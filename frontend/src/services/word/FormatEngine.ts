/**
 * FormatEngine（Agent 格式能力）：整段字体 / 段落格式修改。
 *
 * 复用 PatchEngine.applyEditPlan 骨架：
 *   定位（RangeLocator）→ 乐观锁哈希校验（§15）→ 先建 CC（§24 事务边界）→
 *   trackAll（§25）→ **先读后写**（应用前原值存入事务，拒绝时回写还原）→
 *   软校验（不硬回滚 —— Word 字体名归一化等会造成回读不一致的假失败）→
 *   统计修订数 → finally 恢复 changeTrackingMode。
 *
 * 与文本编辑的关键差异：
 * - Word 可能不把格式修改记录为修订（changeCount 可为 0）—— 事务的拒绝路径
 *   因此改为「回写应用前原值」而非依赖 rejectAll（详见 RevisionService 按 kind 分派）。
 * - 无文本校验（格式改动不改变文字）。
 *
 * 异常回滚：rejectAll（若修订被记录）→ 关修订 → 回写原值 → 删除 CC。
 */
import type { FormatBeforeValues, FormatChanges, FormatPlan } from "@/models/EditPlan";
import type { PatchOutcome, SkippedOp } from "./PatchEngine";
import { CopilotError, toCopilotError } from "@/utils/errors";
import { logger } from "@/utils/logger";
import { textHash } from "@/utils/text";
import { rangeLocator } from "./RangeLocator";
import { WordService } from "./WordService";

export interface FormatOutcome extends PatchOutcome {
  /** 仅含被修改字段的应用前原值（拒绝时回写） */
  beforeValues: FormatBeforeValues;
}

/**
 * alignment 值 → Word 枚举（函数内构造：Word 全局只在宿主内存在，
 * 模块顶层引用会在测试 / Node 环境直接 ReferenceError）。
 */
function alignmentToWord(value: NonNullable<FormatChanges["alignment"]>): Word.Alignment {
  const map: Record<NonNullable<FormatChanges["alignment"]>, Word.Alignment> = {
    left: Word.Alignment.left,
    center: Word.Alignment.centered,
    right: Word.Alignment.right,
    justify: Word.Alignment.justified,
  };
  return map[value];
}

export class FormatEngine {
  /** 应用格式计划。失败（定位 / 哈希 / 写入异常）抛 CopilotError（已自动回滚）。 */
  async applyFormatPlan(plan: FormatPlan): Promise<FormatOutcome> {
    return WordService.run(async (ctx) => {
      // ---- 1. 重新定位 ----
      const located = await rangeLocator.locate(ctx, plan.target);
      if (!located.found) {
        if (located.reason === "ambiguous") {
          throw new CopilotError("RANGE_AMBIGUOUS", "目标内容出现多处匹配");
        }
        throw new CopilotError("RANGE_NOT_FOUND", "目标内容已不存在");
      }

      // ---- 2. 乐观锁校验（段落文本未变） ----
      const targetRange = located.range;
      targetRange.load("text");
      await ctx.sync();
      const currentHash = await textHash(targetRange.text);
      if (currentHash !== plan.target.textHash) {
        throw new CopilotError("DOCUMENT_CHANGED", "目标段落已变化，格式修改未应用。");
      }

      // ---- 3. 保存修订模式 ----
      const previousMode = await WordService.readChangeTrackingMode(ctx);

      let control: Word.ContentControl | null = null;
      let beforeValues: FormatBeforeValues = {};
      const skipped: SkippedOp[] = [];
      let applied = 0;

      try {
        // ---- 4. Content Control 事务边界（先建后写） ----
        control = targetRange.insertContentControl();
        control.tag = plan.contentControlTag;
        control.title = "Word AI";
        try {
          control.appearance = Word.ContentControlAppearance.hidden;
        } catch {
          // 个别宿主不支持 hidden 外观时忽略（默认边界框）
        }

        // ---- 5. 开启 TrackAll ----
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();

        // ---- 6. 读取应用前原值（仅被修改的字段；混合格式读出 null → 不存，回滚跳过） ----
        const content = control.getRange(Word.RangeLocation.content);
        beforeValues = await this.readBeforeValues(ctx, content, plan.changes);

        // ---- 7. 一批写入 ----
        applied = this.applyChanges(content, plan.changes);
        await ctx.sync();

        // ---- 8. 软校验（不一致 → skipped 记录，不回滚：字体名归一化等假失败） ----
        skipped.push(...(await this.softVerify(ctx, content, plan.changes)));

        // ---- 9. 统计实际修订数（Word 可能不记录格式修订 → 0 是合法值） ----
        let changeCount: number | null = null;
        try {
          const changes = control.getTrackedChanges();
          changes.load("text, type");
          await ctx.sync();
          changeCount = changes.items.length;
        } catch (err) {
          logger.warn("统计修订数失败：", err);
        }

        return { contentControlTag: plan.contentControlTag, appliedOps: applied, skipped, changeCount, beforeValues };
      } catch (err) {
        // ---- 异常回滚（尽量还原格式 + 删除 CC） ----
        if (control) {
          await this.rollbackFormat(ctx, control, beforeValues);
        }
        throw toCopilotError(err);
      } finally {
        // ---- 恢复修订模式（§25） ----
        if (previousMode !== null) {
          try {
            ctx.document.changeTrackingMode = previousMode;
            await ctx.sync();
          } catch (err) {
            logger.error("恢复 changeTrackingMode 失败，请手动检查 Word 修订状态：", err);
          }
        }
      }
    });
  }

  // ------------------------------------------------------------------
  // 内部实现
  // ------------------------------------------------------------------

  /**
   * 读取被修改字段的应用前原值（underline / alignment 存原始枚举字符串）。
   * 对齐不在 Range 上（Range 无 paragraphFormat）—— 经 Range.paragraphs
   * 逐段读取（Paragraph.alignment，WordApi 1.1）；多段不一致 → 不存（回滚跳过）。
   */
  private async readBeforeValues(
    ctx: Word.RequestContext,
    content: Word.Range,
    changes: FormatChanges,
  ): Promise<FormatBeforeValues> {
    const fontProps: string[] = [];
    if (changes.bold !== undefined) fontProps.push("bold");
    if (changes.italic !== undefined) fontProps.push("italic");
    if (changes.underline !== undefined) fontProps.push("underline");
    if (changes.strikethrough !== undefined) fontProps.push("strikeThrough");
    if (changes.fontName !== undefined) fontProps.push("name");
    if (changes.fontSize !== undefined) fontProps.push("size");
    if (changes.color !== undefined) fontProps.push("color");
    if (fontProps.length > 0) content.font.load(fontProps.join(","));
    if (changes.alignment !== undefined) content.paragraphs.load("alignment");
    if (fontProps.length === 0 && changes.alignment === undefined) return {};

    await ctx.sync();

    const font = content.font as Word.Font & Record<string, unknown>;
    const before: FormatBeforeValues = {};
    const boolOrSkip = (key: "bold" | "italic" | "strikethrough") => {
      if (typeof font[key === "strikethrough" ? "strikeThrough" : key] === "boolean") {
        before[key] = font[key === "strikethrough" ? "strikeThrough" : key] as boolean;
      }
    };
    if (changes.bold !== undefined) boolOrSkip("bold");
    if (changes.italic !== undefined) boolOrSkip("italic");
    if (changes.strikethrough !== undefined) boolOrSkip("strikethrough");
    if (changes.underline !== undefined) {
      const value = font.underline;
      if (typeof value === "string" && value) before.underline = value;
    }
    if (changes.fontName !== undefined) {
      const value = font.name;
      if (typeof value === "string" && value) before.fontName = value;
    }
    if (changes.fontSize !== undefined) {
      const value = font.size;
      if (typeof value === "number" && value > 0) before.fontSize = value;
    }
    if (changes.color !== undefined) {
      const value = font.color;
      if (typeof value === "string" && value) before.color = value;
    }
    if (changes.alignment !== undefined) {
      const alignments: string[] = content.paragraphs.items.map((p) => String(p.alignment ?? ""));
      const valid = alignments.filter((v) => v !== "" && v !== "Mixed" && v !== "Unknown");
      if (alignments.length > 0 && valid.length === alignments.length && new Set(valid).size === 1) {
        before.alignment = valid[0];
      }
    }
    return before;
  }

  /** 应用格式修改（返回应用字段数）。alignment 逐段写入（items 已由 readBeforeValues 装载）。 */
  private applyChanges(content: Word.Range, changes: FormatChanges): number {
    const font = content.font;
    let applied = 0;
    if (changes.bold !== undefined) {
      font.bold = changes.bold;
      applied++;
    }
    if (changes.italic !== undefined) {
      font.italic = changes.italic;
      applied++;
    }
    if (changes.underline !== undefined) {
      font.underline = changes.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
      applied++;
    }
    if (changes.strikethrough !== undefined) {
      font.strikeThrough = changes.strikethrough;
      applied++;
    }
    if (changes.fontName !== undefined) {
      font.name = changes.fontName;
      applied++;
    }
    if (changes.fontSize !== undefined) {
      font.size = changes.fontSize;
      applied++;
    }
    if (changes.color !== undefined) {
      font.color = changes.color;
      applied++;
    }
    if (changes.alignment !== undefined) {
      const value = alignmentToWord(changes.alignment);
      for (const paragraph of content.paragraphs.items) {
        paragraph.alignment = value;
      }
      applied++;
    }
    return applied;
  }

  /** 软校验：回读已写字段；不一致记 skipped（不回滚）。失败静默（best effort）。 */
  private async softVerify(
    ctx: Word.RequestContext,
    content: Word.Range,
    changes: FormatChanges,
  ): Promise<SkippedOp[]> {
    const skipped: SkippedOp[] = [];
    try {
      const fontProps: string[] = [];
      if (changes.bold !== undefined) fontProps.push("bold");
      if (changes.italic !== undefined) fontProps.push("italic");
      if (changes.underline !== undefined) fontProps.push("underline");
      if (changes.strikethrough !== undefined) fontProps.push("strikeThrough");
      if (changes.fontName !== undefined) fontProps.push("name");
      if (changes.fontSize !== undefined) fontProps.push("size");
      if (changes.color !== undefined) fontProps.push("color");
      if (fontProps.length > 0) content.font.load(fontProps.join(","));
      if (changes.alignment !== undefined) content.paragraphs.load("alignment");
      if (fontProps.length === 0 && changes.alignment === undefined) return skipped;
      await ctx.sync();

      const font = content.font as Word.Font & Record<string, unknown>;
      const mismatch = (name: string, expected: unknown, actual: unknown) => {
        logger.warn(`格式软校验不一致（${name}）：期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`);
        skipped.push({ index: 0, type: `format:${name}`, reason: "value-mismatch" });
      };
      if (changes.bold !== undefined && font.bold !== changes.bold) mismatch("bold", changes.bold, font.bold);
      if (changes.italic !== undefined && font.italic !== changes.italic) mismatch("italic", changes.italic, font.italic);
      if (changes.underline !== undefined) {
        const expected = changes.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        if (font.underline !== expected) mismatch("underline", expected, font.underline);
      }
      if (changes.strikethrough !== undefined && font.strikeThrough !== changes.strikethrough) {
        mismatch("strikethrough", changes.strikethrough, font.strikeThrough);
      }
      if (changes.fontName !== undefined && font.name !== changes.fontName) mismatch("fontName", changes.fontName, font.name);
      if (changes.fontSize !== undefined && font.size !== changes.fontSize) mismatch("fontSize", changes.fontSize, font.size);
      if (changes.color !== undefined && font.color !== changes.color) mismatch("color", changes.color, font.color);
      if (changes.alignment !== undefined) {
        const expected = alignmentToWord(changes.alignment);
        const bad = content.paragraphs.items.some((p) => p.alignment !== expected);
        if (bad) mismatch("alignment", expected, content.paragraphs.items.map((p) => p.alignment).join("|"));
      }
    } catch (err) {
      logger.debug("格式软校验读取失败（忽略）：", err);
    }
    return skipped;
  }

  /**
   * 异常回滚：拒绝已记录的格式修订（若有）→ 关修订回写原值 → 删除 CC。
   * 各步骤独立 try/catch —— 回滚自身失败不掩盖原始异常。
   */
  private async rollbackFormat(
    ctx: Word.RequestContext,
    control: Word.ContentControl,
    before: FormatBeforeValues,
  ): Promise<void> {
    try {
      const changes = control.getTrackedChanges();
      changes.rejectAll(); // 若格式修订被记录，先拒绝还原
    } catch (err) {
      logger.warn("回滚时拒绝格式修订失败（继续回写原值）：", err);
    }
    try {
      // 回写前关修订（否则回写本身会产生新修订；外层 finally 会恢复原模式）
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      const content = control.getRange(Word.RangeLocation.content);
      await this.writeBackBeforeValues(ctx, content, before);
    } catch (err) {
      logger.error("回滚回写格式原值失败，请通过 Word 撤销（Ctrl+Z）还原：", err);
    }
    try {
      control.delete(true);
      await ctx.sync();
    } catch (err) {
      logger.error("回滚删除 Content Control 失败：", err);
    }
  }

  /**
   * 回写应用前原值（RevisionService 拒绝路径也复用）。
   * 注意调用方负责把 changeTrackingMode 切到 off（避免回写产生新修订）。
   * alignment 逐段回写：先装载段落集合再写入（Range 无 paragraphFormat）。
   */
  async writeBackBeforeValues(
    ctx: Word.RequestContext,
    content: Word.Range,
    before: FormatBeforeValues,
  ): Promise<void> {
    const font = content.font;
    if (before.bold !== undefined) font.bold = before.bold;
    if (before.italic !== undefined) font.italic = before.italic;
    if (before.strikethrough !== undefined) font.strikeThrough = before.strikethrough;
    if (before.underline !== undefined) font.underline = before.underline as Word.UnderlineType;
    if (before.fontName !== undefined) font.name = before.fontName;
    if (before.fontSize !== undefined && before.fontSize > 0) font.size = before.fontSize;
    if (before.color !== undefined && before.color) font.color = before.color;
    if (before.alignment !== undefined && before.alignment) {
      content.paragraphs.load("alignment");
      await ctx.sync();
      for (const paragraph of content.paragraphs.items) {
        paragraph.alignment = before.alignment as Word.Alignment;
      }
    }
    await ctx.sync();
  }
}

export const formatEngine = new FormatEngine();

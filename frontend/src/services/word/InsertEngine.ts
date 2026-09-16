/**
 * InsertEngine（Agent 插入能力）：表格 / 公式 / 纯文字段落插入。
 *
 * 插入位置：锚点段落之后，或文档末尾（plan.anchor = null —— 空文档 / 用户
 * 未指明位置时的默认，走 Body.insert*("End")，无需锚点定位与哈希校验）。
 *
 * 表格与文本编辑的结构差异：没有既有文本可包裹 —— 采用**先插入后包裹**：
 *   定位锚点（若有）→ 锚点哈希校验（§15，锚点变 → 插入位置不可信，不应用）→
 *   保存修订模式 → trackAll → 单批插入（Range / Body 的 insertTable / insertOoxml）→
 *   包 Content Control（tag = word_ai_edit:{uuid}，§24 事务边界）→ 一次 sync。
 *
 * 拒绝语义（RevisionService 按 kind 分派）：插入类拒绝 = rejectAll +
 * delete(false)（连壳带内容删除 —— 兜底清除未被 Word 跟踪的插入）。
 *
 * 公式 / 段落 OOXML 统一使用完整 Flat OPC 包；真机验证表明在此格式下可安全
 * 进入 TrackAll + Content Control 管线，保存并重新打开后修订和公式对象均完整。
 *
 * 异常清理：best-effort 删除已插入对象（table.delete() / range.delete()）
 * + REVISION_ERROR；清理失败仅记日志（§15 哈希保证重跑安全）。
 */
import type { FormulaInsertPlan, ParagraphInsertPlan, RangeLocator, TableInsertPlan } from "@/models/EditPlan";
import type { PatchOutcome } from "./PatchEngine";
import { latexToOmml, latexToOoxml } from "./FormulaOoxml";
import { bodyContentToFlatOpc } from "./FlatOpc";
import { CopilotError, toCopilotError } from "@/utils/errors";
import { logger } from "@/utils/logger";
import { stripTrailingMarks, textHash } from "@/utils/text";
import { rangeLocator } from "./RangeLocator";
import { WordService } from "./WordService";

/** XML 文本转义（纯函数，vitest 可测）。 */
export function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Agent 在 paragraph_text 中用 $...$ 或 \(...\) 标注需要成为 Word 对象的行内公式。 */
const INLINE_FORMULA_RE = /\\\((.+?)\\\)|\$([^$\r\n]+)\$/g;

export function containsInlineFormula(lines: readonly string[]): boolean {
  return lines.some((line) => {
    INLINE_FORMULA_RE.lastIndex = 0;
    return INLINE_FORMULA_RE.test(line);
  });
}

function richLineToOoxml(line: string): string {
  let cursor = 0;
  let content = "";
  INLINE_FORMULA_RE.lastIndex = 0;
  for (const match of line.matchAll(INLINE_FORMULA_RE)) {
    const index = match.index ?? 0;
    if (index > cursor) {
      content += `<w:r><w:t xml:space="preserve">${escapeXmlText(line.slice(cursor, index))}</w:t></w:r>`;
    }
    const latex = match[1] ?? match[2] ?? "";
    content += latexToOmml(latex, false);
    cursor = index + match[0].length;
  }
  if (cursor < line.length) {
    content += `<w:r><w:t xml:space="preserve">${escapeXmlText(line.slice(cursor))}</w:t></w:r>`;
  }
  return content || '<w:r><w:t xml:space="preserve"></w:t></w:r>';
}

/**
 * 多段纯文字 → 段落 OOXML（每行一个 <w:p>，可直接传 insertOoxml）。
 * 后端已保证 lines 非空（≤20 行、共 ≤2000 字符）。
 */
export function paragraphsToOoxml(lines: string[]): string {
  const paragraphs = lines
    .map((line) => `<w:p>${richLineToOoxml(line)}</w:p>`)
    .join("");
  return bodyContentToFlatOpc(paragraphs);
}

export class InsertEngine {
  /** 插入表格（需 WordApi 1.3：Range.insertTable / Table 对象模型）。 */
  async applyTableInsertPlan(plan: TableInsertPlan): Promise<PatchOutcome> {
    if (!WordService.isSetSupported("1.3")) {
      throw new CopilotError("UNSUPPORTED_WORD_VERSION", "当前 Word 版本不支持插入表格（需 WordApi 1.3）。");
    }
    return WordService.run(async (ctx) => {
      const anchorRange = await this.locateAnchor(ctx, plan.anchor); // null = 文档末尾

      const previousMode = await WordService.readChangeTrackingMode(ctx);
      let control: Word.ContentControl | null = null;
      let table: Word.Table | null = null;

      try {
        // ---- 先开 TrackAll 再插入（插入本身即修订，接受 / 拒绝语义的基础） ----
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();

        // ---- 单批：插入 → 样式 → 表头 → 包 CC（写后包裹） ----
        const rowCount = plan.values.length;
        const columnCount = plan.values[0]?.length ?? 1;
        table = anchorRange
          ? anchorRange.insertTable(rowCount, columnCount, "After", plan.values)
          : ctx.document.body.insertTable(rowCount, columnCount, "End", plan.values);
        table.styleBuiltIn = "TableGrid";
        table.headerRowCount = plan.header ? 1 : 0;
        const tableRange = table.getRange(Word.RangeLocation.whole);
        control = tableRange.insertContentControl();
        control.tag = plan.contentControlTag;
        control.title = "Word AI";
        try {
          control.appearance = Word.ContentControlAppearance.hidden;
        } catch {
          // 个别宿主不支持 hidden 外观时忽略（默认边界框）
        }
        await ctx.sync();

        return {
          contentControlTag: plan.contentControlTag,
          appliedOps: 1,
          skipped: [],
          changeCount: await this.countChanges(ctx, control),
        };
      } catch (err) {
        // best-effort 清理：CC 已建成 → 拒绝修订 + 删壳带内容；未建成 → 删表
        try {
          if (control) {
            const changes = control.getTrackedChanges();
            changes.rejectAll();
            control.delete(false);
            await ctx.sync();
          } else if (table) {
            table.delete();
            await ctx.sync();
          }
        } catch (cleanupErr) {
          logger.error("表格插入失败清理未完成，请手动删除残留表格：", cleanupErr);
        }
        throw toCopilotError(err);
      } finally {
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

  /** 以 Word 修订事务插入独立公式。 */
  async applyFormulaInsertPlan(plan: FormulaInsertPlan): Promise<PatchOutcome> {
    const ooxml = latexToOoxml(plan.latex, plan.display);
    return this.applyTrackedOoxml(plan.anchor, plan.contentControlTag, ooxml, "公式");
  }

  /** 以 Word 修订事务插入普通或含行内公式的段落。 */
  async applyParagraphInsertPlan(plan: ParagraphInsertPlan): Promise<PatchOutcome> {
    if (plan.lines.length === 0) {
      throw new CopilotError("INVALID_EDIT_RESPONSE", "插入内容为空，未写入文档。");
    }
    const ooxml = paragraphsToOoxml(plan.lines);
    return this.applyTrackedOoxml(plan.anchor, plan.contentControlTag, ooxml, "段落");
  }

  /** Flat OPC 内容统一进入 TrackAll + Content Control 事务管线。 */
  private async applyTrackedOoxml(
    anchor: RangeLocator | null,
    contentControlTag: string,
    ooxml: string,
    label: string,
  ): Promise<PatchOutcome> {
    return WordService.run(async (ctx) => {
      const anchorRange = await this.locateAnchor(ctx, anchor);
      const previousMode = await WordService.readChangeTrackingMode(ctx);
      let control: Word.ContentControl | null = null;
      let inserted: Word.Range | null = null;
      try {
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();
        inserted = anchorRange
          ? anchorRange.insertOoxml(ooxml, "After")
          : ctx.document.body.insertOoxml(ooxml, "End");
        control = inserted.insertContentControl();
        control.tag = contentControlTag;
        control.title = "Word AI";
        try {
          control.appearance = Word.ContentControlAppearance.hidden;
        } catch {
          // 个别宿主不支持 hidden 外观时忽略。
        }
        await ctx.sync();
        return {
          contentControlTag,
          appliedOps: 1,
          skipped: [],
          changeCount: await this.countChanges(ctx, control),
        };
      } catch (err) {
        try {
          if (control) {
            const changes = control.getTrackedChanges();
            changes.rejectAll();
            control.delete(false);
            await ctx.sync();
          } else if (inserted) {
            inserted.delete();
            await ctx.sync();
          }
        } catch (cleanupErr) {
          logger.error(`${label}插入失败清理未完成，请手动检查文档：`, cleanupErr);
        }
        throw toCopilotError(err);
      } finally {
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
   * 定位锚点段落 + 哈希校验（锚点已变 → 插入位置不可信，拒绝应用）。
   * anchor = null → 返回 null（文档末尾模式，跳过定位与校验 —— 空文档可用）。
   */
  private async locateAnchor(ctx: Word.RequestContext, anchor: RangeLocator | null): Promise<Word.Range | null> {
    if (anchor === null) {
      return null;
    }
    const located = await rangeLocator.locate(ctx, anchor);
    if (!located.found) {
      if (located.reason === "ambiguous") {
        throw new CopilotError("RANGE_AMBIGUOUS", "锚点段落出现多处匹配");
      }
      throw new CopilotError("RANGE_NOT_FOUND", "锚点段落已不存在");
    }
    const anchorRange = located.range;
    anchorRange.load("text");
    await ctx.sync();
    // 快照在 DocumentService 中已剥离段落标记；Word 重新读取 Range.text 时，
    // 宿主可能重新附加 \r / \v / 单元格结束符 \x07。结构标记不代表用户修改，
    // 必须先按快照同一规则清洗，否则会误报 DOCUMENT_CHANGED。
    const currentHash = await textHash(stripTrailingMarks(anchorRange.text));
    if (currentHash !== anchor.textHash) {
      throw new CopilotError("DOCUMENT_CHANGED", "锚点段落已变化，插入未应用。");
    }
    return anchorRange;
  }

  /** 统计 CC 内修订数（best effort；插入未被跟踪时为 0 —— 拒绝路径有 delete(false) 兜底）。 */
  private async countChanges(ctx: Word.RequestContext, control: Word.ContentControl): Promise<number | null> {
    try {
      const changes = control.getTrackedChanges();
      changes.load("text, type");
      await ctx.sync();
      return changes.items.length;
    } catch (err) {
      logger.warn("统计修订数失败：", err);
      return null;
    }
  }
}

export const insertEngine = new InsertEngine();

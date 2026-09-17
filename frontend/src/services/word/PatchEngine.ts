/**
 * PatchEngine（对应需求文档 §21 / §24 / §25 / §77）：
 *
 * 在一个 Word.run 批次内完成：
 *   1. RangeLocator 重新定位目标
 *   2. TextHash 乐观锁校验（§15）—— 不一致抛 DOCUMENT_CHANGED，绝不覆盖用户修改
 *   3. 创建 Content Control（tag = word_ai_edit:{uuid}）作为事务边界（§24）
 *   4. 保存 changeTrackingMode → 设置 TrackAll（§25）
 *   5. 从后向前应用 Diff Patch（§21：offset 大 → 小，避免位置漂移）
 *   6. 校验应用结果；不一致自动回滚（rejectAll + 删除 CC）
 *   7. finally 恢复 changeTrackingMode —— 任何异常都不允许 Word 永久停在 TrackAll
 *
 * 事务性：所有 Office.js proxy 对象不跨批次（§47）。
 */
import type { EditOperation, EditPlan } from "@/models/EditPlan";
import { CopilotError } from "@/utils/errors";
import { logger } from "@/utils/logger";
import { normalizeText, reviewedTextEquals, stripTrailingMarks, textHash, toWordText } from "@/utils/text";
import { countOccurrencesBefore, insertAnchorWindow } from "@/utils/match";
import { rangeLocator } from "./RangeLocator";
import { WordService } from "./WordService";

/** insert 锚点长度 */
const ANCHOR_LENGTH = 20;

export interface SkippedOp {
  index: number;
  type: string;
  reason: string;
}

/** 校验失败详情（用于日志与错误信息诊断） */
export interface VerifyFailure {
  final: string;
  expectedFinal: string;
  expectedMarkup: string;
}

export interface PatchOutcome {
  contentControlTag: string;
  appliedOps: number;
  skipped: SkippedOp[];
  /** Word 实际修订条数（best effort） */
  changeCount: number | null;
}

export class PatchEngine {
  /**
   * 应用编辑计划。成功返回事务信息；
   * 失败（目标变化 / 定位失败 / 校验失败）抛 CopilotError。
   */
  async applyEditPlan(plan: EditPlan): Promise<PatchOutcome> {
    return WordService.run(async (ctx) => {
      // ---- 1. 重新定位 ----
      const located = await rangeLocator.locate(ctx, plan.target);
      if (!located.found) {
        if (located.reason === "ambiguous") {
          throw new CopilotError("RANGE_AMBIGUOUS", "目标内容出现多处匹配");
        }
        throw new CopilotError("RANGE_NOT_FOUND", "目标内容已不存在");
      }
      const targetRange = located.range;

      // ---- 2. 乐观锁校验 ----
      targetRange.load("text");
      await ctx.sync();
      const currentText = targetRange.text;
      const currentHash = await textHash(currentText);
      if (currentHash !== plan.target.textHash) {
        throw new CopilotError("DOCUMENT_CHANGED", `目标内容已变化（hash ${currentHash.slice(0, 8)}… ≠ ${plan.target.textHash.slice(0, 8)}…）`);
      }

      // ---- 3. 保存修订模式 ----
      const previousMode = await WordService.readChangeTrackingMode(ctx);

      let control: Word.ContentControl | null = null;
      const skipped: SkippedOp[] = [];
      let applied = 0;

      try {
        // ---- 4. Content Control 事务边界 ----
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

        // ---- 6. 从后向前应用 ----
        const canonicalOriginal = normalizeText(plan.originalText);
        const ops = [...plan.operations].sort((a, b) => b.start - a.start || b.end - a.end);
        for (const op of ops) {
          const ok = await this.applyOperation(ctx, control, op, canonicalOriginal, plan.operations);
          if (ok) {
            applied += 1;
          } else {
            skipped.push({ index: op.start, type: op.type, reason: "anchor-not-found" });
          }
        }
        await ctx.sync();

        // ---- 7. 结果校验（失败自动回滚） ----
        const failure = await this.verifyResult(ctx, control, plan);
        if (failure) {
          await this.rollback(ctx, control);
          control = null;
          logger.warn(
            `修订结果校验不一致（已应用 ${applied}/${plan.operations.length} 处，跳过 ${skipped.length} 处）：\n` +
              `  final=${JSON.stringify(failure.final.slice(0, 120))}\n` +
              `  expectedFinal=${JSON.stringify(failure.expectedFinal.slice(0, 120))}\n` +
              `  expectedMarkup=${JSON.stringify(failure.expectedMarkup.slice(0, 120))}`,
          );
          throw new CopilotError(
            "REVISION_ERROR",
            `修改应用后校验失败，已自动回滚（应用 ${applied}/${plan.operations.length} 处，` +
              `实际文本=${JSON.stringify(failure.final.slice(0, 40))}，` +
              `期望=${JSON.stringify(failure.expectedFinal.slice(0, 40))}）。` +
              `请尝试重新生成，或选择更短的内容。`,
          );
        }

        // ---- 8. 统计实际修订数 ----
        let changeCount: number | null = null;
        try {
          const changes = control.getTrackedChanges();
          changes.load("text, type");
          await ctx.sync();
          changeCount = changes.items.length;
        } catch (err) {
          logger.warn("统计修订数失败：", err);
        }

        return {
          contentControlTag: plan.contentControlTag,
          appliedOps: applied,
          skipped,
          changeCount,
        };
      } finally {
        // ---- 9. 恢复修订模式（成功失败都必须执行 §25） ----
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
   * 应用单个操作。锚定策略：
   * - delete/replace：搜索 oldText，取第 N 个匹配（N = 原文中该位置之前出现的次数）
   * - insert：锚定其后（或其前）的等文本段，插入 Before/After。
   *   后方锚点窗口截止到下一个操作边界 —— 后方操作此刻已应用，
   *   窗口若横跨它们，锚文本已不存在（详见 utils/match.ts insertAnchorWindow）。
   *
   * 返回 false 表示锚点未找到（记入 skipped，由最终校验决定去留）。
   */
  private async applyOperation(
    ctx: Word.RequestContext,
    control: Word.ContentControl,
    op: EditOperation,
    canonicalOriginal: string,
    allOps: readonly EditOperation[],
  ): Promise<boolean> {
    const ctrlRange = control.getRange(Word.RangeLocation.content);
    const oldText = op.oldText ?? "";
    const newText = op.newText ?? "";
    const wordNewText = toWordText(newText);

    // 空锚点（原文为空的极端情况）：直接插到控件开头
    if (!oldText && op.type !== "insert") return false;
    if (op.type === "insert" && canonicalOriginal.length === 0) {
      ctrlRange.insertText(wordNewText, "Start");
      return true;
    }

    if (op.type === "delete" || op.type === "replace") {
      const anchor = this.searchableForm(oldText);
      if (!anchor) return false;
      const occurrence = countOccurrencesBefore(canonicalOriginal, oldText, op.start);
      let outcome = await this.pickMatch(ctx, ctrlRange, anchor, occurrence);
      if (!outcome) {
        // Word search 匹配不到横跨公式（OMML math zone）的文本：Range.text 会把
        // 公式线性化（Diff 坐标系因此成立），但 search 只作用于纯文本 run。
        // 兜底：oldText 前后缀分别锚定，expandTo 覆盖中间不可搜索区域。
        outcome = await this.locateByAnchoredEnds(ctx, ctrlRange, canonicalOriginal, op);
      }
      if (!outcome) return false;
      if (op.type === "delete") {
        outcome.delete();
      } else {
        outcome.insertText(wordNewText, "Replace");
      }
      return true;
    }

    // ---- insert：锚定邻接等文本 ----
    // 优先锚定后方文本（插入到匹配之前）；锚点窗口横跨公式时逐步收缩重试
    const { following, preceding } = insertAnchorWindow(canonicalOriginal, op, allOps, ANCHOR_LENGTH);
    if (following.length >= 1) {
      const outcome = await this.pickInsertAnchor(ctx, ctrlRange, canonicalOriginal, following, op.start, "before");
      if (outcome) {
        outcome.insertText(wordNewText, "Before");
        return true;
      }
    }
    // 锚定前方文本（插入到匹配之后）
    if (preceding.length >= 1) {
      const outcome = await this.pickInsertAnchor(ctx, ctrlRange, canonicalOriginal, preceding, op.start, "after");
      if (outcome) {
        outcome.insertText(wordNewText, "After");
        return true;
      }
    }
    return false;
  }

  /**
   * 插入锚点定位：先按完整窗口搜索（主路径），失败时按减半长度收缩重试。
   *
   * 锚点窗口可能横跨公式（OMML math zone）：线性化的公式文本无法被 search
   * 命中。收缩策略 —— “后方锚点”取窗口前缀（仍从插入点开始）、“前方锚点”
   * 取窗口后缀（仍紧贴插入点），出现序号按缩短后的锚点重算，插入位置不变。
   */
  private async pickInsertAnchor(
    ctx: Word.RequestContext,
    scope: Word.Range,
    original: string,
    window: string,
    insertAt: number,
    location: "before" | "after",
  ): Promise<Word.Range | null> {
    let len = window.length;
    while (len >= 1) {
      const probe = location === "before" ? window.slice(0, len) : window.slice(window.length - len);
      const occurrence = location === "before"
        ? countOccurrencesBefore(original, probe, insertAt)
        : Math.max(0, countOccurrencesBefore(original, probe, insertAt) - 1);
      const outcome = await this.pickMatch(ctx, scope, toWordText(probe), occurrence);
      if (outcome) return outcome;
      len = Math.floor(len / 2);
    }
    return null;
  }

  /**
   * delete/replace 的兜底定位（oldText 横跨公式等不可搜索内容时）：
   * 把 oldText 拆成“最长可搜索前缀 + 最长可搜索后缀”，分别按出现序号
   * 定位后 expandTo 合并 —— 合并范围覆盖中间的不可搜索区域（公式）。
   * 合并范围的 .text 必须与 oldText 的线性文本完全一致，否则视为定位
   * 失败（宁可不改也不错改）。
   */
  private async locateByAnchoredEnds(
    ctx: Word.RequestContext,
    scope: Word.Range,
    original: string,
    op: EditOperation,
  ): Promise<Word.Range | null> {
    const oldText = op.oldText ?? "";
    if (oldText.length < 4) return null; // 过短的文本没有可靠的拆分空间

    const prefixMatch = await this.pickEndAnchor(ctx, scope, original, oldText, op.start, "prefix");
    if (!prefixMatch) return null;
    const suffixMatch = await this.pickEndAnchor(ctx, scope, original, oldText, op.end, "suffix");
    if (!suffixMatch) return null;

    const combined = prefixMatch.expandTo(suffixMatch);
    combined.load("text");
    await ctx.sync();
    const combinedText = stripTrailingMarks(normalizeText(combined.text));
    if (combinedText !== stripTrailingMarks(normalizeText(oldText))) return null;
    return combined;
  }

  /** 锚定 oldText 一端：从半长开始按减半长度收缩，直到该端锚点可搜索。 */
  private async pickEndAnchor(
    ctx: Word.RequestContext,
    scope: Word.Range,
    original: string,
    oldText: string,
    boundary: number,
    end: "prefix" | "suffix",
  ): Promise<Word.Range | null> {
    let len = Math.floor(oldText.length / 2);
    while (len >= 1) {
      const probe = end === "prefix" ? oldText.slice(0, len) : oldText.slice(oldText.length - len);
      const occurrence = end === "prefix"
        ? countOccurrencesBefore(original, probe, boundary)
        : Math.max(0, countOccurrencesBefore(original, probe, boundary) - 1);
      const outcome = await this.pickMatch(ctx, scope, toWordText(probe), occurrence);
      if (outcome) return outcome;
      len = Math.floor(len / 2);
    }
    return null;
  }

  /**
   * 在控件范围内搜索锚文本并选中第 occurrence 个匹配。
   */
  private async pickMatch(
    ctx: Word.RequestContext,
    scope: Word.Range,
    anchor: string,
    occurrence: number,
  ): Promise<Word.Range | null> {
    try {
      const matches = scope.search(anchor, { matchCase: true });
      matches.load("text");
      await ctx.sync();
      // occurrence 是按原文计算出的精确序号。候选不足时绝不能退回最后一个
      // 匹配，否则会把相同短语的另一处内容误改，随后只能靠最终校验回滚。
      if (occurrence < 0 || occurrence >= matches.items.length) return null;
      return matches.items[occurrence];
    } catch (err) {
      logger.debug("锚点搜索失败：", err);
      return null;
    }
  }

  /**
   * 校验最终文本。Word 在开启修订时 range.text 可能返回
   * “最终视图”（= newText）、“含修订标记视图”（= 删除保留 + 插入加入，
   * 且 replace 的旧/新先后顺序因版本而异）—— 以上均视为成功；
   * 都不是则返回不一致详情（用于诊断与错误信息）。
   */
  private async verifyResult(
    ctx: Word.RequestContext,
    control: Word.ContentControl,
    plan: EditPlan,
  ): Promise<VerifyFailure | null> {
    const expectedFinal = normalizeText(plan.newText).replace(/[\n\r\v]+$/, "");
    const range = control.getRange(Word.RangeLocation.content);

    // WordApi 1.4+ 可直接取得“接受全部修订后”的当前版本文本。这比 range.text
    // 稳定：后者受 Word 的修订显示模式影响，可能同时包含删除和插入内容，且多个
    // replace 的旧/新顺序不一定一致。本应用的修订能力门槛是 WordApi 1.6。
    try {
      const reviewed = range.getReviewedText(Word.ChangeTrackingVersion.current);
      await ctx.sync();
      const finalText = normalizeText(reviewed.value).replace(/[\n\r\v]+$/, "");
      if (reviewedTextEquals(finalText, expectedFinal)) return null;
      return {
        final: finalText,
        expectedFinal,
        expectedMarkup: "",
      };
    } catch (err) {
      // 防御性兼容异常宿主：退回旧版的 range.text 多视图判断。
      logger.debug("getReviewedText 校验失败，退回 range.text：", err);
    }

    range.load("text");
    await ctx.sync();
    const finalText = normalizeText(range.text).replace(/[\n\r\v]+$/, "");
    if (finalText === expectedFinal) return null;

    const expectedMarkup = normalizeText(this.buildMarkupExpectation(plan, "old-first")).replace(
      /[\n\r\v]+$/,
      "",
    );
    if (finalText === expectedMarkup) return null;
    const expectedMarkupAlt = normalizeText(this.buildMarkupExpectation(plan, "new-first")).replace(
      /[\n\r\v]+$/,
      "",
    );
    if (finalText === expectedMarkupAlt) return null;

    return { final: finalText, expectedFinal, expectedMarkup };
  }

  /**
   * 构造“含修订标记视图”的期望文本：
   * 原文中依次叠加 —— replace: old+new（或 new+old）；delete: old；insert: new。
   */
  private buildMarkupExpectation(plan: EditPlan, replaceOrder: "old-first" | "new-first"): string {
    const original = normalizeText(plan.originalText);
    const ops = [...plan.operations].sort((a, b) => a.start - b.start);
    let out = "";
    let pos = 0;
    for (const op of ops) {
      out += original.slice(pos, op.start);
      if (op.type === "replace") {
        out += replaceOrder === "old-first"
          ? (op.oldText ?? "") + (op.newText ?? "")
          : (op.newText ?? "") + (op.oldText ?? "");
      } else if (op.type === "delete") {
        out += op.oldText ?? "";
      } else {
        out += op.newText ?? "";
      }
      pos = Math.max(pos, op.end);
    }
    out += original.slice(pos);
    return out;
  }

  /**
   * 回滚：拒绝本次全部修订并移除临时 Content Control（保留正文 → 恢复原状）。
   */
  private async rollback(ctx: Word.RequestContext, control: Word.ContentControl): Promise<void> {
    try {
      const changes = control.getTrackedChanges();
      changes.rejectAll();
      control.delete(true); // 只删除控件外壳，保留（已还原的）正文
      await ctx.sync();
      logger.info("已回滚本次 AI 修改");
    } catch (err) {
      logger.error("回滚失败，请通过 Word 审阅手动拒绝本次修订：", err);
    }
  }

  /** 搜索用文本形式：\n → \r；若包含换行则先尝试可搜索形式 */
  private searchableForm(text: string): string {
    if (!text) return "";
    const wordForm = toWordText(text);
    return wordForm;
  }
}

export const patchEngine = new PatchEngine();

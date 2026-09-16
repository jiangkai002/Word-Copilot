/**
 * RangeLocator（对应需求文档 §14 / §48）：
 * 三级定位 —— Paragraph.uniqueLocalId → 原文精确匹配 → prefix + 原文 + suffix。
 *
 * 重要约束（§47）：
 * - 本类方法必须在 WordService.run 的同一批次内调用；
 * - 返回的 Word.Range proxy 只能在当前批次内使用，不允许存入 Store。
 */
import type { RangeLocator as LocatorModel } from "@/models/EditPlan";
import { normalizeText, stripTrailingMarks } from "@/utils/text";
import { logger } from "@/utils/logger";

export type LocateFailureReason = "not_found" | "ambiguous";

export type LocateResult =
  | { found: true; range: Word.Range }
  | { found: false; reason: LocateFailureReason };

/** 搜索锚点分块长度（规避 Word search 长度限制的兜底策略） */
const ANCHOR_CHUNK = 100;
/** 单段搜索的安全长度上限 */
const SAFE_SEARCH_LENGTH = 200;

interface ParagraphSnapshot {
  id: string | undefined;
  text: string;
  raw: Word.Paragraph;
}

export class RangeLocator {
  /** 入口：三级定位 */
  async locate(ctx: Word.RequestContext, locator: LocatorModel): Promise<LocateResult> {
    const stripped = stripTrailingMarks(locator.originalText);
    if (!stripped) return { found: false, reason: "not_found" };

    const paragraphs = await this.loadParagraphs(ctx);

    // ---- 第一优先级：paragraphId ----
    if (locator.paragraphId) {
      const idx = paragraphs.findIndex((p) => p.id === locator.paragraphId);
      if (idx >= 0) {
        const result = await this.locateInParagraphs(ctx, paragraphs, idx, locator, stripped);
        if (result.found || result.reason === "ambiguous") {
          // ambiguous 时不再降级（§48：多匹配禁止自动修改）
          return result;
        }
        logger.debug("RangeLocator: 段落内定位失败，降级到全文搜索");
      } else {
        logger.debug("RangeLocator: 未找到 paragraphId，降级到全文搜索");
      }
    }

    // ---- 第二优先级：原文精确匹配 ----
    const exact = await this.searchExact(ctx, stripped);
    if (exact) {
      if (exact.matches === 1) {
        const range = this.pickFirst(ctx, exact.results);
        return this.verifyPicked(ctx, range, stripped);
      }
      if (exact.matches > 1) {
        // 多匹配 → 尝试第三级消歧
        if (locator.prefix || locator.suffix) {
          const disambiguated = await this.locateByPrefixSuffix(ctx, locator, stripped);
          if (disambiguated) return disambiguated;
        }
        return { found: false, reason: "ambiguous" };
      }
    }

    // ---- 第三优先级：prefix + 原文 + suffix ----
    if (locator.prefix || locator.suffix) {
      const combined = await this.locateByPrefixSuffix(ctx, locator, stripped);
      if (combined) return combined;
    }

    // ---- 兜底：按段落文本整体扫描 ----
    return this.locateByParagraphScan(ctx, paragraphs, locator, stripped);
  }

  /**
   * 定位成功后校验文本哈希（§15 乐观锁）。
   * 返回 null 表示通过；返回当前文本用于上层报错。
   */
  async readRangeText(ctx: Word.RequestContext, range: Word.Range): Promise<string> {
    range.load("text");
    await ctx.sync();
    return range.text;
  }

  // ------------------------------------------------------------------
  // 内部实现
  // ------------------------------------------------------------------

  private async loadParagraphs(ctx: Word.RequestContext): Promise<ParagraphSnapshot[]> {
    const paras = ctx.document.body.paragraphs;
    paras.load(["uniqueLocalId", "text"]);
    await ctx.sync();
    return paras.items.map((p) => ({
      id: p.uniqueLocalId,
      text: p.text ?? "",
      raw: p,
    }));
  }

  /** 在已定位段落（idx 起）内寻找目标范围 */
  private async locateInParagraphs(
    ctx: Word.RequestContext,
    paragraphs: ParagraphSnapshot[],
    idx: number,
    locator: LocatorModel,
    stripped: string,
  ): Promise<LocateResult> {
    const target = normalizeText(stripped);
    const paraTextN = normalizeText(paragraphs[idx].text);

    if (!target.includes("\n")) {
      // 单段内目标：可能等于整段，也可能是段内选区
      if (stripTrailingMarks(paraTextN) === stripTrailingMarks(target)) {
        // 整段即目标
        return this.locateWholeParagraph(ctx, paragraphs[idx].raw, target);
      }
      // 段内搜索
      const paraRange = paragraphs[idx].raw.getRange(Word.RangeLocation.whole);
      const searchOutcome = await this.trySearch(ctx, paraRange, target);
      if (searchOutcome) {
        if (searchOutcome.matches === 1) {
          return this.verifyPicked(ctx, this.pickFirst(ctx, searchOutcome.results), target);
        }
        if (searchOutcome.matches > 1) {
          // 尝试用 prefix/suffix 消歧：段落文本应等于 prefix+target+suffix
          if (this.prefixSuffixConsistent(locator, target, paraTextN)) {
            const combined = await this.locateByChunks(
              ctx,
              paraRange,
              target,
            );
            if (combined) return combined;
          }
          return { found: false, reason: "ambiguous" };
        }
      }
      // 搜索不可用（长文本等）→ 分块兜底
      if (this.prefixSuffixConsistent(locator, target, paraTextN)) {
        const paraRange2 = paragraphs[idx].raw.getRange(Word.RangeLocation.whole);
        const combined = await this.locateByChunks(ctx, paraRange2, target);
        if (combined) return combined;
      }
      return { found: false, reason: "not_found" };
    }

    // 跨段目标：从 idx 起拼接段落，直到匹配
    let joined = stripTrailingMarks(normalizeText(paragraphs[idx].text));
    let j = idx;
    const strippedTarget = stripTrailingMarks(target);
    while (joined.length < strippedTarget.length && j + 1 < paragraphs.length) {
      j += 1;
      joined += "\n" + stripTrailingMarks(normalizeText(paragraphs[j].text));
    }
    if (joined === strippedTarget) {
      const first = paragraphs[idx].raw.getRange(Word.RangeLocation.content);
      const last = paragraphs[j].raw.getRange(Word.RangeLocation.content);
      const combined = first.expandTo(last);
      combined.load("text");
      await ctx.sync();
      if (normalizeText(combined.text).startsWith(stripTrailingMarks(target))) {
        return { found: true, range: combined };
      }
    }
    return { found: false, reason: "not_found" };
  }

  private async locateWholeParagraph(
    ctx: Word.RequestContext,
    paragraph: Word.Paragraph,
    target: string,
  ): Promise<LocateResult> {
    const whole = paragraph.getRange(Word.RangeLocation.whole);
    // 优先搜索精确范围（排除段落标记）
    const searchOutcome = await this.trySearch(ctx, whole, stripTrailingMarks(target));
    if (searchOutcome && searchOutcome.matches === 1) {
      return this.verifyPicked(ctx, this.pickFirst(ctx, searchOutcome.results), target);
    }
    // 兜底：Content 范围（不含段落标记）并校验
    const content = paragraph.getRange(Word.RangeLocation.content);
    content.load("text");
    await ctx.sync();
    if (normalizeText(content.text).trimEnd() === target.trimEnd()) {
      return { found: true, range: content };
    }
    return { found: false, reason: "not_found" };
  }

  /** 全文精确搜索（返回 null 表示搜索不可用/失败，0 匹配返回 {matches: 0}） */
  private async searchExact(
    ctx: Word.RequestContext,
    stripped: string,
  ): Promise<{ results: Word.RangeCollection; matches: number } | null> {
    if (stripped.length > SAFE_SEARCH_LENGTH) {
      return null; // 长文本交给段落/分块策略
    }
    const outcome = await this.trySearch(ctx, ctx.document.body, stripped);
    return outcome ?? null;
  }

  private async locateByPrefixSuffix(
    ctx: Word.RequestContext,
    locator: LocatorModel,
    stripped: string,
  ): Promise<LocateResult | null> {
    const prefix = stripTrailingMarks(normalizeText(locator.prefix ?? ""));
    const suffix = normalizeText(locator.suffix ?? "").trim();
    if (!prefix && !suffix) return null;
    const combinedText = (prefix + normalizeText(stripped) + suffix).replace(/\n/g, "");
    if (!combinedText || combinedText.length > SAFE_SEARCH_LENGTH) {
      return null;
    }
    try {
      const results = ctx.document.body.search(combinedText, { matchCase: true });
      results.load("text");
      await ctx.sync();
      if (results.items.length === 1) {
        const outer = this.pickFirst(ctx, results);
        // 在外层范围内再定位目标本身
        const inner = await this.trySearch(ctx, outer, stripTrailingMarks(stripped));
        if (inner && inner.matches >= 1) {
          return this.verifyPicked(ctx, this.pickFirst(ctx, inner.results), normalizeText(stripped));
        }
        const chunked = await this.locateByChunks(ctx, outer, normalizeText(stripped));
        if (chunked) return chunked;
      }
      return null;
    } catch {
      return null;
    }
  }

  /** 长文本兜底：首尾分块搜索 + expandTo 合并 + 文本校验 */
  private async locateByChunks(
    ctx: Word.RequestContext,
    scope: Word.Range,
    target: string,
  ): Promise<LocateResult | null> {
    const strippedTarget = stripTrailingMarks(target);
    if (strippedTarget.length <= SAFE_SEARCH_LENGTH) {
      const outcome = await this.trySearch(ctx, scope, strippedTarget);
      if (outcome && outcome.matches === 1) {
        return this.verifyPicked(ctx, this.pickFirst(ctx, outcome.results), target);
      }
      if (outcome && outcome.matches > 1) {
        return { found: false, reason: "ambiguous" };
      }
      return null;
    }
    const firstChunk = strippedTarget.slice(0, ANCHOR_CHUNK);
    const lastChunk = strippedTarget.slice(strippedTarget.length - ANCHOR_CHUNK);
    try {
      const firsts = scope.search(firstChunk, { matchCase: true });
      firsts.load("text");
      await ctx.sync();
      const lasts = scope.search(lastChunk, { matchCase: true });
      lasts.load("text");
      await ctx.sync();
      for (const first of firsts.items) {
        for (const last of lasts.items) {
          const combined = first.expandTo(last);
          combined.load("text");
          await ctx.sync();
          if (normalizeText(combined.text).trimEnd() === strippedTarget.trimEnd()) {
            return { found: true, range: combined };
          }
        }
      }
    } catch (err) {
      logger.debug("locateByChunks 失败：", err);
    }
    return null;
  }

  /** 段落整体扫描兜底（适用于段落模式目标 / 搜索不可用的场景） */
  private async locateByParagraphScan(
    ctx: Word.RequestContext,
    paragraphs: ParagraphSnapshot[],
    _locator: LocatorModel,
    stripped: string,
  ): Promise<LocateResult> {
    const target = stripTrailingMarks(normalizeText(stripped));
    if (!target.includes("\n")) {
      const hits = paragraphs.filter((p) => stripTrailingMarks(normalizeText(p.text)) === target);
      if (hits.length === 1) {
        return this.locateWholeParagraph(ctx, hits[0].raw, target);
      }
      if (hits.length > 1) return { found: false, reason: "ambiguous" };
    } else {
      // 跨段拼接扫描
      const lines = target.split("\n");
      for (let i = 0; i < paragraphs.length; i++) {
        let ok = true;
        for (let k = 0; k < lines.length; k++) {
          if (i + k >= paragraphs.length) { ok = false; break; }
          if (stripTrailingMarks(normalizeText(paragraphs[i + k].text)) !== lines[k]) { ok = false; break; }
        }
        if (ok) {
          const first = paragraphs[i].raw.getRange(Word.RangeLocation.content);
          const last = paragraphs[i + lines.length - 1].raw.getRange(Word.RangeLocation.content);
          const combined = first.expandTo(last);
          combined.load("text");
          await ctx.sync();
          return { found: true, range: combined };
        }
      }
    }
    return { found: false, reason: "not_found" };
  }

  /** 检查 prefix + target + suffix 与段落文本一致（段内消歧的前提） */
  private prefixSuffixConsistent(locator: LocatorModel, target: string, paraTextN: string): boolean {
    if (!locator.prefix && !locator.suffix) return true;
    const rebuilt = normalizeText(locator.prefix ?? "") + target + normalizeText(locator.suffix ?? "");
    return stripTrailingMarks(rebuilt) === stripTrailingMarks(paraTextN);
  }

  /** 安全执行搜索；返回 null 表示搜索失败（如超长） */
  private async trySearch(
    ctx: Word.RequestContext,
    scope: Word.Range | Word.Body,
    text: string,
  ): Promise<{ results: Word.RangeCollection; matches: number } | null> {
    if (!text) return null;
    try {
      const results = scope.search(text, { matchCase: true });
      results.load("text");
      await ctx.sync();
      return { results, matches: results.items.length };
    } catch (err) {
      logger.debug("search 调用失败（可能超长）：", err);
      return null;
    }
  }

  private pickFirst(ctx: Word.RequestContext, results: Word.RangeCollection): Word.Range {
    void ctx;
    return results.items[0];
  }

  private async verifyPicked(
    ctx: Word.RequestContext,
    range: Word.Range,
    expected: string,
  ): Promise<LocateResult> {
    range.load("text");
    await ctx.sync();
    const actual = normalizeText(range.text).trimEnd();
    if (actual === stripTrailingMarks(normalizeText(expected)).trimEnd()) {
      return { found: true, range };
    }
    return { found: false, reason: "not_found" };
  }
}

export const rangeLocator = new RangeLocator();

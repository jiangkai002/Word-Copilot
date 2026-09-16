/**
 * SelectionService（对应需求文档 §9 / §46 / §56）：
 * 读取选区、光标段落及前后文，构造 RangeLocator 与编辑目标。
 *
 * §56：发送 AI 请求时记录目标 Selection 的定位信息，
 * AI 返回后即使光标已移动，也通过 RangeLocator 找回原目标。
 */
import type { RangeLocator } from "@/models/EditPlan";
import type { ParagraphContext } from "@/models/DocumentContext";
import { CopilotError } from "@/utils/errors";
import { normalizeText, stripTrailingMarks, textHash } from "@/utils/text";
import { WordService } from "./WordService";

export interface CapturedTarget {
  /** 目标类型 */
  kind: "selection" | "paragraph";
  /** 目标文本（保持 Word 原始形态，结尾段落标记已剥离） */
  text: string;
  textHash: string;
  locator: RangeLocator;
  /** 同段前缀 / 后缀（段落模式为前后段文本） */
  prefix: string;
  suffix: string;
}

export interface SelectionSnapshot {
  hasSelection: boolean;
  text: string;
  paragraphId?: string;
  paragraphStyle?: string;
  /** 选区在段落内时的同段前后文 */
  prefix: string;
  suffix: string;
}

export class SelectionService {
  /** 当前选区文本（空选区返回空串） */
  async getSelectionText(): Promise<string> {
    return WordService.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      await ctx.sync();
      return sel.text ?? "";
    });
  }

  /**
   * 选区快照（用于 ContextSelector 展示；轻量，不计算哈希）。
   */
  async getSelectionSnapshot(): Promise<SelectionSnapshot> {
    return WordService.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      const paras = sel.paragraphs;
      paras.load(["uniqueLocalId", "text", "style"]);
      await ctx.sync();

      const text = sel.text ?? "";
      const hasSelection = text.trim().length > 0;
      const firstPara = paras.items[0];
      const paragraphText = firstPara?.text ?? "";
      const paragraphId = firstPara?.uniqueLocalId;

      let prefix = "";
      let suffix = "";
      if (hasSelection) {
        const strippedSel = stripTrailingMarks(text);
        const idx = paragraphText.indexOf(strippedSel);
        if (idx >= 0) {
          prefix = paragraphText.slice(0, idx);
          suffix = paragraphText.slice(idx + strippedSel.length);
        }
      }
      return {
        hasSelection,
        text,
        paragraphId,
        paragraphStyle: firstPara?.style,
        prefix,
        suffix,
      };
    });
  }

  /**
   * 光标所在段落及其前后段落（§10）。
   */
  async getParagraphContext(): Promise<{
    current: ParagraphContext;
    previous: ParagraphContext | null;
    next: ParagraphContext | null;
  }> {
    return WordService.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      const paras = sel.paragraphs;
      paras.load(["uniqueLocalId", "text", "style"]);
      await ctx.sync();

      const first = paras.items[0];
      if (!first) {
        throw new CopilotError("UNSUPPORTED_CONTENT", "未找到光标所在段落");
      }
      const currentText = first.text ?? "";
      const current = {
        uniqueLocalId: first.uniqueLocalId,
        text: currentText,
        style: first.style,
        textHash: "",
      };

      const prev = first.getPreviousOrNullObject();
      const next = first.getNextOrNullObject();
      prev.load("text");
      next.load("text");
      await ctx.sync();

      const previous = prev.isNullObject ? null : { text: prev.text ?? "", textHash: "" };
      const nextPara = next.isNullObject ? null : { text: next.text ?? "", textHash: "" };

      current.textHash = await textHash(currentText);
      if (previous) previous.textHash = await textHash(previous.text);
      if (nextPara) nextPara.textHash = await textHash(nextPara.text);

      return { current, previous, next: nextPara };
    });
  }

  /**
   * Agent 模式焦点（§51）：光标 / 选区所在段落（供模型解析「这段」的指代）。
   * 尽力而为 —— 无段落（空文档）或段落为空时返回 null，不抛错。
   */
  async captureAgentFocus(): Promise<{
    paragraphId: string | null;
    paragraphText: string;
    selectedText: string | null;
  } | null> {
    return WordService.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      const paras = sel.paragraphs;
      paras.load(["uniqueLocalId", "text"]);
      await ctx.sync();

      const first = paras.items[0];
      if (!first) return null;
      const paragraphText = stripTrailingMarks(first.text ?? "");
      if (!paragraphText.trim()) return null;
      const rawSelection = sel.text ?? "";
      const selectedText = rawSelection.trim() ? stripTrailingMarks(rawSelection) : null;
      return {
        paragraphId: first.uniqueLocalId ?? null,
        paragraphText,
        selectedText,
      };
    });
  }

  /**
   * 捕获编辑目标（§77 第一步）：
   * - 有选区 → selection 目标（含同段 prefix/suffix）
   * - 无选区 → 光标所在段落（含前后段落）
   *
   * 同时执行复杂内容检测（§38：图片 / 嵌入 ContentControl → 拒绝）。
   */
  async captureEditTarget(maxLength = 5000): Promise<CapturedTarget> {
    const snapshot = await WordService.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      const selParas = sel.paragraphs;
      selParas.load(["uniqueLocalId", "text", "style"]);

      // §38 复杂内容检测：图片 / 现有内容控件
      // （Word 集合无 getCount：加载标量属性后用 items.length）
      const pictures = sel.inlinePictures;
      pictures.load("$all");
      const controls = sel.contentControls;
      controls.load("id");
      await ctx.sync();

      if (pictures.items.length > 0 || controls.items.length > 0) {
        throw new CopilotError("UNSUPPORTED_CONTENT");
      }

      const text = sel.text ?? "";
      const hasSelection = text.trim().length > 0;

      if (hasSelection) {
        const firstPara = selParas.items[0];
        const paragraphText = firstPara?.text ?? "";
        const stripped = stripTrailingMarks(text);
        const idx = paragraphText.indexOf(stripped);
        return {
          kind: "selection" as const,
          text: stripped,
          paragraphId: firstPara?.uniqueLocalId,
          prefix: idx > 0 ? paragraphText.slice(0, idx) : "",
          suffix: idx >= 0 ? paragraphText.slice(idx + stripped.length) : "",
        };
      }

      // 无选区 → 光标段落
      const para = selParas.items[0];
      if (!para) {
        throw new CopilotError("UNSUPPORTED_CONTENT", "未找到可修改的段落，请先选中或定位到正文文本。");
      }
      const paraText = stripTrailingMarks(para.text ?? "");
      if (!paraText.trim()) {
        throw new CopilotError("UNSUPPORTED_CONTENT", "当前段落为空，请选中要修改的文字。");
      }
      const prev = para.getPreviousOrNullObject();
      const next = para.getNextOrNullObject();
      prev.load("text");
      next.load("text");
      await ctx.sync();
      return {
        kind: "paragraph" as const,
        text: paraText,
        paragraphId: para.uniqueLocalId,
        prefix: prev.isNullObject ? "" : stripTrailingMarks(prev.text ?? ""),
        suffix: next.isNullObject ? "" : stripTrailingMarks(next.text ?? ""),
      };
    });

    const normalized = normalizeText(snapshot.text);
    if (normalized.length > maxLength) {
      throw new CopilotError("UNSUPPORTED_CONTENT", `目标内容过长（${normalized.length} 字符 > ${maxLength}），请选择更短的内容。`);
    }
    if (!normalized.trim()) {
      throw new CopilotError("UNSUPPORTED_CONTENT", "请先选中要修改的文字。");
    }

    const hash = await textHash(snapshot.text);
    return {
      kind: snapshot.kind,
      text: snapshot.text,
      textHash: hash,
      prefix: snapshot.prefix,
      suffix: snapshot.suffix,
      locator: {
        paragraphId: snapshot.paragraphId,
        originalText: snapshot.text,
        textHash: hash,
        prefix: snapshot.kind === "selection" ? snapshot.prefix : undefined,
        suffix: snapshot.kind === "selection" ? snapshot.suffix : undefined,
      },
    };
  }
}

export const selectionService = new SelectionService();

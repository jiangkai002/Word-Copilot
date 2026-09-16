/**
 * DocumentService（对应需求文档 §11 / §12 / §50）：
 * 文档大纲、当前章节、全文（带 50k 上限）、超长文档的降级摘要。
 */
import type { DocumentOutlineItem, SectionContext } from "@/models/DocumentContext";
import { logger } from "@/utils/logger";
import { normalizeText, stripTrailingMarks } from "@/utils/text";
import { WordService } from "./WordService";

/** §12：Document 模式全文上限 */
export const DOCUMENT_TEXT_LIMIT = 50_000;
/** §12：超长文档降级时相关段落正文上限 */
export const FALLBACK_RELEVANT_LIMIT = 6_000;
/** 章节内容上限 */
export const SECTION_LIMIT = 30_000;

interface ParagraphScanItem {
  id?: string;
  text: string;
  styleBuiltIn: string;
  style: string;
  level: number | null; // Heading 层级，非标题为 null
  inTable: boolean; // 是否表格单元格内段落（Body.paragraphs 自 1.3 起含表内段）
}

/** Agent 快照段落（§51）：id = uniqueLocalId（缺失时回退为序号） */
export interface SnapshotParagraph {
  id: string;
  /** 已剥离段落标记的段落文本 */
  text: string;
  style: string | null;
  /** Heading 层级（非标题为 null） */
  level: number | null;
  /**
   * 段落位于表格单元格内（Body.paragraphs 自 WordApi 1.3 起包含表内段落）。
   * 表内段落不能作为插入类提案的锚点（会把内容插进单元格）——
   * 后端校验据此拒绝，提示模型改用文档末尾。
   */
  inTable?: boolean;
}

/** Agent 文档快照（§51 / §60）：发送前 UI 必须披露读取范围 */
export interface DocumentSnapshot {
  outline: DocumentOutlineItem[];
  paragraphs: SnapshotParagraph[];
  truncated: boolean;
}

export class DocumentService {
  /**
   * 文档快照（§51 Agent 输入）：全部段落 + 大纲；
   * 正文超 DOCUMENT_TEXT_LIMIT 时截断并标记（大纲同步裁剪到截断点，
   * 避免模型看到截断范围之外的目录结构）。
   * 段落文本与 captureEditTarget 段落模式语义一致（stripTrailingMarks）。
   */
  async getDocumentSnapshot(): Promise<DocumentSnapshot> {
    const scan = await this.scanParagraphs();
    let cutIndex = scan.length;
    let length = 0;
    let truncated = false;
    for (let i = 0; i < scan.length; i++) {
      const text = stripTrailingMarks(scan[i].text);
      if (length + text.length > DOCUMENT_TEXT_LIMIT) {
        cutIndex = i;
        truncated = true;
        break;
      }
      length += text.length;
    }
    const within = scan.slice(0, cutIndex);
    return {
      outline: within
        .filter((p) => p.level !== null)
        .map((p) => ({ level: p.level as number, title: stripTrailingMarks(p.text).trim() })),
      paragraphs: within.map((p, i) => ({
        id: p.id ?? `para-${i}`,
        text: stripTrailingMarks(p.text),
        style: p.style || null,
        level: p.level,
        inTable: p.inTable,
      })),
      truncated,
    };
  }

  /**
   * 文档大纲（§50）：基于 Heading 样式层级。
   */
  async getOutline(): Promise<DocumentOutlineItem[]> {
    const scan = await this.scanParagraphs();
    return scan
      .filter((p) => p.level !== null)
      .map((p) => ({ level: p.level as number, title: stripTrailingMarks(p.text).trim() }));
  }

  /**
   * 全文文本（§12）：≤ 50,000 字符直接返回；超长返回 null（由上层降级）。
   */
  async getDocumentText(limit = DOCUMENT_TEXT_LIMIT): Promise<string | null> {
    return WordService.run(async (ctx) => {
      const body = ctx.document.body;
      body.load("text");
      await ctx.sync();
      const text = body.text ?? "";
      if (text.length > limit) return null;
      return text;
    });
  }

  /**
   * 当前章节（§11）：从光标段落向上找最近标题，
   * 内容直到下一个同级或更高级标题之前。
   */
  async getSectionContext(currentParagraphId?: string): Promise<SectionContext | null> {
    const scan = await this.scanParagraphs();
    if (scan.length === 0) return null;

    let cursorIndex = -1;
    if (currentParagraphId) {
      cursorIndex = scan.findIndex((p) => p.id === currentParagraphId);
    }
    if (cursorIndex < 0) {
      // 无 id 时退化为第一个段落
      cursorIndex = 0;
    }

    // 向上找最近的标题
    let headingIndex = -1;
    for (let i = cursorIndex; i >= 0; i--) {
      if (scan[i].level !== null) {
        headingIndex = i;
        break;
      }
    }
    if (headingIndex < 0) return null;

    const level = scan[headingIndex].level as number;
    const title = stripTrailingMarks(scan[headingIndex].text).trim();

    // 收集内容直到下一个同级/更高级标题
    const parts: string[] = [];
    let length = 0;
    for (let i = headingIndex + 1; i < scan.length; i++) {
      const lv = scan[i].level;
      if (lv !== null && lv <= level) break;
      const line = stripTrailingMarks(scan[i].text);
      if (length + line.length > SECTION_LIMIT) {
        parts.push("…（章节内容过长，已截断）");
        break;
      }
      parts.push(line);
      length += line.length;
    }
    return { title, level, content: parts.join("\n") };
  }

  /**
   * 超长文档降级（§12）：标题结构 + 当前章节 + 当前选区 + 关键词相关段落。
   */
  async getRelevantParagraphs(keywords: string[], limit = FALLBACK_RELEVANT_LIMIT): Promise<string[]> {
    const scan = await this.scanParagraphs();
    const picked: string[] = [];
    let length = 0;
    const terms = keywords.filter((k) => k.length >= 2).slice(0, 8);
    for (const p of scan) {
      if (length >= limit) break;
      if (p.level !== null) continue; // 大纲另行发送
      const line = stripTrailingMarks(p.text);
      if (!line.trim()) continue;
      const hit = terms.some((k) => line.includes(k));
      if (hit && length + line.length <= limit) {
        picked.push(line);
        length += line.length;
      }
    }
    return picked;
  }

  /** 文档统计信息（设置面板展示） */
  async getDocumentStats(): Promise<{ paragraphCount: number; textLength: number; truncated: boolean }> {
    return WordService.run(async (ctx) => {
      const body = ctx.document.body;
      body.load("text");
      // ParagraphCollection 无 getCount：加载轻量标量属性后用 items.length
      const paras = body.paragraphs;
      paras.load("style");
      await ctx.sync();
      const text = body.text ?? "";
      return {
        paragraphCount: paras.items.length,
        textLength: text.length,
        truncated: text.length > DOCUMENT_TEXT_LIMIT,
      };
    });
  }

  // ------------------------------------------------------------------

  private async scanParagraphs(): Promise<ParagraphScanItem[]> {
    return WordService.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load(["uniqueLocalId", "text", "styleBuiltIn", "style"]);
      await ctx.sync();
      // 表内段落探测（WordApi 1.3 null object 模式）：不在表格内时
      // parentTableCellOrNullObject.isNullObject === true。宿主不支持 1.3 时
      // 整批降级为 false（探测失败只损失 inTable 精度，不影响其余扫描）。
      let probeInTable = WordService.isSetSupported("1.3");
      if (probeInTable) {
        try {
          for (const p of paras.items) {
            p.parentTableCellOrNullObject.load("isNullObject");
          }
          await ctx.sync();
        } catch {
          probeInTable = false;
        }
      }
      return paras.items.map((p) => {
        const styleBuiltIn = String((p as { styleBuiltIn?: unknown }).styleBuiltIn ?? "");
        const style = String((p as { style?: unknown }).style ?? "");
        let inTable = false;
        if (probeInTable) {
          try {
            inTable = p.parentTableCellOrNullObject.isNullObject === false;
          } catch {
            inTable = false;
          }
        }
        return {
          id: p.uniqueLocalId,
          text: p.text ?? "",
          styleBuiltIn,
          style,
          level: headingLevel(styleBuiltIn, style),
          inTable,
        };
      });
    });
  }
}

/**
 * 识别 Heading 层级：
 * styleBuiltIn 为语言无关枚举（"Heading1"…）；
 * style 为本地化名称（中文 Word 为 "标题 1"）。
 */
export function headingLevel(styleBuiltIn: string, style: string): number | null {
  const m1 = /^Heading([1-9])$/.exec(styleBuiltIn ?? "");
  if (m1) return Number(m1[1]);
  const m2 = /^标题\s*([1-9])$/.exec(style ?? "");
  if (m2) return Number(m2[1]);
  const m3 = /^Heading\s*([1-9])$/i.exec(style ?? "");
  if (m3) return Number(m3[1]);
  return null;
}

/** 从用户消息提取关键词（简单启发式：按标点/空白切分，取较长词） */
export function extractKeywords(message: string): string[] {
  return message
    .split(/[\s，。；、！？：,.!?;:]+/)
    .map((t) => t.trim())
    .filter((t) => t.length >= 2);
}

/** 规范化全文（供哈希/比较） */
export function normalizedDocumentText(text: string): string {
  return normalizeText(text);
}

export const documentService = new DocumentService();

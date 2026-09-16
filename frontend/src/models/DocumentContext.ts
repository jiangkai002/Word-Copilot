/**
 * 文档上下文模型（对应需求文档 §13 DocumentContext）。
 *
 * 隐私原则（§60）：Selection / Paragraph 模式不携带全文，
 * 只有 @document 模式才允许发送大范围正文。
 */
export type ContextType = "selection" | "paragraph" | "section" | "document";

export interface ParagraphContext {
  /** Word.Paragraph.uniqueLocalId（仅当前会话内有效，WordApi 1.6） */
  uniqueLocalId?: string;
  text: string;
  style?: string;
  index?: number;
  textHash: string;
}

export interface SectionContext {
  title: string;
  level: number;
  content: string;
}

export interface DocumentOutlineItem {
  level: number;
  title: string;
}

export interface SelectionContext {
  text: string;
}

export interface DocumentContext {
  contextType: ContextType;
  documentId?: string;
  selection?: SelectionContext;
  currentParagraph?: ParagraphContext;
  previousParagraph?: ParagraphContext;
  nextParagraph?: ParagraphContext;
  section?: SectionContext;
  outline?: DocumentOutlineItem[];
  /** 仅 document 模式：全文或超长时的摘要正文 */
  documentText?: string;
  /** document 模式超长时截断标记 */
  truncated?: boolean;
}

export const CONTEXT_TYPE_LABELS: Record<ContextType, string> = {
  selection: "当前选区",
  paragraph: "当前段落",
  section: "当前章节",
  document: "整个文档",
};

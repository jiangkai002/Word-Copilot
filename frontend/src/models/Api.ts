/**
 * 与后端 API 交互的数据契约（对应需求文档 §16 / §18 / §44 / §45）。
 */
import type { DocumentContext } from "./DocumentContext";
import type { ContextType } from "./DocumentContext";

export interface HistoryEntry {
  role: "user" | "assistant";
  content: string;
}

export interface ChatStreamRequest {
  conversation_id: string;
  message: string;
  context?: DocumentContext | null;
  history?: HistoryEntry[];
}

/** SSE 事件负载 */
export interface SseTokenEvent {
  text: string;
}

export interface SseErrorEvent {
  code: string;
  message: string;
}

export interface SseDoneEvent {
  usage?: {
    prompt_tokens?: number;
    completion_tokens?: number;
  };
}

export interface EditTargetPayload {
  type: ContextType;
  text: string;
  text_hash: string;
}

export interface EditRequestPayload {
  conversation_id: string;
  instruction: string;
  target: EditTargetPayload;
  context?: {
    previous?: string;
    next?: string;
  } | null;
  /** 重新生成：提示模型避免重复上次输出 */
  regenerate_count?: number;
  avoid_texts?: string[];
}

export interface EditResponsePayload {
  edit_id: string;
  summary: string;
  original_text: string;
  new_text: string;
  warnings: string[];
}

export interface ApiErrorPayload {
  error: {
    code: string;
    message: string;
  };
}

// ---- Agent（§51：批量编辑 + 文档问答）----

export interface AgentParagraphPayload {
  id: string;
  text: string;
  style?: string | null;
  level?: number | null;
  /** 表格单元格内段落（Body.paragraphs 自 WordApi 1.3 起含表内段）——
   *  不能作为插入类提案的锚点，后端据此拒绝并引导模型改用文档末尾 */
  in_table?: boolean;
}

export interface AgentOutlineItemPayload {
  level: number;
  title: string;
}

export interface DocumentSnapshotPayload {
  outline: AgentOutlineItemPayload[];
  paragraphs: AgentParagraphPayload[];
  truncated: boolean;
}

export interface AgentFocusPayload {
  paragraph_id: string;
  paragraph_text: string;
  selected_text?: string | null;
}

export interface AgentStreamRequestPayload {
  conversation_id: string;
  instruction: string;
  history?: HistoryEntry[];
  snapshot: DocumentSnapshotPayload;
  /** Agent 模式：当前光标 / 选区焦点（「这段」的指代） */
  focus?: AgentFocusPayload | null;
}

/** SSE proposal 帧载荷：一条整段替换的修改提案（文本类，帧名 proposal） */
export interface AgentProposalEvent {
  paragraph_id: string;
  original_text: string;
  new_text: string;
  summary: string;
}

// ---- Agent 四类提案（proposal / proposal_format / proposal_table /
// proposal_formula / proposal_paragraph 帧）----

export type ProposalKind = "text" | "format" | "insert-table" | "insert-formula" | "insert-paragraph";

/** 文本提案（AgentApi 解析 proposal 帧时补 kind 判别字段） */
export interface AgentTextProposalEvent extends AgentProposalEvent {
  kind: "text";
}

/** 格式提案：None / 缺省 = 不修改该项；false = 显式关闭 */
export interface AgentFormatProposalEvent {
  kind: "format";
  paragraph_id: string;
  /** 后端从快照回显（前端哈希校验用） */
  original_text?: string;
  summary: string;
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  strikethrough?: boolean | null;
  font_name?: string | null;
  font_size?: number | null;
  color?: string | null;
  alignment?: "left" | "center" | "right" | "justify" | null;
}

/**
 * 插入类提案共通：anchor 为 null / 缺省 = 插入到文档末尾
 * （空文档可用；前端走 Body.insert*("End")）。
 */
export interface AgentInsertAnchorFields {
  anchor_paragraph_id?: string | null;
  anchor_text?: string;
}

/** 表格插入提案：锚点段之后 / 文档末尾 */
export interface AgentTableProposalEvent extends AgentInsertAnchorFields {
  kind: "insert-table";
  summary: string;
  header: boolean;
  values: string[][];
}

/** 公式插入提案：LaTeX → OMML，锚点段之后 / 文档末尾 */
export interface AgentFormulaProposalEvent extends AgentInsertAnchorFields {
  kind: "insert-formula";
  summary: string;
  latex: string;
  display: boolean;
}

/** 纯文字段落插入提案：paragraph_text 换行分段，锚点段之后 / 文档末尾 */
export interface AgentParagraphProposalEvent extends AgentInsertAnchorFields {
  kind: "insert-paragraph";
  summary: string;
  paragraph_text: string;
}

/** Agent 提案判别联合（edits.applyProposal 按 kind 分派） */
export type AgentProposal =
  | AgentTextProposalEvent
  | AgentFormatProposalEvent
  | AgentTableProposalEvent
  | AgentFormulaProposalEvent
  | AgentParagraphProposalEvent;

export interface SseAgentDoneEvent {
  usage?: {
    prompt_tokens?: number;
    completion_tokens?: number;
  };
  proposal_count?: number;
}

export interface HealthResponse {
  status: string;
  version: string;
}

export interface ConfigResponse {
  model: string;
  provider: string;
  supports_stream: boolean;
}

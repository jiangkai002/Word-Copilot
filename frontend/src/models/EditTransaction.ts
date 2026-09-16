/**
 * AI 编辑事务模型（对应需求文档 §23 EditTransaction）。
 *
 * 每一次 AI 修改都是一个独立事务：
 * Content Control（tag = word_ai_edit:{uuid}）框定本次修订范围，
 * Accept / Reject 只作用于该事务，绝不触碰用户人工修订（§55）。
 *
 * kind 判别事务类型（文本 / 格式 / 插入表格 / 插入公式），决定 apply 引擎、
 * 接受 / 拒绝语义与对账模式（RevisionService 按 kind 分派）。
 * 事务仅存于内存（Pinia ref，clear 即清）；若未来持久化，加载时缺失的
 * kind 必须默认为 "text"。
 */
import type { FormatBeforeValues, FormatChanges, RangeLocator } from "./EditPlan";

export type EditTransactionStatus = "pending" | "accepted" | "rejected" | "invalid";

export type TransactionKind =
  | "text"
  | "format"
  | "insert-table"
  | "insert-formula"
  | "insert-paragraph";

export interface EditTransaction {
  /** 展示用事务号，例如 AI_EDIT_001 */
  id: string;
  /** Content Control tag：word_ai_edit:{uuid} */
  contentControlTag: string;
  createdAt: number;
  status: EditTransactionStatus;
  /** 事务类型（accept / reject / 对账按此分派） */
  kind: TransactionKind;
  /** format：目标段文本（= newText）；insert-*：锚点段文本（文档末尾插入为空串） */
  originalText: string;
  /** text：新文本；format：同原文；insert-*：空串 */
  newText: string;
  summary?: string;
  /** 触发本次修改的用户指令 */
  instruction: string;
  /** 实际产生的 Word 修订条数（best effort；格式修改可能为 0 —— Word 不记录格式修订） */
  changeCount: number;
  /**
   * 目标定位信息（用于重新生成 / 乐观锁校验；insert-* 为锚点定位）；
   * insert-* 且锚点为文档末尾时为 null（无目标段落）
   */
  locator: RangeLocator | null;
  /** 目标类型：selection / paragraph（format 与 insert-* 恒为 paragraph） */
  targetKind: "selection" | "paragraph";
  /** 重新生成次数（仅 text 类支持） */
  regenerateCount: number;
  /** invalid 状态时的说明 */
  note?: string;
  // ---- 按类型的载荷 ----
  /** kind === "format"：请求的格式修改 */
  formatChanges?: FormatChanges;
  /** kind === "format"：应用前原值（拒绝时回写还原） */
  formatBefore?: FormatBeforeValues;
  /** kind === "insert-table" */
  tableValues?: string[][];
  tableHeader?: boolean;
  /** kind === "insert-formula" */
  latex?: string;
  displayFormula?: boolean;
  /** kind === "insert-paragraph"：每行一个段落的文本 */
  paragraphLines?: string[];
  /** insert-*：null = 插入到文档末尾（对账 / 展示用） */
  anchorId?: string | null;
}

export const TRANSACTION_STATUS_LABELS: Record<EditTransactionStatus, string> = {
  pending: "待处理",
  accepted: "已接受",
  rejected: "已拒绝",
  invalid: "写入失败",
};

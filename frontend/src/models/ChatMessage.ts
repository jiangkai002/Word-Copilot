/**
 * 聊天消息模型。
 *
 * 类型说明：
 * - user      用户输入
 * - assistant AI 回答（流式时 content 逐步追加）
 * - system    系统提示 / 错误 / 状态信息
 * - edit      AI 编辑结果卡片（关联 EditTransaction）
 */
export type ChatMessageType = "user" | "assistant" | "system" | "edit";

export interface ChatMessage {
  id: string;
  type: ChatMessageType;
  content: string;
  createdAt: number;
  /** assistant 消息流式输出中为 true */
  streaming?: boolean;
  /** system 消息的错误码（用于展示） */
  errorCode?: string;
  /** type === "edit" 时关联的事务 id */
  transactionId?: string;
}

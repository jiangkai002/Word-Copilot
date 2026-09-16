/**
 * 统一错误模型（对应需求文档 §45）。
 *
 * 所有 service 层错误都转为 CopilotError，
 * UI 层根据 code 展示友好中文提示。
 */

export type ErrorCode =
  | "LLM_TIMEOUT"
  | "LLM_ERROR"
  | "INVALID_EDIT_RESPONSE"
  | "DOCUMENT_CHANGED"
  | "UNSUPPORTED_WORD_VERSION"
  | "UNSUPPORTED_CONTENT"
  | "RANGE_NOT_FOUND"
  | "RANGE_AMBIGUOUS"
  | "REVISION_ERROR"
  | "FORMULA_INVALID"
  | "FORMULA_CONVERSION_ERROR"
  | "NETWORK_ERROR"
  | "ABORTED"
  | "PENDING_EDIT_CONFLICT"
  | "WORD_API_ERROR"
  | "BACKEND_ERROR"
  | "UNKNOWN";

const DEFAULT_MESSAGES: Record<ErrorCode, string> = {
  LLM_TIMEOUT: "模型请求超时，请稍后重试。",
  LLM_ERROR: "模型调用出错。",
  INVALID_EDIT_RESPONSE: "模型返回的修改结果无效。",
  DOCUMENT_CHANGED: "目标内容在 AI 生成期间已经发生变化，为避免覆盖你的修改，本次未应用。",
  UNSUPPORTED_WORD_VERSION: "当前 Word 版本不支持 AI 原生修订功能，请升级 Microsoft Word。",
  UNSUPPORTED_CONTENT: "当前内容包含第一版暂不支持的复杂 Word 对象，请重新选择普通文本内容。",
  RANGE_NOT_FOUND: "在文档中找不到目标内容，可能已被移动或删除。",
  RANGE_AMBIGUOUS: "目标内容在文档中出现多处，无法确定修改位置，请重新选中要修改的内容。",
  REVISION_ERROR: "以修订方式写入文档失败。",
  FORMULA_INVALID: "LaTeX 公式语法无法解析，未写入文档。",
  FORMULA_CONVERSION_ERROR: "公式无法转换为 Word 格式，未写入文档。",
  NETWORK_ERROR: "无法连接后端服务，请确认 FastAPI 已启动。",
  ABORTED: "已停止生成。",
  PENDING_EDIT_CONFLICT: "当前内容已有一个尚未处理的 AI 修改，请先接受或拒绝现有修改。",
  WORD_API_ERROR: "Word API 调用失败。",
  BACKEND_ERROR: "后端服务返回错误。",
  UNKNOWN: "发生未知错误。",
};

export class CopilotError extends Error {
  readonly code: ErrorCode;
  readonly detail?: string;

  constructor(code: ErrorCode, messageOrDetail?: string) {
    const base = DEFAULT_MESSAGES[code];
    const isKnownMessage = messageOrDetail && Object.values(DEFAULT_MESSAGES).includes(messageOrDetail);
    super(messageOrDetail && !isKnownMessage ? `${base}（${messageOrDetail}）` : base);
    this.name = "CopilotError";
    this.code = code;
    this.detail = messageOrDetail;
  }
}

/** 把任意异常归一化为 CopilotError */
export function toCopilotError(err: unknown): CopilotError {
  if (err instanceof CopilotError) return err;
  if (err instanceof DOMException && err.name === "AbortError") {
    return new CopilotError("ABORTED");
  }
  if (err instanceof Error) {
    // OfficeExtension.Error 带 code 字段（如 InvalidObjectPath / ItemNotFound / GeneralException）
    const officeCode = (err as { code?: unknown }).code;
    if (typeof officeCode === "string" && officeCode.length > 0) {
      return new CopilotError("WORD_API_ERROR", `${officeCode}: ${err.message}`);
    }
    if (err.name === "TypeError" || /fetch|network/i.test(err.message)) {
      return new CopilotError("NETWORK_ERROR", err.message);
    }
    return new CopilotError("UNKNOWN", err.message);
  }
  return new CopilotError("UNKNOWN", String(err));
}

/** 给用户看的提示文案 */
export function errorUserMessage(err: unknown): string {
  const ce = toCopilotError(err);
  return ce.message;
}

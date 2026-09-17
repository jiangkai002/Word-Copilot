/**
 * API 客户端基础层：
 * baseUrl 由 settings store 注入（避免 store ↔ service 循环依赖）。
 * 统一错误格式（§45）：{ error: { code, message } }
 */
import { CopilotError, type ErrorCode } from "@/utils/errors";
import { setLogBaseUrl } from "@/utils/logger";
import type { ApiErrorPayload } from "@/models/Api";

let baseUrl = "";

export function setApiBaseUrl(url: string): void {
  baseUrl = url.replace(/\/+$/, "");
  setLogBaseUrl(baseUrl);
}

export function getApiBaseUrl(): string {
  return baseUrl;
}

export function resolveUrl(path: string): string {
  return `${baseUrl}${path}`;
}

const BACKEND_ERROR_CODES: readonly string[] = [
  "LLM_TIMEOUT",
  "LLM_ERROR",
  "INVALID_EDIT_RESPONSE",
  "DOCUMENT_CHANGED",
  "UNSUPPORTED_WORD_VERSION",
  "UNSUPPORTED_CONTENT",
  "RANGE_NOT_FOUND",
  "RANGE_AMBIGUOUS",
  "REVISION_ERROR",
];

/** 后端错误码 → 前端错误码（§45 对齐），未知映射为 BACKEND_ERROR */
function mapBackendCode(code: string): ErrorCode {
  return BACKEND_ERROR_CODES.includes(code) ? (code as ErrorCode) : "BACKEND_ERROR";
}

export class ApiError extends CopilotError {
  constructor(code: string, message: string) {
    super(mapBackendCode(code), message);
    this.name = "ApiError";
  }
}

/** 解析非 2xx 响应并抛出 CopilotError */
export async function throwApiError(response: Response): Promise<never> {
  let code = "BACKEND_ERROR";
  let message = `后端返回 ${response.status}`;
  try {
    const body = (await response.json()) as ApiErrorPayload | { detail?: string };
    if (body && typeof body === "object" && "error" in body && body.error) {
      code = body.error.code ?? code;
      message = body.error.message ?? message;
    } else if (body && typeof body === "object" && "detail" in body) {
      message = String((body as { detail?: string }).detail ?? message);
    }
  } catch {
    // 响应体不是 JSON
  }
  throw new ApiError(code, message);
}

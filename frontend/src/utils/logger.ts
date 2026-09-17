/**
 * 日志（对应需求文档 §59）：
 * 开发模式详细输出；生产模式只输出 error / warning。
 */
const isDev = import.meta.env?.DEV ?? true;
let conversationId: string | null = null;
let logBaseUrl = "";

export function setLogConversationId(value: string | null): void {
  conversationId = value;
}

export function setLogBaseUrl(value: string): void {
  logBaseUrl = value.trim().replace(/\/+$/, "");
}

function serializeDetail(value: unknown): unknown {
  if (value instanceof Error) {
    const officeCode = (value as { code?: unknown }).code;
    return {
      name: value.name,
      message: value.message,
      stack: value.stack,
      ...(typeof officeCode === "string" ? { code: officeCode } : {}),
    };
  }
  try {
    JSON.stringify(value);
    return value;
  } catch {
    return String(value);
  }
}

/**
 * 将 warn/error 异步写入后端本地日志。失败时静默，避免日志上报递归影响主流程。
 */
function persist(level: "warn" | "error", args: unknown[]): void {
  if (typeof window === "undefined" || typeof fetch !== "function") return;
  const [head, ...rest] = args;
  const payload = {
    level,
    timestamp: new Date().toISOString(),
    conversation_id: conversationId,
    message: typeof head === "string" ? head : String(head),
    details: rest.map(serializeDetail),
    page: window.location.href,
  };
  void fetch(`${logBaseUrl}/api/v1/logs/client`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    keepalive: true,
  }).catch(() => undefined);
}

function levelEnabled(level: "debug" | "info" | "warn" | "error"): boolean {
  if (isDev) return true;
  return level === "warn" || level === "error";
}

export const logger = {
  debug(...args: unknown[]) {
    if (levelEnabled("debug")) console.debug("[WordAI]", ...args);
  },
  info(...args: unknown[]) {
    if (levelEnabled("info")) console.info("[WordAI]", ...args);
  },
  warn(...args: unknown[]) {
    if (levelEnabled("warn")) console.warn("[WordAI]", ...args);
    persist("warn", args);
  },
  error(...args: unknown[]) {
    if (levelEnabled("error")) console.error("[WordAI]", ...args);
    persist("error", args);
  },
};

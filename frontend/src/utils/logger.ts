/**
 * 日志（对应需求文档 §59）：
 * 开发模式详细输出；生产模式只输出 error / warning。
 */
const isDev = import.meta.env?.DEV ?? true;

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
  },
  error(...args: unknown[]) {
    if (levelEnabled("error")) console.error("[WordAI]", ...args);
  },
};

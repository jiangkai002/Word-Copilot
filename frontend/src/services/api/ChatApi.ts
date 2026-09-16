/**
 * ChatApi（对应需求文档 §16 / §43）：POST /api/v1/chat/stream，SSE 流式返回。
 *
 * 使用 fetch + ReadableStream 解析 SSE（EventSource 不支持 POST）。
 * 支持通过 AbortController 停止生成（§58）。
 */
import type {
  ChatStreamRequest,
  SseDoneEvent,
  SseErrorEvent,
  SseTokenEvent,
  ConfigResponse,
  HealthResponse,
} from "@/models/Api";
import { CopilotError } from "@/utils/errors";
import { resolveUrl, throwApiError } from "./ApiClient";

export interface StreamCallbacks {
  onToken?: (text: string) => void;
  onError?: (code: string, message: string) => void;
  onDone?: (event: SseDoneEvent) => void;
}

export class ChatApi {
  /**
   * 流式对话。resolve() 表示流正常或以 error 事件结束（错误经 onError 回调传递）。
   * 网络层异常 reject（CopilotError）。
   */
  async streamChat(
    request: ChatStreamRequest,
    callbacks: StreamCallbacks,
    signal?: AbortSignal,
  ): Promise<void> {
    let response: Response;
    try {
      response = await fetch(resolveUrl("/api/v1/chat/stream"), {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "text/event-stream" },
        body: JSON.stringify(request),
        signal,
      });
    } catch (err) {
      if (err instanceof DOMException && err.name === "AbortError") throw err;
      throw new CopilotError("NETWORK_ERROR", "无法连接后端（/api/v1/chat/stream）");
    }
    if (!response.ok) {
      await throwApiError(response);
    }
    if (!response.body) {
      throw new CopilotError("NETWORK_ERROR", "后端未返回流式响应体");
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const frames = buffer.split(/\n\n|\r\n\r\n/);
        buffer = frames.pop() ?? "";
        for (const frame of frames) {
          this.handleFrame(frame, callbacks);
        }
      }
      if (buffer.trim()) {
        this.handleFrame(buffer, callbacks);
      }
    } finally {
      reader.releaseLock();
    }
  }

  private handleFrame(frame: string, callbacks: StreamCallbacks): void {
    let event = "message";
    const dataLines: string[] = [];
    for (const line of frame.split(/\n|\r\n/)) {
      if (line.startsWith(":")) continue; // SSE 注释行
      if (line.startsWith("event:")) {
        event = line.slice(6).trim();
      } else if (line.startsWith("data:")) {
        dataLines.push(line.slice(5).trim());
      }
    }
    if (dataLines.length === 0) return;
    const raw = dataLines.join("\n");

    try {
      if (event === "token") {
        const payload = JSON.parse(raw) as SseTokenEvent;
        callbacks.onToken?.(payload.text ?? "");
      } else if (event === "error") {
        const payload = JSON.parse(raw) as SseErrorEvent;
        callbacks.onError?.(payload.code ?? "LLM_ERROR", payload.message ?? "模型调用出错");
      } else if (event === "done") {
        const payload = raw ? (JSON.parse(raw) as SseDoneEvent) : ({} as SseDoneEvent);
        callbacks.onDone?.(payload ?? {});
      }
    } catch {
      // 忽略无法解析的帧
    }
  }

  /** GET /api/v1/health */
  async health(signal?: AbortSignal): Promise<HealthResponse> {
    let response: Response;
    try {
      response = await fetch(resolveUrl("/api/v1/health"), { signal });
    } catch (err) {
      if (err instanceof DOMException && err.name === "AbortError") throw err;
      throw new CopilotError("NETWORK_ERROR", "无法连接后端服务");
    }
    if (!response.ok) await throwApiError(response);
    return (await response.json()) as HealthResponse;
  }

  /** GET /api/v1/config（可选接口 §44） */
  async config(signal?: AbortSignal): Promise<ConfigResponse> {
    let response: Response;
    try {
      response = await fetch(resolveUrl("/api/v1/config"), { signal });
    } catch (err) {
      if (err instanceof DOMException && err.name === "AbortError") throw err;
      throw new CopilotError("NETWORK_ERROR", "无法读取后端配置");
    }
    if (!response.ok) await throwApiError(response);
    return (await response.json()) as ConfigResponse;
  }
}

export const chatApi = new ChatApi();

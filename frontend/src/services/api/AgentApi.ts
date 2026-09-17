/**
 * AgentApi（§51）：POST /api/v1/agent/stream，SSE 双事件通道。
 *
 * - token：自然语言增量（流式气泡）
 * - proposal / proposal_format / proposal_table / proposal_formula /
 *   proposal_paragraph / proposal_heading：六类修改提案（文本 / 格式 / 表格 /
 *   公式 / 纯文字段落 / Word 标题），解析为 AgentProposal 判别联合（text 帧补 kind:"text"
 *   判别字段），收齐后逐条走编辑管线
 *
 * 插入类帧的 anchor_paragraph_id 可为 null（文档末尾）。
 * 畸形提案帧（缺判别字段 / values 非数组等）直接丢弃，不打断流。
 * SSE 解析复用 ChatApi 的 fetch + ReadableStream 模式（EventSource 不支持
 * POST）；AbortController 停止（§58）；错误经 onError 回调传递，网络层异常
 * reject（CopilotError）。
 */
import type {
  AgentFormatProposalEvent,
  AgentFormulaProposalEvent,
  AgentHeadingProposalEvent,
  AgentParagraphProposalEvent,
  AgentProposal,
  AgentStreamRequestPayload,
  AgentTableProposalEvent,
  AgentTextProposalEvent,
  SseAgentDoneEvent,
  SseErrorEvent,
  SseTokenEvent,
} from "@/models/Api";
import { CopilotError } from "@/utils/errors";
import { resolveUrl, throwApiError } from "./ApiClient";

export interface AgentStreamCallbacks {
  onToken?: (text: string) => void;
  onProposal?: (proposal: AgentProposal) => void;
  onError?: (code: string, message: string) => void;
  onDone?: (event: SseAgentDoneEvent) => void;
}

export class AgentApi {
  /**
   * 流式 Agent 任务。resolve() 表示流正常或以 error 事件结束
   * （错误经 onError 回调传递）；网络层异常 reject（CopilotError）。
   */
  async streamAgent(
    request: AgentStreamRequestPayload,
    callbacks: AgentStreamCallbacks,
    signal?: AbortSignal,
  ): Promise<void> {
    let response: Response;
    try {
      response = await fetch(resolveUrl("/api/v1/agent/stream"), {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "text/event-stream" },
        body: JSON.stringify(request),
        signal,
      });
    } catch (err) {
      if (err instanceof DOMException && err.name === "AbortError") throw err;
      throw new CopilotError("NETWORK_ERROR", "无法连接后端（/api/v1/agent/stream）");
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

  private handleFrame(frame: string, callbacks: AgentStreamCallbacks): void {
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
      } else if (event === "proposal") {
        const payload = JSON.parse(raw) as AgentTextProposalEvent;
        // 判别字段由前端补上（后端文本提案不带 kind，线格式向后兼容）
        if (payload.paragraph_id && payload.new_text) {
          callbacks.onProposal?.({ ...payload, kind: "text" });
        }
      } else if (event === "proposal_format") {
        const payload = JSON.parse(raw) as AgentFormatProposalEvent;
        if (this.isValidFormatProposal(payload)) {
          callbacks.onProposal?.(payload);
        }
      } else if (event === "proposal_table") {
        const payload = JSON.parse(raw) as AgentTableProposalEvent;
        if (this.isValidTableProposal(payload)) {
          callbacks.onProposal?.(payload);
        }
      } else if (event === "proposal_formula") {
        const payload = JSON.parse(raw) as AgentFormulaProposalEvent;
        if (this.isValidFormulaProposal(payload)) {
          callbacks.onProposal?.(payload);
        }
      } else if (event === "proposal_paragraph") {
        const payload = JSON.parse(raw) as AgentParagraphProposalEvent;
        if (this.isValidParagraphProposal(payload)) {
          callbacks.onProposal?.(payload);
        }
      } else if (event === "proposal_heading") {
        const payload = JSON.parse(raw) as AgentHeadingProposalEvent;
        if (this.isValidHeadingProposal(payload)) {
          callbacks.onProposal?.(payload);
        }
      } else if (event === "error") {
        const payload = JSON.parse(raw) as SseErrorEvent;
        callbacks.onError?.(payload.code ?? "LLM_ERROR", payload.message ?? "模型调用出错");
      } else if (event === "done") {
        const payload = raw ? (JSON.parse(raw) as SseAgentDoneEvent) : ({} as SseAgentDoneEvent);
        callbacks.onDone?.(payload ?? {});
      }
    } catch {
      // 忽略无法解析的帧
    }
  }

  /** proposal_format 帧校验：指向段落 + 至少一个格式字段（store 侧再做完整校验） */
  private isValidFormatProposal(payload: AgentFormatProposalEvent): boolean {
    if (!payload.paragraph_id) return false;
    const fields = [
      payload.bold,
      payload.italic,
      payload.underline,
      payload.strikethrough,
      payload.font_name,
      payload.font_size,
      payload.color,
      payload.alignment,
    ];
    return fields.some((v) => v !== null && v !== undefined);
  }

  /** proposal_table 帧校验：values 是非空二维字符串数组（anchor 可为 null = 文档末尾） */
  private isValidTableProposal(payload: AgentTableProposalEvent): boolean {
    if (!Array.isArray(payload.values) || payload.values.length === 0) {
      return false;
    }
    return payload.values.every(
      (row) => Array.isArray(row) && row.length > 0 && row.every((cell) => typeof cell === "string"),
    );
  }

  /** proposal_formula 帧校验：非空 latex（anchor 可为 null = 文档末尾） */
  private isValidFormulaProposal(payload: AgentFormulaProposalEvent): boolean {
    return typeof payload.latex === "string" && payload.latex.length > 0;
  }

  /** proposal_paragraph 帧校验：非空 paragraph_text（anchor 可为 null = 文档末尾） */
  private isValidParagraphProposal(payload: AgentParagraphProposalEvent): boolean {
    return typeof payload.paragraph_text === "string" && payload.paragraph_text.trim().length > 0;
  }

  /** proposal_heading：非空单行标题 + Word 标题级别 1～9。 */
  private isValidHeadingProposal(payload: AgentHeadingProposalEvent): boolean {
    return (
      payload.kind === "insert-heading" &&
      typeof payload.heading_text === "string" &&
      payload.heading_text.trim().length > 0 &&
      !/[\r\n\v]/.test(payload.heading_text) &&
      Number.isInteger(payload.level) &&
      payload.level >= 1 &&
      payload.level <= 9
    );
  }
}

export const agentApi = new AgentApi();

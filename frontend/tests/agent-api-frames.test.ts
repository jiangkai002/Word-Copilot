/**
 * AgentApi SSE 帧解析测试：六类提案帧（proposal / proposal_format /
 * proposal_table / proposal_formula / proposal_paragraph / proposal_heading）→ AgentProposal
 * 判别联合；插入类 anchor 可为 null（文档末尾）；畸形提案帧丢弃；
 * error / token / done 帧语义不变。
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { AgentApi } from "@/services/api/AgentApi";
import type { AgentProposal } from "@/models/Api";

/** 手工构造 SSE 流式响应（每帧以空行分隔） */
function sseResponse(frames: string[]): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const frame of frames) controller.enqueue(encoder.encode(frame));
      controller.close();
    },
  });
  return { ok: true, status: 200, body: stream } as unknown as Response;
}

function frame(event: string, data: string): string {
  return `event: ${event}\ndata: ${data}\n\n`;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("AgentApi.streamAgent 帧解析", () => {
  it("六类提案帧依序触发 onProposal（text 帧补 kind 判别字段）", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        sseResponse([
          frame("token", JSON.stringify({ text: "处理中" })),
          frame(
            "proposal",
            JSON.stringify({ paragraph_id: "p1", original_text: "原文", new_text: "新文", summary: "纠错" }),
          ),
          frame(
            "proposal_format",
            JSON.stringify({ kind: "format", paragraph_id: "p2", summary: "设为标题 2", paragraph_style: "heading2" }),
          ),
          frame(
            "proposal_table",
            JSON.stringify({
              kind: "insert-table",
              anchor_paragraph_id: "p3",
              summary: "插入表格",
              header: true,
              values: [
                ["列A", "列B"],
                ["1", "2"],
              ],
            }),
          ),
          frame(
            "proposal_formula",
            JSON.stringify({
              kind: "insert-formula",
              anchor_paragraph_id: "p3",
              summary: "插入公式",
              latex: "E=mc^2",
              display: true,
            }),
          ),
          frame(
            "proposal_paragraph",
            JSON.stringify({
              kind: "insert-paragraph",
              anchor_paragraph_id: "p1",
              anchor_text: "原文",
              paragraph_text: "第一段\n第二段",
              summary: "插入文字",
            }),
          ),
          frame(
            "proposal_heading",
            JSON.stringify({
              kind: "insert-heading",
              anchor_paragraph_id: "p1",
              heading_text: "第一章 系统概述",
              level: 1,
              summary: "插入章标题",
            }),
          ),
          frame("done", JSON.stringify({ proposal_count: 6 })),
        ]),
      ),
    );

    const proposals: AgentProposal[] = [];
    const tokens: string[] = [];
    const dones: unknown[] = [];
    await new AgentApi().streamAgent(
      { conversation_id: "c1", instruction: "x", snapshot: { outline: [], paragraphs: [], truncated: false } },
      {
        onToken: (t) => tokens.push(t),
        onProposal: (p) => proposals.push(p),
        onDone: (d) => dones.push(d),
      },
    );

    expect(tokens).toEqual(["处理中"]);
    expect(proposals.map((p) => p.kind)).toEqual([
      "text",
      "format",
      "insert-table",
      "insert-formula",
      "insert-paragraph",
      "insert-heading",
    ]);
    expect(proposals[0]).toMatchObject({ kind: "text", paragraph_id: "p1", new_text: "新文" });
    expect(proposals[1]).toMatchObject({ kind: "format", paragraph_id: "p2", paragraph_style: "heading2" });
    expect(proposals[2]).toMatchObject({ kind: "insert-table", anchor_paragraph_id: "p3", header: true });
    expect((proposals[2] as { values: string[][] }).values).toEqual([
      ["列A", "列B"],
      ["1", "2"],
    ]);
    expect(proposals[3]).toMatchObject({ kind: "insert-formula", latex: "E=mc^2", display: true });
    expect(proposals[4]).toMatchObject({
      kind: "insert-paragraph",
      paragraph_text: "第一段\n第二段",
      anchor_paragraph_id: "p1",
    });
    expect(proposals[5]).toMatchObject({ kind: "insert-heading", heading_text: "第一章 系统概述", level: 1 });
    expect(dones).toEqual([{ proposal_count: 6 }]);
  });

  it("插入类提案 anchor 为 null（文档末尾 / 空文档）仍为合法帧", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        sseResponse([
          frame(
            "proposal_table",
            JSON.stringify({
              kind: "insert-table",
              anchor_paragraph_id: null,
              summary: "空文档里画表格",
              header: true,
              values: [["列A", "列B"]],
            }),
          ),
          frame(
            "proposal_paragraph",
            JSON.stringify({
              kind: "insert-paragraph",
              anchor_paragraph_id: null,
              anchor_text: "",
              paragraph_text: "随便写点内容",
              summary: "插入文字",
            }),
          ),
          frame(
            "proposal_heading",
            JSON.stringify({
              kind: "insert-heading",
              anchor_paragraph_id: null,
              heading_text: "项目概述",
              level: 1,
              summary: "插入标题",
            }),
          ),
        ]),
      ),
    );

    const proposals: AgentProposal[] = [];
    await new AgentApi().streamAgent(
      { conversation_id: "c1", instruction: "x", snapshot: { outline: [], paragraphs: [], truncated: false } },
      { onProposal: (p) => proposals.push(p) },
    );
    expect(proposals.map((p) => p.kind)).toEqual(["insert-table", "insert-paragraph", "insert-heading"]);
    expect(proposals.every((p) => "anchor_paragraph_id" in p && p.anchor_paragraph_id === null)).toBe(true);
  });

  it("畸形提案帧丢弃：缺判别字段 / 空格式字段 / values 非数组 / 缺 latex", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        sseResponse([
          // proposal 缺 new_text
          frame("proposal", JSON.stringify({ paragraph_id: "p1", original_text: "x", summary: "s" })),
          // proposal_format 无任何格式字段
          frame("proposal_format", JSON.stringify({ kind: "format", paragraph_id: "p2", summary: "s" })),
          // proposal_format 非法内置样式
          frame("proposal_format", JSON.stringify({ kind: "format", paragraph_id: "p2", summary: "s", paragraph_style: "heading10" })),
          // proposal_table values 非二维数组
          frame(
            "proposal_table",
            JSON.stringify({ kind: "insert-table", anchor_paragraph_id: "p3", summary: "s", header: true, values: "不是数组" }),
          ),
          // proposal_table 含非字符串单元格
          frame(
            "proposal_table",
            JSON.stringify({ kind: "insert-table", anchor_paragraph_id: "p3", summary: "s", header: true, values: [[1, 2]] }),
          ),
          // proposal_formula 缺 latex
          frame("proposal_formula", JSON.stringify({ kind: "insert-formula", anchor_paragraph_id: "p3", summary: "s", display: true })),
          // proposal_paragraph 缺 paragraph_text / 空白
          frame(
            "proposal_paragraph",
            JSON.stringify({ kind: "insert-paragraph", anchor_paragraph_id: null, summary: "s" }),
          ),
          frame("proposal_heading", JSON.stringify({ kind: "insert-heading", heading_text: "", level: 1, summary: "s" })),
          frame("proposal_heading", JSON.stringify({ kind: "insert-heading", heading_text: "标题", level: 10, summary: "s" })),
          frame("proposal_heading", JSON.stringify({ kind: "insert-heading", heading_text: "多行\n标题", level: 2, summary: "s" })),
          frame(
            "proposal_paragraph",
            JSON.stringify({ kind: "insert-paragraph", anchor_paragraph_id: null, paragraph_text: "   ", summary: "s" }),
          ),
          // 无法解析的 JSON
          frame("proposal_format", "{bad json"),
        ]),
      ),
    );

    const proposals: AgentProposal[] = [];
    await new AgentApi().streamAgent(
      { conversation_id: "c1", instruction: "x", snapshot: { outline: [], paragraphs: [], truncated: false } },
      { onProposal: (p) => proposals.push(p) },
    );
    expect(proposals).toEqual([]);
  });

  it("error 帧经 onError 回调传递（流仍正常 resolve）", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        sseResponse([
          frame("token", JSON.stringify({ text: "开头" })),
          frame("error", JSON.stringify({ code: "LLM_TIMEOUT", message: "模型调用超时" })),
          frame("done", JSON.stringify({})),
        ]),
      ),
    );

    const errors: Array<[string, string]> = [];
    await new AgentApi().streamAgent(
      { conversation_id: "c1", instruction: "x", snapshot: { outline: [], paragraphs: [], truncated: false } },
      { onError: (code, message) => errors.push([code, message]) },
    );
    expect(errors).toEqual([["LLM_TIMEOUT", "模型调用超时"]]);
  });

  it("网络异常 reject（CopilotError NETWORK_ERROR）", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockRejectedValue(new Error("connection refused")),
    );
    await expect(
      new AgentApi().streamAgent(
        { conversation_id: "c1", instruction: "x", snapshot: { outline: [], paragraphs: [], truncated: false } },
        {},
      ),
    ).rejects.toThrow("无法连接后端");
  });
});

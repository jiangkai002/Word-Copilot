/**
 * ProposalMapper 单元测试（§51 纯函数）：
 * 提案 → CapturedTarget 映射 —— 快照段落文本为权威原文
 * （text/locator.originalText 取快照而非 proposal.original_text，
 * 保证 applyProposal 的回显校验有意义）。
 *
 * 五类提案（判别联合）：text / format 用 paragraph_id 定位目标段；
 * insert-table / insert-formula / insert-paragraph 用 anchor_paragraph_id
 * 定位锚点段；anchor 为 null / 缺省 = 文档末尾（无目标段，映射为 null）。
 */
import { describe, expect, it } from "vitest";
import {
  proposalTargetId,
  proposalToCapturedTarget,
} from "@/services/agent/ProposalMapper";
import type { SnapshotParagraph } from "@/services/word/DocumentService";

const PARAGRAPHS: SnapshotParagraph[] = [
  { id: "p1", text: "第一章 概述", style: "标题 1", level: 1 },
  { id: "p2", text: "系统采用传统架构，可以处理大量数据。", style: "正文", level: null },
  { id: "p3", text: "这个词这个词这个词出现了三次这个词。", style: "正文", level: null },
];

describe("proposalToCapturedTarget", () => {
  it("映射为段落模式目标：快照文本为权威原文 + 前后段上下文", async () => {
    const target = await proposalToCapturedTarget(
      {
        kind: "text",
        paragraph_id: "p2",
        original_text: "系统采用传统架构，可以处理大量数据。",
        new_text: "新文本",
        summary: "纠错",
      },
      PARAGRAPHS,
    );
    expect(target).not.toBeNull();
    expect(target?.kind).toBe("paragraph");
    expect(target?.text).toBe("系统采用传统架构，可以处理大量数据。");
    expect(target?.prefix).toBe("第一章 概述");
    expect(target?.suffix).toBe("这个词这个词这个词出现了三次这个词。");
    expect(target?.locator.paragraphId).toBe("p2");
    expect(target?.locator.originalText).toBe("系统采用传统架构，可以处理大量数据。");
    expect(target?.textHash).toBe(target?.locator.textHash);
  });

  it("段首段落 prefix 为空串；段尾段落 suffix 为空串", async () => {
    const first = await proposalToCapturedTarget(
      {
        kind: "text",
        paragraph_id: "p1",
        original_text: "第一章 概述",
        new_text: "第二章 概述",
        summary: "",
      },
      PARAGRAPHS,
    );
    expect(first?.prefix).toBe("");
    expect(first?.suffix).toBe("系统采用传统架构，可以处理大量数据。");

    const last = await proposalToCapturedTarget(
      {
        kind: "text",
        paragraph_id: "p3",
        original_text: "这个词这个词这个词出现了三次这个词。",
        new_text: "新",
        summary: "",
      },
      PARAGRAPHS,
    );
    expect(last?.suffix).toBe("");
    expect(last?.prefix).toBe("系统采用传统架构，可以处理大量数据。");
  });

  it("权威原文取快照文本而非 proposal.original_text（回显校验才有意义）", async () => {
    const target = await proposalToCapturedTarget(
      {
        kind: "text",
        paragraph_id: "p2",
        original_text: "（被篡改的原文）",
        new_text: "新文本",
        summary: "",
      },
      PARAGRAPHS,
    );
    // target.text 仍为快照段落文本；applyProposal 会因二者不一致拒绝该提案
    expect(target?.text).toBe("系统采用传统架构，可以处理大量数据。");
    expect(target?.text).not.toBe("（被篡改的原文）");
  });

  it("paragraph_id 不在快照中 → null", async () => {
    const target = await proposalToCapturedTarget(
      { kind: "text", paragraph_id: "p99", original_text: "x", new_text: "y", summary: "" },
      PARAGRAPHS,
    );
    expect(target).toBeNull();
  });

  it("format 提案按 paragraph_id 定位目标段", async () => {
    const target = await proposalToCapturedTarget(
      { kind: "format", paragraph_id: "p2", summary: "加粗", bold: true },
      PARAGRAPHS,
    );
    expect(target).not.toBeNull();
    expect(target?.locator.paragraphId).toBe("p2");
    expect(target?.text).toBe("系统采用传统架构，可以处理大量数据。");
  });

  it("insert-table 提案按 anchor_paragraph_id 定位锚点段", async () => {
    const target = await proposalToCapturedTarget(
      {
        kind: "insert-table",
        anchor_paragraph_id: "p3",
        summary: "插入统计表",
        header: true,
        values: [
          ["列A", "列B"],
          ["1", "2"],
        ],
      },
      PARAGRAPHS,
    );
    expect(target).not.toBeNull();
    expect(target?.locator.paragraphId).toBe("p3");
    expect(target?.text).toBe("这个词这个词这个词出现了三次这个词。");
  });

  it("anchor_paragraph_id 不在快照中 → null", async () => {
    const target = await proposalToCapturedTarget(
      {
        kind: "insert-formula",
        anchor_paragraph_id: "p99",
        summary: "插入公式",
        latex: "E=mc^2",
        display: true,
      },
      PARAGRAPHS,
    );
    expect(target).toBeNull();
  });

  it("insert-paragraph 提案按 anchor_paragraph_id 定位锚点段；anchor 为 null → null（文档末尾）", async () => {
    const anchored = await proposalToCapturedTarget(
      {
        kind: "insert-paragraph",
        anchor_paragraph_id: "p1",
        paragraph_text: "新段落",
        summary: "插入文字",
      },
      PARAGRAPHS,
    );
    expect(anchored).not.toBeNull();
    expect(anchored?.locator.paragraphId).toBe("p1");

    // 文档末尾模式：无锚点段（空文档 / 未指明位置）
    const end = await proposalToCapturedTarget(
      {
        kind: "insert-paragraph",
        anchor_paragraph_id: null,
        paragraph_text: "随便写的内容",
        summary: "插入文字",
      },
      PARAGRAPHS,
    );
    expect(end).toBeNull();

    // anchor 缺省（字段不存在）同 null
    const omitted = await proposalToCapturedTarget(
      {
        kind: "insert-table",
        summary: "空文档画表格",
        header: true,
        values: [["a", "b"]],
      } as unknown as import("@/models/Api").AgentTableProposalEvent,
      PARAGRAPHS,
    );
    expect(omitted).toBeNull();
  });
});

describe("proposalTargetId", () => {
  it("text / format → paragraph_id；insert-* → anchor_paragraph_id", () => {
    expect(
      proposalTargetId({
        kind: "text",
        paragraph_id: "p1",
        original_text: "第一章 概述",
        new_text: "改",
        summary: "",
      }),
    ).toBe("p1");
    expect(proposalTargetId({ kind: "format", paragraph_id: "p2", summary: "", bold: true })).toBe("p2");
    expect(
      proposalTargetId({
        kind: "insert-table",
        anchor_paragraph_id: "p3",
        summary: "",
        header: false,
        values: [["a"]],
      }),
    ).toBe("p3");
    expect(
      proposalTargetId({
        kind: "insert-formula",
        anchor_paragraph_id: "p2",
        summary: "",
        latex: "a+b",
        display: false,
      }),
    ).toBe("p2");
    expect(
      proposalTargetId({
        kind: "insert-paragraph",
        anchor_paragraph_id: "p1",
        paragraph_text: "新段落",
        summary: "",
      }),
    ).toBe("p1");
    expect(
      proposalTargetId({
        kind: "insert-heading",
        anchor_paragraph_id: "p2",
        heading_text: "新增小节",
        level: 2,
        summary: "",
      }),
    ).toBe("p2");
  });

  it("insert-* 的文档末尾模式 → null（anchor 为 null 或缺省）", () => {
    expect(
      proposalTargetId({
        kind: "insert-paragraph",
        anchor_paragraph_id: null,
        paragraph_text: "文字",
        summary: "",
      }),
    ).toBeNull();
    expect(
      proposalTargetId({
        kind: "insert-table",
        summary: "",
        header: true,
        values: [["a"]],
      } as unknown as import("@/models/Api").AgentTableProposalEvent),
    ).toBeNull();
    expect(
      proposalTargetId({
        kind: "insert-heading",
        anchor_paragraph_id: null,
        heading_text: "第一章",
        level: 1,
        summary: "",
      }),
    ).toBeNull();
  });
});

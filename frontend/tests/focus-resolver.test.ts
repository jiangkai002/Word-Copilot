/**
 * FocusResolver 单元测试（§51 纯函数）：
 * 光标 / 选区焦点 → 请求 focus 载荷 —— 段落 id 与文本以快照为权威
 * （uniqueLocalId 未命中时按文本唯一匹配回退，无法消歧则不发送焦点）。
 */
import { describe, expect, it } from "vitest";
import { resolveFocus, type AgentFocusInput } from "@/services/agent/FocusResolver";
import type { SnapshotParagraph } from "@/services/word/DocumentService";

const PARAGRAPHS: SnapshotParagraph[] = [
  { id: "p1", text: "第一章 概述", style: "标题 1", level: 1 },
  { id: "p2", text: "系统采用传统架构，可以处理大量数据。", style: "正文", level: null },
  { id: "p3", text: "同一个文本。", style: "正文", level: null },
  { id: "p4", text: "同一个文本。", style: "正文", level: null },
];

describe("resolveFocus", () => {
  it("id 命中快照 → 采用快照 id 与文本（选中文本透传）", () => {
    const focus: AgentFocusInput = {
      paragraphId: "p2",
      paragraphText: "系统采用传统架构，可以处理大量数据。",
      selectedText: "传统架构",
    };
    expect(resolveFocus(focus, PARAGRAPHS)).toEqual({
      paragraph_id: "p2",
      paragraph_text: "系统采用传统架构，可以处理大量数据。",
      selected_text: "传统架构",
    });
  });

  it("id 未命中但文本唯一匹配 → 回退到快照段落（快照 id 为权威）", () => {
    // 快照 id 缺失时前端用 para-{i} 兜底，与 uniqueLocalId 不同
    const focus: AgentFocusInput = {
      paragraphId: "uuid-xyz-不在快照",
      paragraphText: "第一章 概述",
      selectedText: null,
    };
    expect(resolveFocus(focus, PARAGRAPHS)?.paragraph_id).toBe("p1");
  });

  it("文本在快照中重复 → 无法消歧 → null", () => {
    const focus: AgentFocusInput = {
      paragraphId: null,
      paragraphText: "同一个文本。",
      selectedText: null,
    };
    expect(resolveFocus(focus, PARAGRAPHS)).toBeNull();
  });

  it("光标段不在快照内（如被截断）→ null", () => {
    const focus: AgentFocusInput = {
      paragraphId: "p9",
      paragraphText: "快照之外的段落。",
      selectedText: null,
    };
    expect(resolveFocus(focus, PARAGRAPHS)).toBeNull();
  });

  it("paragraphId 为 null 时仅按文本解析", () => {
    const focus: AgentFocusInput = {
      paragraphId: null,
      paragraphText: "系统采用传统架构，可以处理大量数据。",
      selectedText: null,
    };
    expect(resolveFocus(focus, PARAGRAPHS)?.paragraph_id).toBe("p2");
  });
});

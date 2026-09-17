/**
 * 公式 / 段落插入事务：先登记审核卡，再写入 Word 修订；
 * 普通文字与含行内公式的段落使用同一审核语义。
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import { createPinia, setActivePinia } from "pinia";
import { insertEngine } from "@/services/word/InsertEngine";
import { useEditStore } from "@/stores/edits";

vi.mock("@/services/word/InsertEngine", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@/services/word/InsertEngine")>();
  return {
    ...actual,
    insertEngine: {
      applyTableInsertPlan: vi.fn(),
      applyFormulaInsertPlan: vi.fn(),
      applyParagraphInsertPlan: vi.fn(),
      applyHeadingInsertPlan: vi.fn(),
      commitFormulaInsertPlan: vi.fn(),
      commitParagraphInsertPlan: vi.fn(),
      commitHeadingInsertPlan: vi.fn(),
    },
  };
});

const outcome = {
  contentControlTag: "word_ai_edit:test",
  appliedOps: 1,
  skipped: [],
  changeCount: 1,
};

describe("公式与段落修订事务", () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
  });

  it("独立公式以 Word 修订写入并保留待审卡片", async () => {
    vi.mocked(insertEngine.applyFormulaInsertPlan).mockResolvedValue(outcome);
    const store = useEditStore();
    const result = await store.applyProposal(
      "插入公式",
      {
        kind: "insert-formula",
        anchor_paragraph_id: null,
        summary: "插入质能方程",
        latex: "E=mc^2",
        display: true,
      },
      [],
    );

    expect(result).toEqual({ ok: true });
    expect(insertEngine.applyFormulaInsertPlan).toHaveBeenCalledOnce();
    expect(store.pendingList).toHaveLength(1);
    expect(store.pendingList[0].changeCount).toBe(1);
  });

  it("热更新后新 Store 可回退调用旧版公式插入方法", async () => {
    const engine = insertEngine as unknown as {
      applyFormulaInsertPlan?: typeof insertEngine.applyFormulaInsertPlan;
      commitFormulaInsertPlan: typeof insertEngine.commitFormulaInsertPlan;
    };
    const currentApply = engine.applyFormulaInsertPlan;
    engine.applyFormulaInsertPlan = undefined;
    vi.mocked(engine.commitFormulaInsertPlan).mockResolvedValue(outcome);

    try {
      const store = useEditStore();
      const result = await store.applyProposal(
        "插入公式",
        {
          kind: "insert-formula",
          anchor_paragraph_id: null,
          summary: "插入质能方程",
          latex: "E=mc^2",
          display: true,
        },
        [],
      );

      expect(result).toEqual({ ok: true });
      expect(engine.commitFormulaInsertPlan).toHaveBeenCalledOnce();
    } finally {
      engine.applyFormulaInsertPlan = currentApply;
    }
  });

  it.each([
    ["纯文字", "该公式说明指数、复数和三角函数之间存在紧密联系。"],
    ["含行内公式", "质能方程 $E=mc^2$ 表明质量与能量可以相互转化。"],
  ])("%s段落使用同一 Word 修订管线", async (_label, paragraphText) => {
    vi.mocked(insertEngine.applyParagraphInsertPlan).mockResolvedValue(outcome);
    const store = useEditStore();
    const result = await store.applyProposal(
      "写入段落",
      {
        kind: "insert-paragraph",
        anchor_paragraph_id: null,
        summary: "写入说明",
        paragraph_text: paragraphText,
      },
      [],
    );

    expect(result).toEqual({ ok: true });
    expect(insertEngine.applyParagraphInsertPlan).toHaveBeenCalledOnce();
    expect(store.pendingList).toHaveLength(1);
    expect(store.pendingList[0].changeCount).toBe(1);
  });

  it("Word 写入尚未完成时审核卡已经登记", async () => {
    let finish!: (value: typeof outcome) => void;
    vi.mocked(insertEngine.applyParagraphInsertPlan).mockReturnValue(
      new Promise((resolve) => {
        finish = resolve;
      }),
    );
    const store = useEditStore();
    const applying = store.applyProposal(
      "补充解释",
      {
        kind: "insert-paragraph",
        anchor_paragraph_id: null,
        summary: "补充解释",
        paragraph_text: "这是需要审核的解释文字。",
      },
      [],
    );

    await Promise.resolve();
    expect(store.pendingList).toHaveLength(1);
    finish(outcome);
    expect(await applying).toEqual({ ok: true });
  });

  it("Word 标题使用独立修订事务并保存级别", async () => {
    vi.mocked(insertEngine.applyHeadingInsertPlan).mockResolvedValue(outcome);
    const store = useEditStore();
    const result = await store.applyProposal(
      "创建第一章",
      {
        kind: "insert-heading",
        anchor_paragraph_id: null,
        summary: "创建章标题",
        heading_text: "第一章 系统概述",
        level: 1,
      },
      [],
    );

    expect(result).toEqual({ ok: true });
    expect(insertEngine.applyHeadingInsertPlan).toHaveBeenCalledOnce();
    expect(store.pendingList[0]).toMatchObject({
      kind: "insert-heading",
      headingText: "第一章 系统概述",
      headingLevel: 1,
      changeCount: 1,
    });
  });

  it("Word 写入失败时保留失败卡片，不留下无卡片状态", async () => {
    vi.mocked(insertEngine.applyParagraphInsertPlan).mockRejectedValue(new Error("write failed"));
    const store = useEditStore();
    const result = await store.applyProposal(
      "补充解释",
      {
        kind: "insert-paragraph",
        anchor_paragraph_id: null,
        summary: "补充解释",
        paragraph_text: "这是需要审核的解释文字。",
      },
      [],
    );

    expect(result.ok).toBe(false);
    const tx = store.getTransaction(store.order[0]);
    expect(tx?.status).toBe("invalid");
    expect(tx?.note).toContain("写入 Word 修订失败");
  });
});

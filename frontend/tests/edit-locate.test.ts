/**
 * 编辑卡片点击定位（stores/edits.locate → RevisionService.jumpToTransaction）：
 * - pending：按事务 tag 跳转（事务控件存在），兜底文本原文优先
 * - accepted：控件已移除，兜底文本新文本优先（正文已是新文本）
 * - 未能定位 / 异常：给出系统提示，working 状态正确复位
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import { createPinia, setActivePinia } from "pinia";
import { insertEngine } from "@/services/word/InsertEngine";
import { revisionService } from "@/services/word/RevisionService";
import { useChatStore } from "@/stores/chat";
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

vi.mock("@/services/word/RevisionService", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@/services/word/RevisionService")>();
  return {
    ...actual,
    revisionService: {
      jumpToTransaction: vi.fn(),
    },
  };
});

const outcome = {
  contentControlTag: "word_ai_edit:test",
  appliedOps: 1,
  skipped: [],
  changeCount: 1,
};

/** 经 applyProposal(insert-paragraph) 登记一个事务并返回其 id */
async function createParagraphTx(paragraphText: string): Promise<string> {
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
  if (!result.ok) throw new Error("事务登记失败");
  return store.order[0];
}

describe("编辑卡片点击定位", () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
  });

  it("pending：按事务 tag 跳转，兜底文本原文优先（插入类原文为空）", async () => {
    vi.mocked(revisionService.jumpToTransaction).mockResolvedValue(true);
    const store = useEditStore();
    const id = await createParagraphTx("第一段内容\n第二段内容");

    await store.locate(id);

    expect(revisionService.jumpToTransaction).toHaveBeenCalledWith("word_ai_edit:test", [
      "",
      "第一段内容\n第二段内容",
    ]);
    expect(store.working).toBe(false);
  });

  it("accepted：控件已移除，兜底文本新文本优先", async () => {
    vi.mocked(revisionService.jumpToTransaction).mockResolvedValue(true);
    const store = useEditStore();
    const id = await createParagraphTx("第一段内容\n第二段内容");
    store.transactions[id].status = "accepted";

    await store.locate(id);

    expect(revisionService.jumpToTransaction).toHaveBeenCalledWith("word_ai_edit:test", [
      "第一段内容\n第二段内容",
      "",
    ]);
  });

  it("未能定位时给出系统提示", async () => {
    vi.mocked(revisionService.jumpToTransaction).mockResolvedValue(false);
    const store = useEditStore();
    const chatStore = useChatStore();
    const id = await createParagraphTx("第一段内容");

    await store.locate(id);

    const system = chatStore.messages.filter((m) => m.type === "system");
    expect(system.some((m) => m.content.includes("未能定位"))).toBe(true);
    expect(store.working).toBe(false);
  });

  it("定位抛错时转为系统提示，不外抛", async () => {
    vi.mocked(revisionService.jumpToTransaction).mockRejectedValue(new Error("word offline"));
    const store = useEditStore();
    const chatStore = useChatStore();
    const id = await createParagraphTx("第一段内容");

    await expect(store.locate(id)).resolves.toBeUndefined();

    const system = chatStore.messages.filter((m) => m.type === "system");
    expect(system.some((m) => m.content.includes("定位失败"))).toBe(true);
    expect(store.working).toBe(false);
  });

  it("working 期间点击不触发定位", async () => {
    vi.mocked(revisionService.jumpToTransaction).mockResolvedValue(true);
    const store = useEditStore();
    const id = await createParagraphTx("第一段内容");
    store.working = true;

    await store.locate(id);

    expect(revisionService.jumpToTransaction).not.toHaveBeenCalled();
    expect(store.working).toBe(true);
  });

  it("事务不存在时静默返回", async () => {
    const store = useEditStore();
    await expect(store.locate("AI_EDIT_999")).resolves.toBeUndefined();
    expect(revisionService.jumpToTransaction).not.toHaveBeenCalled();
  });
});

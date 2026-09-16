/**
 * DiffEngine 单元测试（对应需求文档 §68）。
 */
import { describe, expect, it } from "vitest";
import { diffEngine } from "@/services/diff/DiffEngine";
import type { DiffOperation, EditOperation } from "@/models/EditPlan";

/**
 * 模拟 PatchEngine 的纯文本应用策略：从后向前（offset 降序）应用操作。
 * EditOperation 坐标正确 ⇔ 该函数能从 original 重建出 newText。
 */
function applyEditOperations(original: string, ops: EditOperation[]): string {
  let text = original;
  const sorted = [...ops].sort((a, b) => b.start - a.start || b.end - a.end);
  for (const op of sorted) {
    text = text.slice(0, op.start) + (op.newText ?? "") + text.slice(op.end);
  }
  return text;
}

function types(ops: DiffOperation[]): string {
  return ops.map((o) => o.type[0]).join("");
}

describe("DiffEngine.diff（§20 字符级差异）", () => {
  it("§68 Test 1：传统 → 微服务 得到最小差异", () => {
    const ops = diffEngine.diff("系统采用传统架构。", "系统采用微服务架构。");
    expect(ops).toEqual([
      { type: "equal", text: "系统采用" },
      { type: "delete", text: "传统" },
      { type: "insert", text: "微服务" },
      { type: "equal", text: "架构。" },
    ]);
  });

  it("§68 Test 2：扩写不得退化为整段删除 + 整段插入", () => {
    const oldText = "系统能够处理数据。";
    const newText = "系统能够高效、稳定地处理大规模实时数据。";
    const ops = diffEngine.diff(oldText, newText);

    // 不得删除整段（删除总量远小于原文）
    const deleted = ops
      .filter((o) => o.type === "delete")
      .map((o) => o.text)
      .join("");
    expect(deleted.length).toBeLessThan(oldText.length);
    expect(deleted).not.toBe(oldText);

    // 差异是局部的：原文仍有等文本段被保留
    expect(ops.filter((o) => o.type === "equal").length).toBeGreaterThanOrEqual(1);
    expect(ops.some((o) => o.type === "equal" && o.text.startsWith("系统能够"))).toBe(true);
    expect(ops.some((o) => o.type === "equal" && o.text.includes("数据"))).toBe(true);

    // 插入的内容确实是扩充的文字
    const inserted = ops
      .filter((o) => o.type === "insert")
      .map((o) => o.text)
      .join("");
    expect(newText).toContain(inserted);
    expect(inserted).toContain("高效");
  });

  it("相同文本返回空数组", () => {
    expect(diffEngine.diff("完全一致", "完全一致")).toEqual([]);
  });

  it("行尾规范化后参与比较（\\r\\n / \\r / \\v 等价）", () => {
    const ops = diffEngine.diff("第一段\r\n第二段", "第一段\r第二段");
    expect(ops).toEqual([]);
  });

  it("空原文 → 全部为插入", () => {
    expect(diffEngine.diff("", "全新内容")).toEqual([{ type: "insert", text: "全新内容" }]);
  });

  it("多段落替换产生跨段差异", () => {
    const ops = diffEngine.diff("第一段\n第二段", "第一段\n改动段");
    expect(ops.some((o) => o.type === "delete" && o.text === "第二")).toBe(true);
    expect(ops.some((o) => o.type === "insert" && o.text === "改动")).toBe(true);
    expect(types(ops).startsWith("e")).toBe(true);
  });
});

describe("DiffEngine.toEditOperations（§22 折叠）", () => {
  it("§68 Test 1 折叠为单个 replace", () => {
    const ops = diffEngine.computeEditOperations("系统采用传统架构。", "系统采用微服务架构。");
    expect(ops).toHaveLength(1);
    expect(ops[0]).toMatchObject({
      type: "replace",
      start: 4,
      end: 6,
      oldText: "传统",
      newText: "微服务",
    });
  });

  it("§68 Test 2 折叠后可精确重建（不退化为整段删除 + 整段插入）", () => {
    const oldText = "系统能够处理数据。";
    const newText = "系统能够高效、稳定地处理大规模实时数据。";
    const ops = diffEngine.computeEditOperations(oldText, newText);

    expect(ops.length).toBeGreaterThanOrEqual(1);

    // 不存在覆盖整段原文的纯删除
    const wholeDelete = ops.find((o) => o.type === "delete" && o.start === 0 && o.end === oldText.length);
    expect(wholeDelete).toBeUndefined();

    // 坐标一致性：从后向前应用可精确重建 newText
    expect(applyEditOperations(oldText, ops)).toBe(newText);
  });

  it("句首插入", () => {
    const ops = diffEngine.computeEditOperations("世界", "你好世界");
    expect(ops).toHaveLength(1);
    expect(ops[0]).toMatchObject({ type: "insert", start: 0, end: 0, newText: "你好" });
  });

  it("句尾插入", () => {
    const ops = diffEngine.computeEditOperations("你好", "你好！");
    expect(ops).toHaveLength(1);
    expect(ops[0]).toMatchObject({ type: "insert", start: 2, end: 2, newText: "！" });
  });

  it("整段删除折叠为单个 delete", () => {
    const ops = diffEngine.computeEditOperations("删除整句话", "");
    expect(ops).toHaveLength(1);
    expect(ops[0]).toMatchObject({ type: "delete", start: 0, end: 5, oldText: "删除整句话" });
  });

  it("操作按文档顺序排列（start 升序）", () => {
    const ops = diffEngine.computeEditOperations(
      "one two three four five",
      "ONE two three FOUR five!",
    );
    expect(ops.length).toBeGreaterThanOrEqual(2);
    for (let i = 1; i < ops.length; i += 1) {
      expect(ops[i].start).toBeGreaterThanOrEqual(ops[i - 1].start);
    }
  });
});

describe("EditOperation 坐标一致性（PatchEngine 文本应用模拟）", () => {
  const CASES: Array<[string, string]> = [
    ["系统采用传统架构。", "系统采用微服务架构。"],
    ["系统能够处理数据。", "系统能够高效、稳定地处理大规模实时数据。"],
    ["one two three four five", "ONE two three FOUR five!"],
    ["The quick brown fox", "The slow brown dog"],
    ["", "从零开始的内容"],
    ["全部删除", ""],
    ["重复词，重复词，重复词。", "重复词，去重词，重复词。"],
    ["第一段\r\n第二段\r\n第三段", "第一段\n修改段\n第三段"],
    ["多个  空格", "多个空格"],
    ["Hello, 世界! Mixed content.", "你好, 世界! Mixed edits."],
  ];

  it.each(CASES)("从后向前应用可精确重建 newText：%j", (oldText, newText) => {
    const ops = diffEngine.computeEditOperations(oldText, newText);
    if (oldText === newText) {
      expect(ops).toHaveLength(0);
      return;
    }
    const normalizedOld = oldText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\v/g, "\n");
    expect(applyEditOperations(normalizedOld, ops)).toBe(
      newText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\v/g, "\n"),
    );
  });

  it("多操作互不干扰（从后向前应用顺序正确）", () => {
    const oldText = "aaa bbb ccc ddd eee";
    const newText = "AAA bbb CcC ddd eee";
    const ops = diffEngine.computeEditOperations(oldText, newText);
    const replaceCount = ops.filter((o) => o.type === "replace").length;
    expect(replaceCount).toBeGreaterThanOrEqual(2);
    expect(applyEditOperations(oldText, ops)).toBe(newText);
  });
});

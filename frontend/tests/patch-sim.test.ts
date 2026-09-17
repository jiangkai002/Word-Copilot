/**
 * PatchEngine 锚定策略回归测试（纯字符串模拟，不依赖 Word）。
 *
 * 背景：真实润色场景曾出现“插入@0 的后方锚点取原文 [0,20) 窗口，
 * 但从后向前应用时窗口内位置已被替换操作改动，锚文本不存在 → 插入被跳过
 * → 校验失败回滚”。insertAnchorWindow 将窗口截止到下一操作边界修复该问题。
 * 本文件用当时的真实操作序列钉住该回归。
 */
import { describe, expect, it } from "vitest";
import { countOccurrencesBefore, insertAnchorWindow } from "@/utils/match";
import type { EditOperation } from "@/models/EditPlan";

const ANCHOR_LENGTH = 20;

/** 非重叠第 N 次（0 起）出现位置；不存在返回 -1 —— 镜像 Word search 语义 */
function nthIndexOf(hay: string, needle: string, n: number): number {
  if (!needle) return -1;
  let i = 0;
  for (let k = 0; k <= n; k++) {
    const found = hay.indexOf(needle, i);
    if (found === -1) return -1;
    if (k === n) return found;
    i = found + needle.length;
  }
  return -1;
}

/**
 * 模拟 PatchEngine.applyEditPlan 的从后向前应用：
 * - replace/delete：搜索 oldText 第 N 个匹配（N = 原文中 op.start 之前的出现次数）
 * - insert：insertAnchorWindow 取锚点，Before/After 插入
 * 返回最终文本与被跳过的操作。
 */
function simulateApply(original: string, ops: EditOperation[]): { final: string; skipped: EditOperation[] } {
  let text = original;
  const skipped: EditOperation[] = [];
  const sorted = [...ops].sort((a, b) => b.start - a.start || b.end - a.end);

  for (const op of sorted) {
    if (op.type === "delete" || op.type === "replace") {
      const anchor = op.oldText ?? "";
      const occurrence = countOccurrencesBefore(original, anchor, op.start);
      const idx = nthIndexOf(text, anchor, occurrence);
      if (idx < 0) {
        skipped.push(op);
        continue;
      }
      text = text.slice(0, idx) + (op.newText ?? "") + text.slice(idx + anchor.length);
      continue;
    }

    // insert
    const { following, preceding } = insertAnchorWindow(original, op, ops, ANCHOR_LENGTH);
    if (following.length >= 1) {
      const occurrence = countOccurrencesBefore(original, following, op.start);
      const idx = nthIndexOf(text, following, occurrence);
      if (idx >= 0) {
        text = text.slice(0, idx) + (op.newText ?? "") + text.slice(idx);
        continue;
      }
    }
    if (preceding.length >= 1) {
      const occurrence = Math.max(0, countOccurrencesBefore(original, preceding, op.start) - 1);
      const idx = nthIndexOf(text, preceding, occurrence);
      if (idx >= 0) {
        const at = idx + preceding.length;
        text = text.slice(0, at) + (op.newText ?? "") + text.slice(at);
        continue;
      }
    }
    skipped.push(op);
  }
  return { final: text, skipped };
}

/** 2026-09-15 真实润色场景（deepseek）的 original / new 与 DiffEngine 输出 */
const REAL_ORIGINAL = "系统采用传统架构，可以处理大量数据。";
const REAL_NEW = "本系统采用传统架构，具备处理海量数据的能力。";
const REAL_OPS: EditOperation[] = [
  { type: "insert", start: 0, end: 0, oldText: "", newText: "本" },
  { type: "replace", start: 9, end: 11, oldText: "可以", newText: "具备" },
  { type: "replace", start: 13, end: 14, oldText: "大", newText: "海" },
  { type: "insert", start: 17, end: 17, oldText: "", newText: "的能力" },
];

describe("insertAnchorWindow（插入锚点窗口）", () => {
  it("插入@0 的后方窗口截止到下一操作边界（@9），不横跨已改动区域", () => {
    const op = REAL_OPS[0];
    const { following, preceding } = insertAnchorWindow(REAL_ORIGINAL, op, REAL_OPS, ANCHOR_LENGTH);
    expect(following).toBe("系统采用传统架构，"); // [0, 9)，不含已被 replace 的“可以/大”
    expect(preceding).toBe(""); // 插入点在段首，无前方文本
  });

  it("插入@17（段尾）后方窗口止于句号，回退到前方锚点", () => {
    const op = REAL_OPS[3];
    const { following, preceding } = insertAnchorWindow(REAL_ORIGINAL, op, REAL_OPS, ANCHOR_LENGTH);
    expect(following).toBe("。"); // [17, 18)
    expect(preceding.length).toBeGreaterThan(2);
  });

  it("旧实现的 20 字符窗口在应用中途确实不存在（钉住 bug 成因）", () => {
    // 模拟 @13、@9 已应用后的文本：旧锚点（原文 [0,20)）已不可搜到
    let evolved = REAL_ORIGINAL;
    const sorted = [...REAL_OPS].sort((a, b) => b.start - a.start || b.end - a.end);
    for (const op of sorted.slice(0, 3)) {
      // 前 3 个（按时间）= 插入@17、替换@13、替换@9
      if (op.type === "replace") {
        evolved = evolved.replace(op.oldText ?? "", op.newText ?? "");
      }
    }
    const staleWindow = REAL_ORIGINAL.slice(0, 20);
    expect(evolved.includes(staleWindow)).toBe(false); // ← 旧实现在此失败
  });
});

describe("PatchEngine 应用模拟（真实润色操作序列）", () => {
  it("从后向前应用 4 个操作：全部命中，最终文本 == new_text", () => {
    const { final, skipped } = simulateApply(REAL_ORIGINAL, REAL_OPS);
    expect(skipped).toEqual([]);
    expect(final).toBe(REAL_NEW);
  });

  it("多段重复词场景：连续两处替换 + 段尾插入", () => {
    const original = "这个词这个词这个词出现了三次这个词。";
    const ops: EditOperation[] = [
      { type: "replace", start: 0, end: 3, oldText: "这个词", newText: "单词A" },
      { type: "replace", start: 6, end: 9, oldText: "这个词", newText: "单词B" },
      { type: "insert", start: 18, end: 18, oldText: "", newText: "共计" },
    ];
    // 第 2 处“这个词”（[3,6)）未被替换；插入@18 落在句号之后
    const { final, skipped } = simulateApply(original, ops);
    expect(skipped).toEqual([]);
    expect(final).toBe("单词A这个词单词B出现了三次这个词。共计");
  });

  it("相邻操作只留下单字符锚点时仍能应用插入", () => {
    const original = "步骤1.2 统一BIM坐标系";
    const ops: EditOperation[] = [
      { type: "insert", start: 0, end: 0, oldText: "", newText: "（一）" },
      { type: "replace", start: 1, end: 2, oldText: "骤", newText: "阶段" },
    ];

    const { final, skipped } = simulateApply(original, ops);
    expect(skipped).toEqual([]);
    expect(final).toBe("（一）步阶段1.2 统一BIM坐标系");
  });
});

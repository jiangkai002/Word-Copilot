/**
 * 出现次数统计单元测试 —— PatchEngine 第 N 个匹配锚定的正确性基础。
 * 语义必须与 Word search（从左到右、非重叠）一致。
 */
import { describe, expect, it } from "vitest";
import { countOccurrences, countOccurrencesBefore } from "@/utils/match";

describe("countOccurrences（非重叠、从左到右）", () => {
  it("基础计数", () => {
    expect(countOccurrences("abcabcabc", "abc")).toBe(3);
  });

  it("重叠模式只按非重叠计（aaaa 中 aa = 2 次）", () => {
    expect(countOccurrences("aaaa", "aa")).toBe(2);
  });

  it("中文重复词", () => {
    expect(countOccurrences("重复词，重复词，重复词。", "重复词")).toBe(3);
  });

  it("不含出现 0 次", () => {
    expect(countOccurrences("abcdef", "xyz")).toBe(0);
  });

  it("空 needle 返回 0", () => {
    expect(countOccurrences("abc", "")).toBe(0);
  });
});

describe("countOccurrencesBefore（[0, limit) 内完整落入的匹配）", () => {
  it("limit 截断：匹配必须完整落在区间内", () => {
    // "abc" 在 [0,4) 内只有第 1 个完整出现（位置 0）；位置 3 的出现跨越边界
    expect(countOccurrencesBefore("abcabcabc", "abc", 4)).toBe(1);
    expect(countOccurrencesBefore("abcabcabc", "abc", 6)).toBe(2);
    expect(countOccurrencesBefore("abcabcabc", "abc", 9)).toBe(3);
  });

  it("limit 小于 needle 长度时为 0", () => {
    expect(countOccurrencesBefore("abc", "abc", 2)).toBe(0);
  });

  it("恰好等于 needle 长度时为 1", () => {
    expect(countOccurrencesBefore("abc", "abc", 3)).toBe(1);
  });

  it("空 needle 返回 0", () => {
    expect(countOccurrencesBefore("abc", "", 3)).toBe(0);
  });

  it("PatchEngine 场景：原文中第 N 次出现的序号计算", () => {
    // 原文 “好词好词好词”，要删除第 2 个 “好词”（op.start = 4）
    const original = "好词好词好词";
    const occurrence = countOccurrencesBefore(original, "好词", 4);
    expect(occurrence).toBe(2); // 第 3 个匹配（0 基序号 2）是目标
  });

  it("跨段文本不受 \\n 影响", () => {
    const original = "词\n词\n词";
    // “词”位于 0、2、4；[0,5) 内全部完整落入
    expect(countOccurrencesBefore(original, "词", 5)).toBe(3);
    // [0,4) 内位置 4 的匹配跨越边界，不计入
    expect(countOccurrencesBefore(original, "词", 4)).toBe(2);
  });
});

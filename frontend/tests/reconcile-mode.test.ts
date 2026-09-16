/**
 * resolveReconcileMode 单元测试（§29 对账模式判定）：
 * text 恒 count；非 text 仅在 apply 时测得修订 > 0 才信任 count 信号，
 * 否则 existence（CC 存在即 pending —— 格式未跟踪 / 插入未跟踪场景）。
 */
import { describe, expect, it } from "vitest";
import { resolveReconcileMode } from "@/services/word/RevisionService";
import type { TransactionKind } from "@/models/EditTransaction";

describe("resolveReconcileMode", () => {
  it("text 恒为 count（修订条数是唯一信号，与测得数量无关）", () => {
    expect(resolveReconcileMode("text", 0)).toBe("count");
    expect(resolveReconcileMode("text", 5)).toBe("count");
    expect(resolveReconcileMode("text", null)).toBe("count");
    expect(resolveReconcileMode("text", undefined)).toBe("count");
  });

  it("format 且 apply 时测得修订 > 0 → count", () => {
    expect(resolveReconcileMode("format", 1)).toBe("count");
    expect(resolveReconcileMode("format", 12)).toBe("count");
  });

  it("format 且修订为 0（Word 常不记录格式修订）→ existence", () => {
    expect(resolveReconcileMode("format", 0)).toBe("existence");
  });

  it("insert-* 且修订为 0 / null / undefined（未跟踪插入）→ existence", () => {
    expect(resolveReconcileMode("insert-table", 0)).toBe("existence");
    expect(resolveReconcileMode("insert-formula", 0)).toBe("existence");
    expect(resolveReconcileMode("insert-paragraph", 0)).toBe("existence");
    expect(resolveReconcileMode("insert-formula", null)).toBe("existence");
    expect(resolveReconcileMode("insert-table", undefined)).toBe("existence");
  });

  it("insert-* 且 apply 时测得修订 > 0 → count（插入被跟踪，条数可信）", () => {
    expect(resolveReconcileMode("insert-table", 3)).toBe("count");
    expect(resolveReconcileMode("insert-formula", 2)).toBe("count");
    expect(resolveReconcileMode("insert-paragraph", 1)).toBe("count");
  });

  it("全 kind 组合的矩阵", () => {
    const kinds: TransactionKind[] = ["text", "format", "insert-table", "insert-formula", "insert-paragraph"];
    for (const kind of kinds) {
      expect(resolveReconcileMode(kind, 7)).toBe("count");
    }
    for (const kind of kinds.slice(1)) {
      expect(resolveReconcileMode(kind, 0)).toBe("existence");
    }
  });
});

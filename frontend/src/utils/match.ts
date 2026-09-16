/**
 * 非重叠出现次数统计。
 *
 * PatchEngine 用它在“原文”里推算某次操作对应的搜索匹配序号，
 * 语义必须与 Word search（从左到右、非重叠）一致。
 */
import type { EditOperation } from "@/models/EditPlan";

/**
 * 统计 needle 在 hay[0, limit) 中非重叠出现的次数。
 * 匹配必须完全落在 [0, limit) 内。
 */
export function countOccurrencesBefore(hay: string, needle: string, limit: number): number {
  if (!needle) return 0;
  const end = Math.min(limit, hay.length);
  if (needle.length > end) return 0;
  let count = 0;
  let i = 0;
  while (i <= end - needle.length) {
    const found = hay.indexOf(needle, i);
    if (found === -1 || found + needle.length > end) break;
    count += 1;
    i = found + needle.length;
  }
  return count;
}

/** 统计 needle 在 hay 中非重叠出现的总次数 */
export function countOccurrences(hay: string, needle: string): number {
  return countOccurrencesBefore(hay, needle, hay.length);
}

/**
 * 插入操作的锚点窗口（PatchEngine §21 锚定策略的纯函数部分）。
 *
 * 插入点 p 的“后方锚点”窗口取原文 [p, p+len)，
 * 但必须截止到 p 之后最近一个操作的边界 —— 从后向前应用时，
 * 那些操作已经改动过文档，窗口若横跨它们，锚文本在文档中已不存在，
 * 会导致插入被跳过（曾导致真实润色场景校验失败回滚）。
 * “前方锚点” [p-len, p) 中的操作均未应用，天然安全，仅按长度截取。
 *
 * 返回的窗口均为单行（不含 \n），与 Word search 的可搜索形式一致。
 */
export function insertAnchorWindow(
  original: string,
  op: EditOperation,
  allOps: readonly EditOperation[],
  anchorLength: number,
): { following: string; preceding: string } {
  // p 之后最近一个其他操作的起点（原文坐标）
  let boundary = Infinity;
  for (const other of allOps) {
    if (other === op || other.start <= op.start) continue;
    if (other.start < boundary) boundary = other.start;
  }
  const followingMax = Math.min(op.start + anchorLength, boundary);
  const following = original.slice(op.start, followingMax).split("\n")[0];

  const preceding = original
    .slice(Math.max(0, op.start - anchorLength), op.start)
    .split("\n")
    .pop() ?? "";
  return { following, preceding };
}

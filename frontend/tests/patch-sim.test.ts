/**
 * PatchEngine 锚定策略回归测试（纯字符串模拟，不依赖 Word）。
 *
 * 背景：真实润色场景曾出现“插入@0 的后方锚点取原文 [0,20) 窗口，
 * 但从后向前应用时窗口内位置已被替换操作改动，锚文本不存在 → 插入被跳过
 * → 校验失败回滚”。insertAnchorWindow 将窗口截止到下一操作边界修复该问题。
 *
 * 2026-09-17 回归（真实 Word 调用失败，见 backend/logs/conversations.jsonl）：
 * 专利文档段落内含行内公式（OMML）。Range.text 会把公式线性化成普通文本
 * （Diff 坐标系因此成立），但 Word search 只作用于纯文本 run —— 横跨公式的
 * oldText / 锚点窗口搜不到 → 操作被跳过 → 校验失败回滚。修复：
 * - delete/replace：前后缀分别锚定 + expandTo 覆盖中间公式 + 线性文本校验
 * - insert：锚点窗口按减半长度收缩为可搜索前缀/后缀，插入位置不变
 * 本文件用该次失败的真实操作序列钉住回归（simulateApply 以 mathZones
 * 模拟公式区间：区间内的文本不产生 search 匹配）。
 */
import { describe, expect, it } from "vitest";
import { countOccurrencesBefore, insertAnchorWindow } from "@/utils/match";
import { diffEngine } from "@/services/diff/DiffEngine";
import type { EditOperation } from "@/models/EditPlan";

const ANCHOR_LENGTH = 20;

/** 公式区间（OMML math zone，当前文本坐标）：其中的文本不可被 search 命中 */
interface MathZone {
  start: number;
  end: number;
}

/** 模拟 Word search：非重叠扫描，跳过与公式区间重叠的匹配，返回全部可见匹配位置 */
function searchableMatches(hay: string, needle: string, zones: readonly MathZone[]): number[] {
  if (!needle) return [];
  const out: number[] = [];
  let i = 0;
  while (i <= hay.length - needle.length) {
    const found = hay.indexOf(needle, i);
    if (found === -1) return out;
    const overlaps = zones.some((z) => found < z.end && found + needle.length > z.start);
    if (overlaps) {
      i = found + 1; // 公式内的文本不产生匹配
      continue;
    }
    out.push(found);
    i = found + needle.length; // 非重叠
  }
  return out;
}

/**
 * 模拟 PatchEngine.applyEditPlan 的从后向前应用：
 * - replace/delete：搜索 oldText 第 N 个匹配（N = 原文中 op.start 之前的出现次数）；
 *   搜不到（横跨公式）时镜像 locateByAnchoredEnds —— 前后缀锚定 + 区间合并 + 线性文本校验
 * - insert：insertAnchorWindow 取锚点，Before/After 插入；
 *   锚点搜不到（横跨公式）时镜像 pickInsertAnchor —— 按减半长度收缩重试
 * 返回最终文本与被跳过的操作。
 */
function simulateApply(
  original: string,
  ops: EditOperation[],
  mathZones: MathZone[] = [],
): { final: string; skipped: EditOperation[] } {
  let text = original;
  let zones = mathZones.map((z) => ({ ...z }));
  const skipped: EditOperation[] = [];
  const sorted = [...ops].sort((a, b) => b.start - a.start || b.end - a.end);

  /** 第 occurrence 个（0 起）可搜索匹配位置；不存在返回 -1 */
  const searchAt = (needle: string, occurrence: number): number => {
    const matches = searchableMatches(text, needle, zones);
    return occurrence >= 0 && occurrence < matches.length ? matches[occurrence] : -1;
  };

  /** 在 [from, to) 应用文本手术并同步公式区间坐标 */
  const splice = (from: number, to: number, replacement: string): void => {
    text = text.slice(0, from) + replacement + text.slice(to);
    const delta = replacement.length - (to - from);
    zones = zones
      .filter((z) => !(z.start >= from && z.end <= to)) // 区间内的公式随删除消失
      .map((z) => (z.start >= to ? { start: z.start + delta, end: z.end + delta } : z));
  };

  /** 镜像 PatchEngine.locateByAnchoredEnds：前后缀锚定 + expandTo + 线性文本校验 */
  const locateByEnds = (op: EditOperation): { start: number; end: number } | null => {
    const oldText = op.oldText ?? "";
    if (oldText.length < 4) return null;
    const pickEnd = (end: "prefix" | "suffix"): number => {
      let len = Math.floor(oldText.length / 2);
      while (len >= 1) {
        const probe = end === "prefix" ? oldText.slice(0, len) : oldText.slice(oldText.length - len);
        const occurrence = end === "prefix"
          ? countOccurrencesBefore(original, probe, op.start)
          : Math.max(0, countOccurrencesBefore(original, probe, op.end) - 1);
        const idx = searchAt(probe, occurrence);
        if (idx >= 0) return end === "prefix" ? idx : idx + probe.length;
        len = Math.floor(len / 2);
      }
      return -1;
    };
    const start = pickEnd("prefix");
    if (start < 0) return null;
    const end = pickEnd("suffix");
    if (end < 0) return null;
    if (text.slice(start, end) !== oldText) return null; // 线性文本校验，防错误配对
    return { start, end };
  };

  /** 镜像 PatchEngine.pickInsertAnchor：锚点窗口减半收缩，插入位置不变 */
  const insertAnchorPos = (window: string, location: "before" | "after", insertAt: number): number => {
    let len = window.length;
    while (len >= 1) {
      const probe = location === "before" ? window.slice(0, len) : window.slice(window.length - len);
      const occurrence = location === "before"
        ? countOccurrencesBefore(original, probe, insertAt)
        : Math.max(0, countOccurrencesBefore(original, probe, insertAt) - 1);
      const idx = searchAt(probe, occurrence);
      if (idx >= 0) return location === "before" ? idx : idx + probe.length;
      len = Math.floor(len / 2);
    }
    return -1;
  };

  for (const op of sorted) {
    if (op.type === "delete" || op.type === "replace") {
      const oldText = op.oldText ?? "";
      const occurrence = countOccurrencesBefore(original, oldText, op.start);
      let from = searchAt(oldText, occurrence);
      let to = from >= 0 ? from + oldText.length : -1;
      if (from < 0) {
        const ends = locateByEnds(op);
        if (ends) {
          from = ends.start;
          to = ends.end;
        }
      }
      if (from < 0) {
        skipped.push(op);
        continue;
      }
      splice(from, to, op.newText ?? "");
      continue;
    }

    // insert
    const { following, preceding } = insertAnchorWindow(original, op, ops, ANCHOR_LENGTH);
    if (following.length >= 1) {
      const at = insertAnchorPos(following, "before", op.start);
      if (at >= 0) {
        splice(at, at, op.newText ?? "");
        continue;
      }
    }
    if (preceding.length >= 1) {
      const at = insertAnchorPos(preceding, "after", op.start);
      if (at >= 0) {
        splice(at, at, op.newText ?? "");
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

/**
 * 2026-09-17 真实失败场景（专利文档表格内段落，agent propose_edit）：
 * original/new 取自 conversations.jsonl 中该次调用的提案原文。
 * 段内两处公式（原文坐标）：[53,62) = “T_cam^BIM”、[106,121) = “T_cam^BIM∈SE(3)”。
 * 删除操作 oldText 横跨 [106,121)，Word search 无法整串命中（本次失败根因）。
 */
const PATENT_ORIGINAL =
  "步骤1.2 统一BIM坐标系、巡检定位坐标系和长度单位，确定坐标轴方向、楼层高程和原点转换关系。以齐次变换T_cam^BIM表示相机坐标系到BIM坐标系的变换，使相机局部点和方向能够映射到建筑世界坐标。以齐次变换T_cam^BIM∈SE(3)表示由相机坐标系到BIM坐标系的坐标变换。约定上标BIM表示目标坐标系、下标cam表示源坐标系，T_cam^BIM=[R|t;0 1]，其中R为旋转矩阵、t为平移向量。对相机坐标系中的点p_c，其在BIM坐标系中的坐标为p_BIM=R·p_c+t（齐次形式取w=1）；对相机坐标系中的方向d_c，其在BIM坐标系中的方向为d_BIM=R·d_c（齐次形式取w=0），即方向仅受旋转作用、不受平移影响。对第i个相机，记T_cam,i^BIM=[R_i|C_i]，其中R_i为相机在BIM坐标系中的旋转，C_i为相机中心在BIM坐标系中的位置。";
const PATENT_NEW =
  "步骤1.2 统一BIM坐标系、巡检定位坐标系和长度单位，确定坐标轴方向、楼层高程和原点转换关系。以齐次变换T_cam^BIM∈SE(3)表示相机坐标系到BIM坐标系的坐标变换，使相机局部点和方向能够映射到建筑世界坐标。约定上标BIM表示目标坐标系、下标cam表示源坐标系，T_cam^BIM=[R|t;0 1]，其中R为旋转矩阵、t为平移向量。对相机坐标系中的点p_c，其在BIM坐标系中的坐标为p_BIM=R·p_c+t（齐次形式取w=1）；对相机坐标系中的方向d_c，其在BIM坐标系中的方向为d_BIM=R·d_c（齐次形式取w=0），即方向仅受旋转作用、不受平移影响。对第i个相机，记T_cam,i^BIM=[R_i|C_i]，其中R_i为相机在BIM坐标系中的旋转，C_i为相机中心在BIM坐标系中的位置。";
const PATENT_OPS: EditOperation[] = [
  { type: "insert", start: 62, end: 62, oldText: "", newText: "∈SE(3)" },
  { type: "insert", start: 77, end: 77, oldText: "", newText: "坐标" },
  {
    type: "delete",
    start: 101,
    end: 142,
    oldText: "以齐次变换T_cam^BIM∈SE(3)表示由相机坐标系到BIM坐标系的坐标变换。",
    newText: "",
  },
];
const PATENT_MATH_ZONES: MathZone[] = [
  { start: 53, end: 62 },
  { start: 106, end: 121 },
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

describe("公式（OMML math zone）场景 —— 2026-09-17 真实失败回归", () => {
  it("DiffEngine 对该 original/new 的输出与失败现场的操作序列一致", () => {
    expect(diffEngine.computeEditOperations(PATENT_ORIGINAL, PATENT_NEW)).toEqual(PATENT_OPS);
  });

  it("横跨公式的 oldText 无法被 search 整串命中（钉住根因）", () => {
    const deleteOp = PATENT_OPS[2];
    const matches = searchableMatches(PATENT_ORIGINAL, deleteOp.oldText ?? "", PATENT_MATH_ZONES);
    expect(matches).toEqual([]); // ← 修复前：操作被跳过，校验失败回滚
  });

  it("横跨公式的删除经前后缀锚点兜底应用：3 个操作全部生效", () => {
    const { final, skipped } = simulateApply(PATENT_ORIGINAL, PATENT_OPS, PATENT_MATH_ZONES);
    expect(skipped).toEqual([]);
    expect(final).toBe(PATENT_NEW);
  });

  it("插入锚点窗口横跨公式时收缩为可搜索前缀，插入位置不变", () => {
    const original = "甲乙丙丁戊己庚辛壬癸子丑寅卯辰巳午未申酉";
    const zones: MathZone[] = [{ start: 5, end: 12 }]; // [5,12) 视为公式
    const ops: EditOperation[] = [
      { type: "insert", start: 0, end: 0, oldText: "", newText: "【新】" },
    ];
    // 后方窗口 = 原文 [0,20)，横跨公式 → 完整搜索失败 → 收缩到 [0,5)
    const { final, skipped } = simulateApply(original, ops, zones);
    expect(skipped).toEqual([]);
    expect(final).toBe("【新】" + original);
  });

  it("无公式时兜底逻辑不改变既有行为（专利场景去掉公式区间）", () => {
    const { final, skipped } = simulateApply(PATENT_ORIGINAL, PATENT_OPS, []);
    expect(skipped).toEqual([]);
    expect(final).toBe(PATENT_NEW);
  });
});

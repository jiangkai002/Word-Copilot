/**
 * DiffEngine（对应需求文档 §19–§22）：
 * 后端只返回 original_text / new_text，字符级 Diff 必须在 Word 客户端完成。
 *
 * 使用 diff-match-patch + diff_cleanupSemantic，
 * 输出最小化的 DiffOperation 序列，并折叠为 EditOperation（insert/delete/replace）。
 *
 * 坐标系：normalizeText 后的“规范化原文”（\n 表示段落分隔）。
 * 应用到 Word 时由 PatchEngine 负责把 \n 转回 \r。
 */
import * as dmpModule from "diff-match-patch";
import type { DiffOperation, EditOperation } from "@/models/EditPlan";
import { normalizeText } from "@/utils/text";

/** diff-match-patch 的运行时形态（CJS 包，导入形态做防御式处理） */
type DmpLike = {
  diff_main(text1: string, text2: string): Array<[number, string]>;
  diff_cleanupSemantic(diffs: Array<[number, string]>): void;
  /** diff 超时（秒），超时返回次优结果而不是失败 */
  Diff_Timeout: number;
};
type DmpCtor = new () => DmpLike;

function resolveDmpCtor(): DmpCtor {
  const mod = dmpModule as unknown as Record<string, unknown>;
  const candidate = mod.diff_match_patch ?? mod.default ?? mod;
  if (typeof candidate !== "function") {
    throw new Error("无法加载 diff-match-patch");
  }
  return candidate as DmpCtor;
}

const createDmp = (() => {
  let ctor: DmpCtor | null = null;
  return () => {
    if (!ctor) ctor = resolveDmpCtor();
    const dmp = new ctor();
    return dmp;
  };
})();

export class DiffEngine {
  /**
   * 字符级 Diff（§20）。输入会先规范化（\r\n / \r / \v → \n）。
   */
  diff(oldText: string, newText: string): DiffOperation[] {
    const oldNorm = normalizeText(oldText);
    const newNorm = normalizeText(newText);
    if (oldNorm === newNorm) return [];

    const dmp = createDmp();
    dmp.Diff_Timeout = 5; // 秒；超时返回次优结果而不是失败
    const diffs = dmp.diff_main(oldNorm, newNorm);
    dmp.diff_cleanupSemantic(diffs);

    return diffs
      .filter(([, text]) => text.length > 0)
      .map(([op, text]) => ({
        type: op === 0 ? ("equal" as const) : op === -1 ? ("delete" as const) : ("insert" as const),
        text,
      }));
  }

  /**
   * 由 DiffOperation 折叠出 EditOperation（§22）。
   * 相邻的 delete+insert 合并为 replace；连续 delete 合并；连续 insert 合并。
   * 返回操作按文档顺序排列（start 升序）。
   */
  toEditOperations(operations: DiffOperation[]): EditOperation[] {
    const result: EditOperation[] = [];
    let pos = 0; // 规范化原文中的游标
    let pendingDelete: { start: number; end: number; text: string } | null = null;
    let pendingInsert: string | null = null;

    const flush = () => {
      if (pendingDelete && pendingInsert !== null) {
        result.push({
          type: "replace",
          start: pendingDelete.start,
          end: pendingDelete.end,
          oldText: pendingDelete.text,
          newText: pendingInsert,
        });
      } else if (pendingDelete) {
        result.push({
          type: "delete",
          start: pendingDelete.start,
          end: pendingDelete.end,
          oldText: pendingDelete.text,
          newText: "",
        });
      } else if (pendingInsert !== null) {
        result.push({
          type: "insert",
          // 插入发生在当前游标处（其前的 equal/delete 已推进 pos，其后的 equal 尚未推进）
          start: pos,
          end: pos,
          oldText: "",
          newText: pendingInsert,
        });
      }
      pendingDelete = null;
      pendingInsert = null;
    };

    for (const op of operations) {
      if (op.type === "equal") {
        flush();
        pos += op.text.length;
      } else if (op.type === "delete") {
        if (!pendingDelete) {
          pendingDelete = { start: pos, end: pos + op.text.length, text: op.text };
        } else {
          pendingDelete.end += op.text.length;
          pendingDelete.text += op.text;
        }
        pos += op.text.length;
      } else {
        // insert
        pendingInsert = pendingInsert === null ? op.text : pendingInsert + op.text;
      }
    }
    flush();
    return result;
  }

  /**
   * 一步到位：oldText → newText 的 EditOperation 列表。
   */
  computeEditOperations(oldText: string, newText: string): EditOperation[] {
    return this.toEditOperations(this.diff(oldText, newText));
  }
}

export const diffEngine = new DiffEngine();

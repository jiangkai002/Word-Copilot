/**
 * RevisionService / RevisionManager（对应需求文档 §26–§29 / §54 / §55）：
 *
 * Accept / Reject 只作用于事务对应的 ContentControl 范围内的修订，
 * 绝不调用 document 级别的 acceptAll —— 这是强制要求（§55 AI 修订与人工修订隔离）。
 *
 * reconcile（§29）：用户可能用 Word 原生审阅处理了修订，
 * 刷新时发现控件内 0 条修订即视为已处理，清理临时 ContentControl，不报错。
 *
 * 按 kind 分派（Agent 格式 / 插入能力引入的差异）：
 * - text：修订条数是可信信号（现状语义）。
 * - format：Word 可能不把格式修改记录为修订（count=0 是常态）——
 *   count=0 的拒绝改为「回写应用前原值」还原；对账走 existence 模式
 *   （CC 存在即 pending，不误判 handled_externally）。
 * - insert-table / insert-paragraph：未被跟踪的插入（count=0）拒绝时
 *   delete(false) 连壳带内容删除，保证内容一定被移除。
 * - insert-formula 与段落同样通过完整 Flat OPC 进入修订管线。
 * 非文本事务仅在 apply 时测得修订数 > 0 的情况下才信任 count 信号。
 */
import type { FormatBeforeValues } from "@/models/EditPlan";
import type { TransactionKind } from "@/models/EditTransaction";
import { toCopilotError } from "@/utils/errors";
import { logger } from "@/utils/logger";
import { contentControlService } from "./ContentControlService";
import { formatEngine } from "./FormatEngine";
import { WordService } from "./WordService";

export interface TrackedChangeInfo {
  type: string;
  text: string;
}

export type TransactionOutcome = "accepted" | "rejected" | "handled_externally" | "not_found" | "pending";

/** 事务引用：RevisionService 按此分派接受 / 拒绝 / 对账语义 */
export interface TransactionRef {
  tag: string;
  kind: TransactionKind;
  /** apply 时测得的修订条数（tx.changeCount）—— >0 时 count 信号对非 text 事务可信 */
  appliedChangeCount?: number | null;
  /** kind === "format"：无修订记录路径拒绝时回写的应用前原值 */
  formatBefore?: FormatBeforeValues;
}

export type ReconcileMode = "count" | "existence";

/**
 * 对账模式判定（纯函数）：
 * - text → count（修订条数是唯一信号，现状语义）
 * - 非 text 且 apply 时测得修订 > 0 → count（修订被跟踪，条数可信）
 * - 其余 → existence（CC 存在即 pending —— 格式未跟踪 / 插入未跟踪场景）
 */
export function resolveReconcileMode(
  kind: TransactionKind,
  appliedChangeCount: number | null | undefined,
): ReconcileMode {
  if (kind === "text") return "count";
  return (appliedChangeCount ?? 0) > 0 ? "count" : "existence";
}

export class RevisionService {
  /**
   * 接受事务（§78）：ContentControl.getTrackedChanges().acceptAll()，
   * 然后删除临时 ContentControl，保留正文。
   *
   * format / 未跟踪插入：值已直接生效（count=0）—— 删壳即完成接受。
   */
  async acceptTransaction(ref: TransactionRef): Promise<TransactionOutcome> {
    return WordService.run(async (ctx) => {
      const control = await contentControlService.find(ctx, ref.tag);
      if (!control) return "not_found";

      const changeCount = await this.countChanges(ctx, control);
      if (changeCount === 0) {
        if (ref.kind !== "text" && (ref.appliedChangeCount ?? 0) > 0) {
          // apply 时有修订、现在没有了 —— 已被 Word 原生审阅处理（§29）
          control.delete(true);
          await ctx.sync();
          return "handled_externally";
        }
        // text：已被 Word 原生处理（§29）；format / 未跟踪插入：值已生效
        control.delete(true);
        await ctx.sync();
        return ref.kind === "text" ? "handled_externally" : "accepted";
      }

      try {
        const changes = control.getTrackedChanges();
        changes.acceptAll();
        control.delete(true);
        await ctx.sync();
        return "accepted";
      } catch (err) {
        throw toCopilotError(err);
      }
    });
  }

  /**
   * 拒绝事务（§79）：恢复原文后移除临时 Content Control。
   *
   * - count > 0（所有 kind）：rejectAll + delete(true / insert:false)
   * - format count=0：关修订 → 回写 formatBefore 原值 → delete(true)
   * - insert-* count=0：delete(false) 连壳带内容删除（兜底清除未跟踪插入）
   * - text count=0：已被 Word 原生处理（§29 现状语义）
   */
  async rejectTransaction(ref: TransactionRef): Promise<TransactionOutcome> {
    return WordService.run(async (ctx) => {
      const control = await contentControlService.find(ctx, ref.tag);
      if (!control) return "not_found";

      const changeCount = await this.countChanges(ctx, control);
      if (changeCount === 0) {
        if (ref.kind === "text") {
          // §29：已被 Word 原生处理（正文已由 Word 还原/保留）
          control.delete(true);
          await ctx.sync();
          return "handled_externally";
        }
        if (ref.kind === "format") {
          // Word 未跟踪格式修订 —— 主动回写应用前原值还原
          try {
            ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            const content = control.getRange(Word.RangeLocation.content);
            if (ref.formatBefore && Object.keys(ref.formatBefore).length > 0) {
              await formatEngine.writeBackBeforeValues(ctx, content, ref.formatBefore);
            } else {
              logger.warn(`${ref.tag} 缺少格式原值快照，拒绝后无法自动还原（可 Ctrl+Z 撤销）`);
            }
          } catch (err) {
            logger.error("回写格式原值失败（可 Ctrl+Z 撤销）：", toCopilotError(err).message);
          }
          control.delete(true);
          await ctx.sync();
          return "rejected";
        }
        // insert-*：未跟踪的插入 —— 连壳带内容删除
        control.delete(false);
        await ctx.sync();
        return "rejected";
      }

      try {
        const changes = control.getTrackedChanges();
        changes.rejectAll();
        // insert 拒绝后 CC 已空，delete(false) 兜底清除残壳 / 未跟踪内容
        control.delete(ref.kind.startsWith("insert-") ? false : true);
        await ctx.sync();
        return "rejected";
      } catch (err) {
        throw toCopilotError(err);
      }
    });
  }

  /**
   * 读取事务范围内的修订列表（§26 getTransactionChanges）。
   */
  async getTransactionChanges(tag: string): Promise<TrackedChangeInfo[]> {
    return WordService.run(async (ctx) => {
      const control = await contentControlService.find(ctx, tag);
      if (!control) return [];
      return this.listChanges(ctx, control);
    });
  }

  /**
   * 对账（§29 reconcileTransactions）：检查所有待处理事务。
   * 返回每个 tag 的最新状态。
   *
   * count 模式（text / 曾测得修订的非文本）：0 条修订 → 已被 Word 原生处理。
   * existence 模式（未跟踪格式 / 插入）：CC 存在即 pending —— 卡片保持可操作，
   * 由用户显式接受 / 拒绝收尾（count 信号对这些事务不可信）。
   */
  async reconcile(refs: TransactionRef[]): Promise<Map<string, TransactionOutcome>> {
    const results = new Map<string, TransactionOutcome>();
    if (refs.length === 0) return results;
    await WordService.run(async (ctx) => {
      for (const ref of refs) {
        try {
          const control = await contentControlService.find(ctx, ref.tag);
          if (!control) {
            results.set(ref.tag, "not_found");
            continue;
          }
          const mode = resolveReconcileMode(ref.kind, ref.appliedChangeCount);
          if (mode === "existence") {
            results.set(ref.tag, "pending");
            continue;
          }
          const count = await this.countChanges(ctx, control);
          if (count === 0) {
            // 已被 Word 原生接受/拒绝：清理临时控件与事务状态
            control.delete(ref.kind.startsWith("insert-") ? false : true);
            await ctx.sync();
            results.set(ref.tag, "handled_externally");
          } else {
            results.set(ref.tag, "pending");
          }
        } catch (err) {
          logger.warn(`reconcile ${ref.tag} 失败：`, toCopilotError(err).message);
          results.set(ref.tag, "not_found");
        }
      }
    });
    return results;
  }

  // ------------------------------------------------------------------

  private async countChanges(
    ctx: Word.RequestContext,
    control: Word.ContentControl,
  ): Promise<number> {
    try {
      const changes = control.getTrackedChanges();
      changes.load("text, type");
      await ctx.sync();
      return changes.items.length;
    } catch (err) {
      // 已知问题：文档含“移动内容”修订时 getTrackedChanges 可能失败（office-js #5535）
      logger.warn("读取修订列表失败，按 0 条处理：", toCopilotError(err).message);
      return 0;
    }
  }

  private async listChanges(
    ctx: Word.RequestContext,
    control: Word.ContentControl,
  ): Promise<TrackedChangeInfo[]> {
    try {
      const changes = control.getTrackedChanges();
      changes.load("text, type");
      await ctx.sync();
      return changes.items.map((c) => ({
        type: String((c as { type?: string }).type ?? ""),
        text: (c as { text?: string }).text ?? "",
      }));
    } catch (err) {
      logger.warn("读取修订列表失败：", toCopilotError(err).message);
      return [];
    }
  }
}

export const revisionService = new RevisionService();

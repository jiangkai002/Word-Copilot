/**
 * ContentControlService（对应需求文档 §24）：
 * AI 事务边界 —— tag = word_ai_edit:{UUID}。
 * Content Control 的目的不是显示 UI，而是确定一次 AI 修改对应的 Revision 范围。
 *
 * 注意：所有方法在 WordService.run 批次内调用，不跨批次持有 proxy。
 */
import { logger } from "@/utils/logger";

export const EDIT_CONTROL_TAG_PREFIX = "word_ai_edit:";
export const EDIT_CONTROL_TITLE = "Word AI";

export class ContentControlService {
  /** 按 tag 查找事务 Content Control；找不到返回 null */
  async find(
    ctx: Word.RequestContext,
    tag: string,
  ): Promise<Word.ContentControl | null> {
    try {
      const collection = ctx.document.contentControls.getByTag(tag);
      const first = collection.getFirstOrNullObject();
      first.load("tag");
      await ctx.sync();
      if (first.isNullObject) return null;
      return first;
    } catch (err) {
      logger.debug("查找 ContentControl 失败：", err);
      return null;
    }
  }

  /** 是否存在指定 tag 的控件 */
  async exists(ctx: Word.RequestContext, tag: string): Promise<boolean> {
    return (await this.find(ctx, tag)) !== null;
  }

  /**
   * 移除临时 Content Control，但保留正文（§27/§28）。
   */
  async removeKeepingContent(ctx: Word.RequestContext, tag: string): Promise<boolean> {
    const control = await this.find(ctx, tag);
    if (!control) return false;
    control.delete(true); // keepContent = true
    await ctx.sync();
    logger.debug(`已移除事务控件 ${tag}（保留正文）`);
    return true;
  }
}

export const contentControlService = new ContentControlService();

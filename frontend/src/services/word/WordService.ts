/**
 * WordService（对应需求文档 §46）：
 * 所有 Office.js 调用的统一入口，禁止 Vue 组件直接 Word.run(...)。
 *
 * - 封装 Word.run 批处理
 * - WordApi 能力检测（§5：不支持 1.6 时 AI Revision 禁用但不抛异常）
 * - Office.js 错误归一化
 */
import { CopilotError, toCopilotError } from "@/utils/errors";
import { logger } from "@/utils/logger";

export class WordService {
  /** Office.js 是否已加载（浏览器开发预览时为 false） */
  static isOfficeReady(): boolean {
    return typeof Word !== "undefined" && typeof Word.run === "function";
  }

  /** 是否支持指定 WordApi 版本 */
  static isSetSupported(version: "1.1" | "1.3" | "1.4" | "1.6"): boolean {
    try {
      return Boolean(Office.context?.requirements?.isSetSupported?.("WordApi", version));
    } catch {
      return false;
    }
  }

  /**
   * AI 修订能力：需要 WordApi 1.6
   * （changeTrackingMode 为 1.4；TrackedChange / uniqueLocalId 为 1.6）
   */
  static get revisionSupported(): boolean {
    return this.isSetSupported("1.6");
  }

  /** 表格插入能力：Range.insertTable / Table 对象模型为 WordApi 1.3 */
  static get tableInsertSupported(): boolean {
    return this.isSetSupported("1.3");
  }

  /**
   * 统一的 Word.run 包装：回调内的异常统一转为 CopilotError。
   * 注意：proxy 对象（Range / Paragraph / ContentControl）不允许跨出本批次（§47）。
   */
  static async run<T>(batch: (ctx: Word.RequestContext) => Promise<T>): Promise<T> {
    if (!this.isOfficeReady()) {
      throw new CopilotError("UNSUPPORTED_WORD_VERSION", "Office.js 尚未就绪（请在 Word 任务窗格中运行插件）。");
    }
    try {
      return await Word.run(async (ctx: Word.RequestContext) => batch(ctx));
    } catch (err) {
      throw toCopilotError(err);
    }
  }

  /**
   * 若 WordApi 1.6 不可用，返回给用户的降级提示（§5：不抛异常，仅禁用修订功能）。
   */
  static revisionUnsupportedMessage(): string {
    return "当前 Word 版本不支持 AI 原生修订功能，请升级 Microsoft Word。（聊天与问答仍可继续使用。）";
  }

  /**
   * 读取 document.changeTrackingMode；读取失败时返回 null（个别平台/版本问题）。
   * 旧宿主可能返回字符串字面量（"Off" / "TrackAll" / "TrackMineOnly"），统一归一化为枚举。
   */
  static async readChangeTrackingMode(ctx: Word.RequestContext): Promise<Word.ChangeTrackingMode | null> {
    try {
      const doc = ctx.document;
      doc.load("changeTrackingMode");
      await ctx.sync();
      return WordService.normalizeChangeTrackingMode(doc.changeTrackingMode);
    } catch (err) {
      logger.warn("读取 changeTrackingMode 失败：", toCopilotError(err).message);
      return null;
    }
  }

  private static normalizeChangeTrackingMode(
    value: Word.ChangeTrackingMode | string,
  ): Word.ChangeTrackingMode {
    switch (value) {
      case Word.ChangeTrackingMode.off:
      case "Off":
        return Word.ChangeTrackingMode.off;
      case Word.ChangeTrackingMode.trackMineOnly:
      case "TrackMineOnly":
        return Word.ChangeTrackingMode.trackMineOnly;
      default:
        return Word.ChangeTrackingMode.trackAll;
    }
  }
}

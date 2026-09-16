/**
 * Document Store（对应需求文档 §8–§12 / §41）：
 * 上下文模式（@selection / @paragraph / @section / @document）、选区轻量跟踪、
 * 按需构造 DocumentContext（发送时才读取 —— 保证新鲜度并遵循隐私原则 §60）。
 */
import { defineStore } from "pinia";
import { ref, computed } from "vue";
import type { ContextType, DocumentContext } from "@/models/DocumentContext";
import { CONTEXT_TYPE_LABELS } from "@/models/DocumentContext";
import { documentService, extractKeywords } from "@/services/word/DocumentService";
import { selectionService, type CapturedTarget } from "@/services/word/SelectionService";
import { WordService } from "@/services/word/WordService";
import { logger } from "@/utils/logger";
import { toCopilotError } from "@/utils/errors";
import { useChatStore } from "./chat";

const CONTEXT_MODE_STORAGE_KEY = "word-ai-copilot.contextMode";

export const useDocumentStore = defineStore("document", () => {
  const hostReady = ref(false);
  const revisionSupported = ref(false);
  const contextMode = ref<ContextType>("paragraph");
  const selectionAvailable = ref(false);
  const selectionTextPreview = ref("");
  const selectionLength = ref(0);
  const trackingSelection = ref(false);
  const lastError = ref("");

  const contextLabel = computed(() => CONTEXT_TYPE_LABELS[contextMode.value]);

  let selectionDebounceTimer: ReturnType<typeof setTimeout> | null = null;
  let selectionHandlerAdded = false;

  function init(): void {
    hostReady.value = WordService.isOfficeReady();
    revisionSupported.value = hostReady.value && WordService.revisionSupported;
    try {
      const saved = localStorage.getItem(CONTEXT_MODE_STORAGE_KEY);
      if (saved === "selection" || saved === "paragraph" || saved === "section" || saved === "document") {
        contextMode.value = saved;
      }
    } catch {
      /* ignore */
    }
    void refreshSelection();
    startSelectionTracking();
  }

  function setContextMode(mode: ContextType): void {
    contextMode.value = mode;
    try {
      localStorage.setItem(CONTEXT_MODE_STORAGE_KEY, mode);
    } catch {
      /* ignore */
    }
  }

  /** 轻量选区刷新（选区变化事件防抖后调用） */
  async function refreshSelection(): Promise<void> {
    if (!hostReady.value) return;
    try {
      const snapshot = await selectionService.getSelectionSnapshot();
      selectionAvailable.value = snapshot.hasSelection;
      selectionTextPreview.value = snapshot.text.slice(0, 60);
      selectionLength.value = snapshot.text.trim().length;
      lastError.value = "";
    } catch (err) {
      lastError.value = toCopilotError(err).message;
      logger.debug("刷新选区失败：", lastError.value);
    }
  }

  /**
   * 监听 Word 选区变化（common API DocumentSelectionChanged），防抖 500ms。
   */
  function startSelectionTracking(): void {
    if (!hostReady.value || selectionHandlerAdded) return;
    try {
      Office.context?.document?.addHandlerAsync?.(
        Office.EventType.DocumentSelectionChanged,
        () => {
          trackingSelection.value = true;
          if (selectionDebounceTimer) clearTimeout(selectionDebounceTimer);
          selectionDebounceTimer = setTimeout(() => {
            void refreshSelection();
          }, 500);
        },
        (result: Office.AsyncResult<void>) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            logger.debug("选区跟踪不可用：", result.error?.message);
          }
        },
      );
      selectionHandlerAdded = true;
    } catch (err) {
      logger.debug("注册选区监听失败：", err);
    }
  }

  /**
   * 按当前模式构造 DocumentContext（发送 AI 请求时调用）。
   * 失败时返回 null 并写入系统消息（不中断发送）。
   */
  async function buildDocumentContext(userMessage?: string): Promise<DocumentContext | null> {
    if (!hostReady.value) return null;
    const chatStore = useChatStore();
    try {
      let mode = contextMode.value;
      if (mode === "selection" && !selectionAvailable.value) {
        // 选区为空时自动降级为段落模式（§9/§10）
        mode = "paragraph";
      }

      if (mode === "selection" || mode === "paragraph") {
        const paragraphCtx = await selectionService.getParagraphContext();
        let selectionText: string | undefined;
        if (mode === "selection") {
          selectionText = (await selectionService.getSelectionText()).trim();
        }
        return {
          contextType: mode,
          selection: selectionText ? { text: selectionText } : undefined,
          currentParagraph: paragraphCtx.current,
          previousParagraph: paragraphCtx.previous ?? undefined,
          nextParagraph: paragraphCtx.next ?? undefined,
        };
      }

      if (mode === "section") {
        const snapshot = await selectionService.getSelectionSnapshot();
        const section = await documentService.getSectionContext(snapshot.paragraphId);
        const outline = await documentService.getOutline();
        const context: DocumentContext = {
          contextType: "section",
          outline,
          section: section ?? undefined,
        };
        if (!section) {
          chatStore.addSystem("未识别到当前章节（光标上方没有标题），已按文档问答处理。");
          context.contextType = "document";
          context.documentText = (await documentService.getDocumentText()) ?? undefined;
        }
        return context;
      }

      // document 模式（§12）
      const fullText = await documentService.getDocumentText();
      if (fullText !== null) {
        return { contextType: "document", documentText: fullText };
      }
      // 超长：标题结构 + 当前章节 + 相关段落
      const outline = await documentService.getOutline();
      const snapshot = await selectionService.getSelectionSnapshot();
      const section = await documentService.getSectionContext(snapshot.paragraphId);
      const relevant = await documentService.getRelevantParagraphs(extractKeywords(userMessage ?? ""));
      const parts: string[] = [];
      if (section) parts.push(`【当前章节：${section.title}】\n${section.content}`);
      if (relevant.length > 0) parts.push(`【与问题相关的段落】\n${relevant.join("\n")}`);
      return {
        contextType: "document",
        truncated: true,
        outline,
        section: section ?? undefined,
        documentText: parts.join("\n\n") || "（文档过长且未找到相关段落）",
      };
    } catch (err) {
      const ce = toCopilotError(err);
      logger.warn("构造文档上下文失败：", ce.message);
      chatStore.addSystem(`读取文档上下文失败：${ce.message}`);
      return null;
    }
  }

  /**
   * 捕获编辑目标（§9）：优先选区，否则光标段落。
   */
  async function captureEditTarget(): Promise<CapturedTarget> {
    return selectionService.captureEditTarget();
  }

  return {
    hostReady,
    revisionSupported,
    contextMode,
    selectionAvailable,
    selectionTextPreview,
    selectionLength,
    lastError,
    contextLabel,
    init,
    setContextMode,
    refreshSelection,
    buildDocumentContext,
    captureEditTarget,
  };
});

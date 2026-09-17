/**
 * Chat Store（对应需求文档 §16 / §41 / §58）：
 * 多轮对话 + SSE 流式输出 + AbortController 停止生成。
 * 聊天历史仅保存在当前会话（Pinia 内存），关闭 Word 后允许消失（§41）。
 */
import { defineStore } from "pinia";
import { ref, computed } from "vue";
import type { ChatMessage } from "@/models/ChatMessage";
import type { HistoryEntry } from "@/models/Api";
import type { ChatMode } from "@/commands/intent";
import { chatApi } from "@/services/api/ChatApi";
import { uuid } from "@/utils/id";
import { toCopilotError } from "@/utils/errors";
import { logger, setLogConversationId } from "@/utils/logger";
import { useDocumentStore } from "./document";

const MAX_HISTORY = 10;
const MODE_STORAGE_KEY = "wordai.chatMode";

/** 恢复持久化的模式（无效值回退 auto） */
function restoreMode(): ChatMode {
  try {
    const saved = localStorage.getItem(MODE_STORAGE_KEY);
    if (saved === "chat" || saved === "agent" || saved === "auto") return saved;
  } catch {
    /* ignore */
  }
  return "auto";
}

export const useChatStore = defineStore("chat", () => {
  const conversationId = ref(uuid());
  const messages = ref<ChatMessage[]>([]);
  const streaming = ref(false);
  const abortController = ref<AbortController | null>(null);
  /** 输入模式（对话 / 智能 / Agent），仅影响自由输入的路由 */
  const mode = ref<ChatMode>(restoreMode());

  function setMode(next: ChatMode): void {
    mode.value = next;
    try {
      localStorage.setItem(MODE_STORAGE_KEY, next);
    } catch {
      /* ignore */
    }
  }

  const recentHistory = computed<HistoryEntry[]>(() =>
    messages.value
      .filter((m) => (m.type === "user" || m.type === "assistant") && !m.errorCode && m.content.trim())
      .slice(-MAX_HISTORY)
      .map((m) => ({ role: m.type as "user" | "assistant", content: m.content })),
  );

  function pushMessage(message: Omit<ChatMessage, "id" | "createdAt"> & { id?: string }): ChatMessage {
    const full: ChatMessage = {
      id: message.id ?? uuid(),
      type: message.type,
      content: message.content ?? "",
      createdAt: Date.now(),
      streaming: message.streaming,
      errorCode: message.errorCode,
      transactionId: message.transactionId,
    };
    messages.value.push(full);
    // 必须返回响应式代理（而非原始对象）：后续对 content/streaming/errorCode 的修改
    // 只有经过 proxy 的 set trap 才会触发视图更新 —— 直接改原始对象 UI 不会刷新。
    return messages.value[messages.value.length - 1];
  }

  function addSystem(content: string, errorCode?: string): void {
    pushMessage({ type: "system", content, errorCode });
  }

  function addEditMessage(transactionId: string): void {
    pushMessage({ type: "edit", content: "", transactionId });
  }

  /** 发送消息（流式） */
  async function send(text: string): Promise<void> {
    const message = text.trim();
    if (!message || streaming.value) return;
    const documentStore = useDocumentStore();
    setLogConversationId(conversationId.value);

    const context = await documentStore.buildDocumentContext(message);
    // 先取历史再 push —— 历史不含本次用户消息
    const history = recentHistory.value;
    pushMessage({ type: "user", content: message });

    const assistant = pushMessage({ type: "assistant", content: "", streaming: true });
    streaming.value = true;
    const controller = new AbortController();
    abortController.value = controller;

    let hadError = false;
    try {
      await chatApi.streamChat(
        {
          conversation_id: conversationId.value,
          message,
          context,
          history,
        },
        {
          onToken: (t) => {
            assistant.content += t;
          },
          onError: (code, msg) => {
            hadError = true;
            assistant.errorCode = code;
            assistant.content = assistant.content || msg;
          },
          onDone: () => {
            /* 流结束 */
          },
        },
        controller.signal,
      );
    } catch (err) {
      const ce = toCopilotError(err);
      if (ce.code === "ABORTED") {
        assistant.content = assistant.content || "（已停止生成）";
      } else {
        hadError = true;
        assistant.errorCode = ce.code;
        assistant.content = assistant.content || ce.message;
        logger.error("chat 失败：", ce.message);
      }
    } finally {
      assistant.streaming = false;
      streaming.value = false;
      abortController.value = null;
      if (hadError && !assistant.content) {
        assistant.content = "请求失败，请重试。";
      }
    }
  }

  /** 停止生成（§58：必须可用；用户取消不允许产生 Word 修改） */
  function stop(): void {
    abortController.value?.abort();
  }

  function clear(): void {
    if (streaming.value) stop();
    messages.value = [];
    conversationId.value = uuid();
    setLogConversationId(conversationId.value);
  }

  return {
    conversationId,
    messages,
    streaming,
    abortController,
    mode,
    setMode,
    recentHistory,
    pushMessage,
    addSystem,
    addEditMessage,
    send,
    stop,
    clear,
  };
});

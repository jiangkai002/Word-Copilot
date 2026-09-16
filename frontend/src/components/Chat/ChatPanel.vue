<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue";
import { useChatStore } from "@/stores/chat";
import { useEditStore } from "@/stores/edits";
import ChatMessage from "./ChatMessage.vue";
import EditResult from "@/components/Edit/EditResult.vue";

const chat = useChatStore();
const editsStore = useEditStore();

const container = ref<HTMLElement | null>(null);

const isEmpty = computed(() => chat.messages.length === 0);

async function scrollToBottom(): Promise<void> {
  await nextTick();
  const el = container.value;
  if (el) el.scrollTop = el.scrollHeight;
}

// 新消息或流式内容追加时自动滚到底部
watch(() => chat.messages.length, () => void scrollToBottom());
watch(
  () => {
    const last = chat.messages[chat.messages.length - 1];
    return last ? last.content.length : 0;
  },
  () => void scrollToBottom(),
);
</script>

<template>
  <div ref="container" class="chat-panel">
    <div v-if="isEmpty" class="empty-hint">
      <p class="hint-title">你好，我是 Word AI 助手</p>
      <p>选中文字，或将光标放在段落中，然后：</p>
      <ul>
        <li>直接提问 —— AI 按上方选择的范围读取上下文</li>
        <li>下达修改指令（如“润色这段话”）或点击快捷命令 ——<br />AI 修改已有文字，以 Word 修订写入文档</li>
        <li>全文 / 批量指令（如“全文纠错”）—— AI 读取全文快照，<br />逐段提交修改提案，生成多张编辑卡片</li>
      </ul>
      <p class="hint-mode">
        输入框上方可切换模式：<b>智能</b>（自动识别指令）· <b>对话</b>（仅聊天）·
        <b>Agent</b>（AI 自主决定是否修改文档）
      </p>
      <p class="hint-note">
        AI 的修改以 Word 修订（Track Changes）写入，可逐条接受或拒绝，
        也可随时用 Word 原生审阅处理。<br />
        注意：AI 只能修改文档中已有的文字（单段修改限选区 / 光标所在段落，
        批量指令可覆盖全文各段），不能在空白处凭空写入新内容。
      </p>
    </div>

    <template v-for="m in chat.messages" :key="m.id">
      <ChatMessage v-if="m.type === 'user' || m.type === 'assistant'" :message="m" />
      <div v-else-if="m.type === 'system'" class="system-msg" :class="{ error: !!m.errorCode }">
        {{ m.content }}
      </div>
      <EditResult
        v-else-if="m.type === 'edit' && m.transactionId && editsStore.getTransaction(m.transactionId)"
        :transaction-id="m.transactionId"
      />
    </template>
  </div>
</template>

<style scoped>
.chat-panel {
  flex: 1;
  overflow-y: auto;
  padding: 10px;
  min-height: 0;
}

.empty-hint {
  color: var(--text-secondary);
  font-size: 12.5px;
  padding: 18px 6px;
}

.hint-title {
  font-size: 13.5px;
  font-weight: 600;
  color: var(--text);
  margin-bottom: 8px;
}

.empty-hint ul {
  margin: 6px 0;
  padding-left: 18px;
}

.hint-mode {
  margin-top: 6px;
  font-size: 12px;
  color: var(--text-secondary);
}

.hint-mode b {
  color: var(--text);
  font-weight: 600;
}

.hint-note {
  margin-top: 10px;
  padding: 8px;
  background: var(--panel-bg);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  font-size: 12px;
}

.system-msg {
  margin: 6px 0;
  padding: 6px 10px;
  font-size: 12px;
  color: var(--text-secondary);
  background: var(--panel-bg);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  white-space: pre-wrap;
  word-break: break-word;
}

.system-msg.error {
  color: var(--danger);
  background: var(--danger-weak);
  border-color: var(--danger);
}
</style>

<script setup lang="ts">
import { computed, ref } from "vue";
import { useChatStore } from "@/stores/chat";
import { useAgentStore } from "@/stores/agent";
import { useEditStore } from "@/stores/edits";
import { QUICK_COMMANDS, type QuickCommand } from "@/commands/QuickCommands";
import { routeInput, type ChatMode } from "@/commands/intent";

const chat = useChatStore();
const agentStore = useAgentStore();
const editsStore = useEditStore();
const draft = ref("");

/** 模式开关（自由输入的路由方式；快捷命令自带类型不受影响） */
const MODES: { value: ChatMode; label: string; title: string }[] = [
  { value: "auto", label: "智能", title: "自动识别指令类型：批量任务 → Agent，修改类指令 → 单段修改，其余 → 对话" },
  { value: "chat", label: "对话", title: "仅对话：所有输入都按聊天处理，不修改文档" },
  { value: "agent", label: "Agent", title: "AI 自主决定：读取全文快照，需要修改文档时提交修改提案（逐条审阅）" },
];

const placeholder = computed(
  () =>
    ({
      auto: "提问，或下达修改指令（如“润色这段话”）；Enter 发送…",
      chat: "对话模式：仅提问与讨论，不会修改文档；Enter 发送…",
      agent: "Agent 模式：提问或下达任务，AI 自主决定是否修改文档；Enter 发送…",
    })[chat.mode],
);

async function submit(): Promise<void> {
  const text = draft.value.trim();
  if (!text || chat.streaming || agentStore.working) return;
  draft.value = "";
  // 模式 + §17 意图路由：chat 模式一律对话；agent 模式一律交给 Agent
  //（模型自主决定回答还是提案）；auto 模式关键词启发式 batch > edit > chat
  const route = routeInput(text, chat.mode);
  if (route === "agent") {
    await agentStore.runAgent(text);
  } else if (route === "edit") {
    chat.pushMessage({ type: "user", content: text });
    await editsStore.sendEdit(text);
  } else {
    await chat.send(text);
  }
}

function runCommand(cmd: QuickCommand): void {
  if (cmd.kind === "chat") {
    void chat.send(cmd.instruction);
  } else if (cmd.kind === "agent") {
    void agentStore.runAgent(cmd.instruction);
  } else {
    void editsStore.sendEdit(cmd.instruction);
  }
}

/** Enter 发送 / Shift+Enter 换行；中文输入法组合中的 Enter 不发送 */
function onKeydown(e: KeyboardEvent): void {
  if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
    e.preventDefault();
    void submit();
  }
}
</script>

<template>
  <div class="chat-input-area">
    <div class="mode-row" role="group" aria-label="输入模式">
      <button
        v-for="m in MODES"
        :key="m.value"
        class="mode-pill"
        :class="{ active: chat.mode === m.value }"
        :title="m.title"
        @click="chat.setMode(m.value)"
      >
        {{ m.label }}
      </button>
    </div>

    <div class="quick-commands">
      <button
        v-for="cmd in QUICK_COMMANDS"
        :key="cmd.id"
        class="cmd-chip"
        :class="cmd.kind"
        :title="cmd.instruction"
        :disabled="chat.streaming || editsStore.working || agentStore.working"
        @click="runCommand(cmd)"
      >
        {{ cmd.label }}
      </button>
    </div>

    <div class="input-row">
      <textarea
        v-model="draft"
        class="input-box"
        rows="2"
        :placeholder="placeholder"
        :disabled="chat.streaming"
        @keydown="onKeydown"
      />
      <button v-if="chat.streaming" class="btn stop-btn" title="停止生成" @click="chat.stop()">
        停止
      </button>
      <button v-else class="btn primary send-btn" :disabled="!draft.trim()" @click="submit()">
        发送
      </button>
    </div>
  </div>
</template>

<style scoped>
.chat-input-area {
  flex-shrink: 0;
  border-top: 1px solid var(--border);
  background: var(--bg);
  padding: 6px 10px 10px;
}

.mode-row {
  display: flex;
  gap: 3px;
  margin-bottom: 6px;
}

.mode-pill {
  font-size: 11.5px;
  padding: 2px 10px;
  border: 1px solid var(--border);
  border-radius: 9px;
  color: var(--text-secondary);
  background: var(--bg);
  cursor: pointer;
}

.mode-pill.active {
  border-color: var(--accent);
  color: var(--accent);
  background: var(--accent-weak);
  font-weight: 600;
}

.quick-commands {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  margin-bottom: 6px;
}

.cmd-chip {
  font-size: 12px;
  padding: 2px 8px;
  border: 1px solid var(--border-strong);
  border-radius: 10px;
  color: var(--text-secondary);
  background: var(--bg);
}

.cmd-chip:hover:not(:disabled) {
  border-color: var(--accent);
  color: var(--accent);
  background: var(--accent-weak);
}

.cmd-chip.edit:hover:not(:disabled) {
  border-color: var(--success);
  color: var(--success);
  background: var(--success-weak);
}

.cmd-chip:disabled {
  opacity: 0.5;
}

.input-row {
  display: flex;
  gap: 6px;
  align-items: flex-end;
}

.input-box {
  flex: 1;
  resize: none;
  font: inherit;
  font-size: 13px;
  color: var(--text);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: var(--radius);
  padding: 6px 8px;
  max-height: 120px;
}

.input-box:focus {
  outline: none;
  border-color: var(--accent);
}

.input-box:disabled {
  opacity: 0.6;
}

.send-btn,
.stop-btn {
  height: 32px;
  min-width: 52px;
}

.stop-btn {
  border-color: var(--danger);
  color: var(--danger);
}
</style>

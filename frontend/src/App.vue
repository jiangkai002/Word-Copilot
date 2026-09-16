<script setup lang="ts">
import { onMounted, onUnmounted } from "vue";
import { useSettingsStore } from "@/stores/settings";
import { useDocumentStore } from "@/stores/document";
import { useEditStore } from "@/stores/edits";
import { useChatStore } from "@/stores/chat";
import ChatPanel from "@/components/Chat/ChatPanel.vue";
import ChatInput from "@/components/Chat/ChatInput.vue";
import ContextSelector from "@/components/Context/ContextSelector.vue";
import SettingsPanel from "@/components/Settings/SettingsPanel.vue";

const settings = useSettingsStore();
const documentStore = useDocumentStore();
const editsStore = useEditStore();
const chatStore = useChatStore();

/**
 * 任务窗格重新可见 / 获得焦点时：
 * - 对账 pending 事务（用户可能刚在 Word 原生审阅中处理了修订，§29）
 * - 刷新选区状态
 */
function onWake(): void {
  if (document.visibilityState !== "visible") return;
  void editsStore.reconcile();
  void documentStore.refreshSelection();
}

onMounted(() => {
  document.addEventListener("visibilitychange", onWake);
  window.addEventListener("focus", onWake);
});

onUnmounted(() => {
  document.removeEventListener("visibilitychange", onWake);
  window.removeEventListener("focus", onWake);
});
</script>

<template>
  <div class="app">
    <header class="app-header">
      <div class="app-title">
        Word AI
        <span v-if="editsStore.pendingList.length > 0" class="pending-badge" :title="`${editsStore.pendingList.length} 个待处理修改`">
          {{ editsStore.pendingList.length }}
        </span>
      </div>
      <div class="header-actions">
        <span class="status-dot" :class="settings.backendStatus" :title="`后端：${settings.statusLabel}`" />
        <button class="ghost-btn" title="开始新对话" @click="chatStore.clear()">新对话</button>
        <button class="ghost-btn" title="设置" @click="settings.settingsOpen = true">设置</button>
      </div>
    </header>

    <div v-if="documentStore.hostReady && !documentStore.revisionSupported" class="cap-banner">
      当前 Word 版本不支持修订功能（需 WordApi 1.6）。对话可用，AI 修改已禁用。
    </div>

    <ContextSelector v-if="documentStore.hostReady" />

    <ChatPanel />

    <ChatInput />

    <SettingsPanel v-if="settings.settingsOpen" />
  </div>
</template>

<style scoped>
.app {
  display: flex;
  flex-direction: column;
  height: 100%;
}

.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 6px 10px;
  border-bottom: 1px solid var(--border);
  background: var(--panel-bg);
  flex-shrink: 0;
}

.app-title {
  font-weight: 600;
  font-size: 13px;
  display: flex;
  align-items: center;
  gap: 6px;
}

.pending-badge {
  font-size: 11px;
  font-weight: 600;
  line-height: 1;
  padding: 2px 6px;
  border-radius: 8px;
  background: var(--accent-weak);
  color: var(--accent);
}

.header-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}

.ghost-btn {
  font-size: 12px;
  color: var(--text-secondary);
  padding: 2px 6px;
  border-radius: var(--radius);
}

.ghost-btn:hover {
  background: var(--panel-bg-2);
  color: var(--text);
}

.status-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: var(--text-muted);
}

.status-dot.ok {
  background: var(--success);
}

.status-dot.error {
  background: var(--danger);
}

.status-dot.checking {
  background: var(--warning-border);
}

.cap-banner {
  flex-shrink: 0;
  padding: 5px 10px;
  font-size: 12px;
  background: var(--warning-bg);
  border-bottom: 1px solid var(--warning-border);
  color: var(--text-secondary);
}
</style>

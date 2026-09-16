<script setup lang="ts">
import { computed } from "vue";
import { useDocumentStore } from "@/stores/document";
import type { ContextType } from "@/models/DocumentContext";

const documentStore = useDocumentStore();

interface ModeOption {
  value: ContextType;
  label: string;
  title: string;
}

const MODES: ModeOption[] = [
  { value: "selection", label: "选区", title: "仅发送当前选中的文字" },
  { value: "paragraph", label: "段落", title: "发送光标所在段落及前后段落" },
  { value: "section", label: "章节", title: "发送当前章节全文与文档标题结构" },
  { value: "document", label: "全文", title: "发送整篇文档（过长时自动截取相关部分）" },
];

/** §13：UI 明示 AI 将读取的内容 */
const contextHint = computed(() => {
  if (documentStore.contextMode === "selection") {
    if (documentStore.selectionAvailable) {
      const preview =
        documentStore.selectionTextPreview.length > 24
          ? `${documentStore.selectionTextPreview.slice(0, 24)}…`
          : documentStore.selectionTextPreview;
      return `AI 将读取选区（${documentStore.selectionLength} 字）：${preview}`;
    }
    return "未选择文字 —— 将按段落模式读取";
  }
  if (documentStore.contextMode === "paragraph") return "AI 将读取光标段落及其前后段";
  if (documentStore.contextMode === "section") return "AI 将读取当前章节与文档结构";
  return "AI 将读取整篇文档";
});
</script>

<template>
  <div class="context-bar">
    <div class="context-modes" role="tablist">
      <button
        v-for="mode in MODES"
        :key="mode.value"
        class="mode-btn"
        :class="{ active: documentStore.contextMode === mode.value }"
        :title="mode.title"
        @click="documentStore.setContextMode(mode.value)"
      >
        {{ mode.label }}
      </button>
    </div>
    <div class="context-hint" :title="contextHint">{{ contextHint }}</div>
  </div>
</template>

<style scoped>
.context-bar {
  flex-shrink: 0;
  border-bottom: 1px solid var(--border);
  background: var(--bg);
  padding: 6px 10px 7px;
}

.context-modes {
  display: flex;
  gap: 2px;
}

.mode-btn {
  flex: 1;
  font-size: 12px;
  padding: 3px 0;
  border-radius: var(--radius);
  color: var(--text-secondary);
  border: 1px solid transparent;
}

.mode-btn:hover {
  background: var(--panel-bg-2);
}

.mode-btn.active {
  background: var(--accent-weak);
  color: var(--accent);
  border-color: var(--accent);
  font-weight: 600;
}

.context-hint {
  margin-top: 5px;
  font-size: 11.5px;
  color: var(--text-muted);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
</style>

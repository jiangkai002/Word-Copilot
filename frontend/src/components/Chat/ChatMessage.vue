<script setup lang="ts">
import { computed } from "vue";
import type { ChatMessage } from "@/models/ChatMessage";
import { renderMarkdown } from "@/utils/markdown";
// KaTeX 样式在组件层引入（markdown.ts 保持纯 JS —— CSS 导入会拖挂 vitest 的
// 资源管线；公式 HTML 由渲染器产出，样式属于展示层）
import "katex/dist/katex.min.css";

const props = defineProps<{ message: ChatMessage }>();

/** AI 回复渲染受限 Markdown（转义先行 + 白名单标签，无 XSS 面，见 utils/markdown） */
const html = computed(() =>
  props.message.type === "assistant" ? renderMarkdown(props.message.content) : "",
);
</script>

<template>
  <div class="chat-message" :class="[message.type, { error: !!message.errorCode }]">
    <div class="msg-role">{{ message.type === "user" ? "你" : "AI" }}</div>
    <div class="msg-body">
      <!-- AI 回复：markdown 渲染（流式时增量重渲染，未闭合代码块自然补全） -->
      <template v-if="message.type === 'assistant'">
        <div v-if="html" class="md-body" v-html="html" />
        <span v-if="message.streaming" class="cursor" />
      </template>
      <!-- 用户输入：按原文展示 -->
      <span v-else class="msg-text">{{ message.content }}</span>
    </div>
  </div>
</template>

<style scoped>
.chat-message {
  margin: 8px 0;
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.msg-role {
  font-size: 11px;
  font-weight: 600;
  color: var(--text-muted);
}

.chat-message.user .msg-role {
  color: var(--accent);
}

.msg-body {
  font-size: 13px;
  white-space: pre-wrap;
  word-break: break-word;
}

.chat-message.user .msg-text {
  display: inline-block;
  padding: 5px 10px;
  background: var(--accent-weak);
  border-radius: var(--radius);
}

.chat-message.error .msg-text {
  color: var(--danger);
}

.chat-message.error .md-body {
  color: var(--danger);
}

/* ---- AI 回复 markdown（v-html 内容，须 :deep） ---- */
.md-body {
  white-space: normal;
  line-height: 1.6;
}

.md-body :deep(p) {
  margin: 3px 0;
}

.md-body :deep(h1),
.md-body :deep(h2),
.md-body :deep(h3),
.md-body :deep(h4),
.md-body :deep(h5),
.md-body :deep(h6) {
  margin: 8px 0 4px;
  line-height: 1.4;
  font-weight: 600;
}

.md-body :deep(h1) {
  font-size: 15px;
}

.md-body :deep(h2) {
  font-size: 14px;
}

.md-body :deep(h3),
.md-body :deep(h4),
.md-body :deep(h5),
.md-body :deep(h6) {
  font-size: 13.5px;
}

.md-body :deep(ul),
.md-body :deep(ol) {
  margin: 3px 0;
  padding-left: 20px;
}

.md-body :deep(li) {
  margin: 2px 0;
}

.md-body :deep(code) {
  font-family: Consolas, "Courier New", monospace;
  font-size: 12px;
  background: var(--panel-bg);
  border: 1px solid var(--border);
  border-radius: 3px;
  padding: 0 4px;
}

.md-body :deep(pre) {
  margin: 4px 0;
  padding: 8px 10px;
  background: var(--panel-bg);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  overflow-x: auto;
}

.md-body :deep(pre code) {
  background: none;
  border: none;
  padding: 0;
  white-space: pre;
}

.md-body :deep(blockquote) {
  margin: 4px 0;
  padding: 3px 10px;
  border-left: 3px solid var(--accent);
  background: var(--panel-bg);
  color: var(--text-secondary);
}

.md-body :deep(a) {
  color: var(--accent);
}

.md-body :deep(hr) {
  border: none;
  border-top: 1px solid var(--border);
  margin: 8px 0;
}

/* ---- GFM 表格 ---- */
.md-body :deep(table) {
  margin: 6px 0;
  border-collapse: collapse;
  max-width: 100%;
}

.md-body :deep(th),
.md-body :deep(td) {
  border: 1px solid var(--border);
  padding: 3px 8px;
  font-size: 12.5px;
  word-break: break-word;
  white-space: normal;
}

.md-body :deep(th) {
  background: var(--panel-bg);
  font-weight: 600;
}

/* ---- KaTeX 公式（样式主体在 katex.min.css，此处做面板适配） ---- */
.md-body :deep(.katex) {
  font-size: 1.05em;
}

.md-body :deep(.katex-display) {
  margin: 6px 0;
  overflow-x: auto;
  overflow-y: hidden;
  padding: 2px 0;
}

.cursor {
  display: inline-block;
  width: 7px;
  height: 14px;
  margin-left: 2px;
  vertical-align: -2px;
  background: var(--accent);
  animation: blink 1s step-end infinite;
}

@keyframes blink {
  50% {
    opacity: 0;
  }
}
</style>

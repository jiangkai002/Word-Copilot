<script setup lang="ts">
import { ref, watch } from "vue";
import { useSettingsStore } from "@/stores/settings";

const settings = useSettingsStore();
const urlDraft = ref(settings.backendBaseUrl);
const saving = ref(false);

watch(
  () => settings.settingsOpen,
  (open) => {
    if (open) urlDraft.value = settings.backendBaseUrl;
  },
);

async function saveAndTest(): Promise<void> {
  saving.value = true;
  settings.setBackendBaseUrl(urlDraft.value);
  await settings.checkBackend();
  saving.value = false;
}

function close(): void {
  settings.settingsOpen = false;
}
</script>

<template>
  <div class="settings-overlay" @click.self="close">
    <div class="settings-panel">
      <div class="panel-head">
        <span>设置</span>
        <button class="close-btn" title="关闭" @click="close">×</button>
      </div>

      <label class="field-label">后端服务地址</label>
      <div class="url-row">
        <input
          v-model="urlDraft"
          class="url-input"
          type="text"
          placeholder="留空 = 相对路径（开发模式经 Vite 代理）"
          spellcheck="false"
        />
        <button class="btn primary" :disabled="saving" @click="saveAndTest">
          {{ saving ? "测试中…" : "保存并测试" }}
        </button>
      </div>
      <p class="field-hint">
        开发环境默认留空即可：任务窗格经 Vite 开发服务器代理到
        http://localhost:8000。生产部署请填写后端实际地址（须 HTTPS）。
      </p>

      <div class="status-row">
        <span class="status-dot" :class="settings.backendStatus" />
        <span>{{ settings.statusLabel }}</span>
        <span v-if="settings.backendError" class="status-error" :title="settings.backendError">
          {{ settings.backendError }}
        </span>
      </div>

      <div v-if="settings.modelInfo" class="model-row">
        <span class="field-label">模型</span>
        <code class="model-code">{{ settings.modelInfo.model }}</code>
        <span class="model-provider">{{ settings.modelInfo.provider }}</span>
      </div>
    </div>
  </div>
</template>

<style scoped>
.settings-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.32);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 16px;
}

.settings-panel {
  width: 100%;
  max-width: 320px;
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: var(--radius);
  padding: 12px 14px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.18);
}

.panel-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  font-weight: 600;
  font-size: 13px;
  margin-bottom: 10px;
}

.close-btn {
  font-size: 16px;
  line-height: 1;
  color: var(--text-secondary);
  padding: 0 4px;
}

.close-btn:hover {
  color: var(--text);
}

.field-label {
  display: block;
  font-size: 12px;
  color: var(--text-secondary);
  margin-bottom: 4px;
}

.url-row {
  display: flex;
  gap: 6px;
}

.url-input {
  flex: 1;
  min-width: 0;
  font: inherit;
  font-size: 12.5px;
  font-family: Consolas, "Courier New", monospace;
  color: var(--text);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: var(--radius);
  padding: 4px 8px;
}

.url-input:focus {
  outline: none;
  border-color: var(--accent);
}

.field-hint {
  margin: 6px 0 10px;
  font-size: 11.5px;
  color: var(--text-muted);
}

.status-row {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12.5px;
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

.status-error {
  color: var(--danger);
  font-size: 11.5px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.model-row {
  margin-top: 10px;
  display: flex;
  align-items: center;
  gap: 6px;
}

.model-row .field-label {
  margin: 0;
}

.model-code {
  font-size: 11.5px;
  background: var(--code-bg);
  padding: 1px 6px;
  border-radius: var(--radius);
  word-break: break-all;
}

.model-provider {
  font-size: 11.5px;
  color: var(--text-muted);
}
</style>

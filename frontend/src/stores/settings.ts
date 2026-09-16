/**
 * Settings Store：后端地址 + 连接状态 + 模型信息。
 * baseUrl 持久化到 localStorage；其余为会话状态。
 */
import { defineStore } from "pinia";
import { ref, computed } from "vue";
import type { ConfigResponse, HealthResponse } from "@/models/Api";
import { chatApi } from "@/services/api/ChatApi";
import { setApiBaseUrl } from "@/services/api/ApiClient";
import { logger } from "@/utils/logger";

const STORAGE_KEY = "word-ai-copilot.settings";

export type BackendStatus = "unknown" | "checking" | "ok" | "error";

export const useSettingsStore = defineStore("settings", () => {
  const backendBaseUrl = ref("");
  const backendStatus = ref<BackendStatus>("unknown");
  const backendError = ref("");
  const modelInfo = ref<ConfigResponse | null>(null);
  const settingsOpen = ref(false);

  const statusLabel = computed(() => {
    switch (backendStatus.value) {
      case "ok":
        return "已连接";
      case "checking":
        return "连接中…";
      case "error":
        return "未连接";
      default:
        return "未检测";
    }
  });

  function init(): void {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved) as { backendBaseUrl?: string };
        if (typeof parsed.backendBaseUrl === "string") {
          backendBaseUrl.value = parsed.backendBaseUrl;
        }
      }
    } catch {
      logger.warn("读取本地设置失败，使用默认值");
    }
    setApiBaseUrl(backendBaseUrl.value);
  }

  function setBackendBaseUrl(url: string): void {
    const normalized = url.trim().replace(/\/+$/, "");
    backendBaseUrl.value = normalized;
    setApiBaseUrl(normalized);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ backendBaseUrl: normalized }));
    } catch {
      logger.warn("保存设置失败");
    }
  }

  async function checkBackend(): Promise<HealthResponse | null> {
    backendStatus.value = "checking";
    backendError.value = "";
    try {
      const health = await chatApi.health();
      try {
        modelInfo.value = await chatApi.config();
      } catch {
        modelInfo.value = null; // /config 是可选接口
      }
      backendStatus.value = "ok";
      return health;
    } catch (err) {
      backendStatus.value = "error";
      backendError.value = err instanceof Error ? err.message : String(err);
      return null;
    }
  }

  return {
    backendBaseUrl,
    backendStatus,
    backendError,
    modelInfo,
    settingsOpen,
    statusLabel,
    init,
    setBackendBaseUrl,
    checkBackend,
  };
});

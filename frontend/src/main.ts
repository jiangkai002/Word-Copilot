/**
 * 入口：等待 Office.onReady 后再挂载应用（保证 stores 能立即探测宿主能力）。
 * 在纯浏览器中打开（开发调试）时 Office 未定义 —— 优雅降级为普通页面。
 */
import { createApp } from "vue";
import { createPinia } from "pinia";
import App from "./App.vue";
import "./styles/main.css";
import { useSettingsStore } from "./stores/settings";
import { useDocumentStore } from "./stores/document";
import { logger } from "./utils/logger";

function bootstrap(): void {
  const app = createApp(App);
  const pinia = createPinia();
  app.use(pinia);
  app.mount("#app");

  // stores 初始化（pinia 已激活，可安全在组件外调用）
  useSettingsStore(pinia).init();
  useDocumentStore(pinia).init();
  void useSettingsStore(pinia).checkBackend();
}

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    logger.info(`Office 主机就绪：${info.host}`);
    bootstrap();
  });
} else {
  logger.warn("未检测到 Office.js —— 以浏览器模式运行（仅可调试 UI）");
  bootstrap();
}

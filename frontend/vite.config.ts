import { fileURLToPath, URL } from "node:url";
import { createRequire } from "node:module";
import { defineConfig } from "vitest/config";
import vue from "@vitejs/plugin-vue";

const require = createRequire(import.meta.url);

/** vitest 运行时不启用 HTTPS / 开发证书 —— 避免 worker 内部 fetch 自签证书超时 */
const runningVitest = process.env.VITEST !== undefined || process.env.NODE_ENV === "test";

/**
 * 开发阶段使用 Office Add-in 官方推荐的 localhost 开发证书（office-addin-dev-certs）。
 * v3 的 getHttpsServerOptions() 是异步的（返回 Promise<{ca, cert, key}>），
 * 必须等待完成后再交给 Vite —— 否则任务窗格的 TLS 握手会失败。
 * 首次运行前请执行：npm install && npm run dev-certs
 */
async function httpsServerOptions(): Promise<Record<string, unknown> | undefined> {
  if (runningVitest) return undefined;
  try {
    const devCerts = require("office-addin-dev-certs") as {
      getHttpsServerOptions: () => Promise<Record<string, unknown>>;
    };
    return await devCerts.getHttpsServerOptions();
  } catch (err) {
    console.warn(
      "[vite] office-addin-dev-certs 不可用（尚未 npm install？），开发服务器暂时使用 HTTP。" +
        "请先执行 npm install && npm run dev-certs，否则 Word 无法加载 https://localhost:3000 的任务窗格。",
    );
    return undefined;
  }
}

export default defineConfig(async () => ({
  plugins: [vue()],
  resolve: {
    alias: {
      "@": fileURLToPath(new URL("./src", import.meta.url)),
    },
  },
  server: {
    port: 3000,
    strictPort: true,
    https: await httpsServerOptions(),
    // 开发阶段把 /api 反向代理到本地 FastAPI，任务窗格(https)不会遇到混合内容问题
    // 端口 8100：本机 8000 常被其他服务占用（可用环境变量 BACKEND_PORT 覆盖）
    proxy: {
      "/api": {
        target: `http://localhost:${process.env.BACKEND_PORT ?? 8100}`,
        changeOrigin: false,
      },
    },
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: fileURLToPath(new URL("./taskpane.html", import.meta.url)),
    },
  },
  test: {
    environment: "node",
    include: ["tests/**/*.test.ts"],
  },
}));

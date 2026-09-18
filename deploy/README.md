# 分发部署包（只分发 manifest）

把 Word AI 插件给别人用的最小分发物：**一个 manifest.xml**。
任务窗格前端与 FastAPI 后端都运行在服务方（你）的机器上，使用者的电脑
上不安装任何代码、不配置 API Key —— Word 只是从你的机器加载网页。

```
使用者 Word ──manifest 注册──► 你的机器（一个 HTTPS 源）
                                 ├── taskpane.html + 静态资源（frontend/dist）
                                 └── /api 反代 → 127.0.0.1:8100（FastAPI）
```

## 一、服务方（你）的前置条件

manifest 只是"指针"，它指向的源必须真实可用：

1. **构建前端**：`cd frontend && npm run build`（产出 `dist/`）。
2. **一个 HTTPS 源**同时提供 `dist/` 静态文件，并把同源 `/api` 反代到本机
   FastAPI（8100）。任选其一：
   - Caddy（自动内网 CA，最省心）：`file_server` + `reverse_proxy /api/*`；
   - nginx / 其他反代同理。
   - 注意：**不能用** `npm run dev` 的 localhost 证书对外 —— 那是本机专用。
3. **证书**必须被使用者机器信任：
   - 内网：自建 CA，把根证书 `.crt` 发给使用者导入一次（见下文）；
   - 公网：域名证书（如 Cloudflare Tunnel，使用者无需导入任何证书）。
4. **CORS**：`backend/.env` 的 `BACKEND_CORS_ORIGINS` 加入该 HTTPS 源
   （逗号分隔，禁止 `*`）。
5. **防火墙**放行对外端口；你的电脑在线插件才可用（关机 = 全员不可用）。
6. **隐私告知**（建议明确告知使用者）：
   - LLM API Key 只在你的机器上，计费走你的账号；
   - `backend/logs/conversations.jsonl` 默认记录对话与文档正文到你机器
     （可设 `CONVERSATION_LOG_INCLUDE_CONTENT=false` 关闭）；
   - `backend/data/skills.json` 为所有使用者共享。

## 二、生成分发版 manifest

```bash
# 仓库根目录
py deploy\make_manifest.py --base https://你的机器名:端口
# 可选：
#   --new-id   生成新插件 GUID（见下文"插件身份"）
#   -o 路径    指定输出位置（默认 deploy\dist\manifest.xml）
```

脚本做的事：把 `frontend/manifest.xml` 中全部 8 处
`https://localhost:3000`（任务窗格 / 图标 / AppDomain）替换为你给的源，
并做 HTTPS / 非 localhost / XML 良构校验。产物即分发的全部内容。

使用者**无需**在任务窗格"设置"里填后端地址 —— 同源 `/api` 直达。

## 三、使用者安装（二选一）

### 方式 1：共享文件夹受信目录（推荐，使用者零命令行）

1. 把生成的 `manifest.xml` 放进一个网络共享目录（如 `\\你的机器\wordai`）。
2. 使用者在 Word 中：**文件 → 选项 → 信任中心 → 信任中心设置 →
   受信任的加载项目录**，添加该共享路径，确定后**完全退出并重启 Word**。
3. **插入 → 获取加载项 → 共享文件夹**（或"我的加载项"→"共享文件夹"），
   选择 Word AI → 添加。功能区"开始"选项卡出现"AI 助手"分组。

### 方式 2：命令行 sideload（使用者需装 Node.js）

```bash
npx office-addin-dev-settings register manifest.xml
```

（用户级注册表注册，持久；卸载用 `unregister manifest.xml`。）

### 一次性：信任内网证书（仅方式 A 自建 CA 时需要）

双击根证书 `.crt` → 安装证书 → 存储位置"当前用户" →
"将所有的证书都放入下列存储" → **受信任的根证书颁发机构**。
未导入时 Word 会**静默拒绝**加载任务窗格（功能区按钮在、窗格空白，无报错）。

## 四、插件身份（Id）

- 默认沿用源 manifest 的 GUID：不同机器互不影响；但**同一台机器**上，
  分发版与 localhost 开发版（`npm run sideload`）互相顶替 —— 同一个 Id
  只能注册一份。
- 想在开发机上同时保留两个入口：生成时加 `--new-id`。

## 五、使用者环境要求

- Windows 桌面版 Microsoft 365 Word（AI 修订需 WordApi 1.6；旧版自动
  降级为"仅对话"，不报错）；
- 能访问你的 HTTPS 源（同网段或隧道可达）。

常见排查：

| 症状 | 原因 |
| --- | --- |
| 功能区按钮在、任务窗格空白 | 证书不受信任（静默失败）或源不可达 |
| 任务窗格打开但请求全失败 | `/api` 反代未配好，或 CORS 未加源 |
| 对话可用、修改报"不支持修订" | 对方 Word < WordApi 1.6（预期降级） |

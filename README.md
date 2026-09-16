# Word AI Copilot

Word 任务窗格 AI 助手：在 Microsoft Word 中与 AI 对话（基于选区 / 段落 / 章节 / 全文上下文），
并让 AI 直接修改文档 —— 所有修改以 **Word 原生修订（Track Changes）** 写入，
可逐条接受 / 拒绝 / 重新生成，与人工修订完全隔离。

Agent 模式下 AI 除整段改写外，还可提交四类结构化提案（同样走修订与事务卡片管线）：
**字体 / 段落格式修改**（加粗、斜体、下划线、删除线、字体、字号、颜色、对齐）、
**插入表格**（TableGrid 样式、表头行）、**插入数学公式**（LaTeX → OMML，Word
原生公式对象，可双击编辑）、**插入纯文字段落**（换行分段）。插入类内容可落在
锚点段之后或**文档末尾**（锚点可省略 —— 空文档也能直接生成内容）。

## 核心原则

- **LLM 永不直接操作 Word**：`LLM → EditPlan → 校验 → Diff → Patch → Office.js → Track Changes`
  （agent 模式同理：`Agent → 工具提案 → 逐条独立管线`）
- **绝不静默覆盖文档**：AI 修改全部经由修订（Track Changes）写入
- **修订隔离**：每次修改是一个事务（Content Control 边界），接受 / 拒绝只作用于该事务，
  绝不调用文档级 acceptAll —— 人工修订不受影响
- **乐观锁**：应用修改前校验目标文本哈希（SHA-256），文档变化即拒绝（`DOCUMENT_CHANGED`）
- **Diff 在客户端计算**：后端只返回 `original_text / new_text`，
  字符级差异（diff-match-patch）在任务窗格内计算并从后向前应用
- **隐私**：只发送完成请求所需的文档内容（选区 / 段落模式不发送全文），UI 明示 AI 读取范围
  （agent / @document 模式发送全文快照，发送前在聊天区披露段数与字数）

## 架构

```
┌─────────────────────────────  Word（桌面 / 365） ─────────────────────────────┐
│  Task Pane（iframe）                                                          │
│  ┌────────────────────────────────────────────────────────────────────────┐  │
│  │ Vue 3 + Pinia + TypeScript（frontend/）                                 │  │
│  │  UI 组件 ── Pinia Stores（chat / edits / agent / document / settings）  │  │
│  │                │                                                       │  │
│  │                ├─ services/word/   Office.js 全部封装于此               │  │
│  │                │   SelectionService / DocumentService / RangeLocator    │  │
│  │                │   PatchEngine / FormatEngine / InsertEngine /          │  │
│  │                │   FormulaOoxml(LaTeX→OMML) / RevisionService / …      │  │
│  │                │        │（Office.js，proxy 不跨 Word.run 批次）        │  │
│  │                ├─ services/agent/  提案 → 编辑目标映射（纯函数）         │  │
│  │                ├─ services/diff/   DiffEngine（diff-match-patch）       │  │
│  │                └─ services/api/    ChatApi(SSE) / EditApi / AgentApi    │  │
│  └───────────────────────────────────┼────────────────────────────────────┘  │
└──────────────────────────────────────┼───────────────────────────────────────┘
                                       │ HTTPS（开发期经 Vite 代理 /api → :8100）
┌──────────────────────────────────────▼───────────────────────────────────────┐
│  Backend（backend/，FastAPI + Python 3.12，无数据库）                          │
│  api/（chat SSE / edit / agent SSE / health / config）                         │
│  services/（Prompt 构造、JSON 解析、Pydantic 校验、agent 编排）                │
│  llm/（LLMProvider 抽象 → OpenAICompatibleProvider，httpx）                   │
│  agent_framework（Microsoft Agent Framework；懒 import，未装时后端仍可启动）  │
│  只做文本理解 / 生成，不知道任何 Word 对象（§75）                                │
└──────────────────────────────────────────────────────────────────────────────┘
```

编辑数据流（§77）：

```
用户指令 → 捕获目标（Selection/段落 + RangeLocator + textHash）
        → POST /api/v1/edit（只发目标与前后文）
        → 后端返回 { original_text, new_text, summary }（Pydantic 校验）
        → 客户端 DiffEngine 计算字符级 EditOperation
        → PatchEngine：定位 → 哈希校验 → Content Control 边界 → TrackAll
          → 从后向前应用 → 结果校验（失败自动回滚）→ 恢复原修订模式
        → EditTransaction（pending）→ 编辑卡片（接受 / 拒绝 / 重新生成）
```

Agent 数据流（§51：文档问答 + 修改提案，模式开关「Agent」下由模型自主决策）：

```
Agent 入口（「全文纠错」快捷命令 / 智能模式 isBatchIntent 或 isAgentToolIntent 命中 /
          Agent 模式一切指令）
        → DocumentService.getDocumentSnapshot()（全部段落 + 大纲，50k 截断）
        → SelectionService.captureAgentFocus()（光标/选区段 → 快照 id 解析 → focus）
        → 聊天区披露读取范围（§60）→ POST /api/v1/agent/stream（SSE）
        → 后端 Agent（Microsoft Agent Framework，快照以闭包注入工具）：
            get_outline / read_paragraph / search_paragraphs（只读快照）
            propose_edit(paragraph_id, new_text, summary)（整段替换提案）
            propose_format(paragraph_id, bold/italic/…/alignment)（格式提案）
            insert_table(anchor_paragraph_id?, values, header)（表格插入提案）
            insert_formula(anchor_paragraph_id?, latex, display)（公式插入提案）
            insert_paragraph(anchor_paragraph_id?, paragraph_text)（纯文字段落提案）
            —— 插入类工具的 anchor 可省略（null = 插入到文档末尾，空文档可用）
            —— 锚点不能是表格单元格内段落（快照标注 in_table /「（表格内）」，
               校验拒绝并引导改用文档末尾 —— 防止内容写进单元格）
        → SSE 双通道：token（回答流式气泡）+ 五类 proposal 帧
        → 流结束 → 逐条 proposalToCapturedTarget → applyProposal（按 kind 分派）：
            text   → §57 冲突检查 → 回显校验 → Diff → PatchEngine → 事务卡片
            format → 冲突检查 → FormatEngine（原值快照 + 软校验）→ 事务卡片
            insert-table  → WordApi 1.3 门 → InsertEngine（先插后包 CC）→ 事务卡片
            insert-formula→ LaTeX→OMML→Flat OPC → 先登记卡片 → TrackAll + CC 写入修订
            insert-paragraph→ 换行分段；若含 $LaTeX$ 则生成 w:r + m:oMath 富段落；
                              先登记卡片 → TrackAll + CC 写入修订
            （插入类 anchor 为 null 时走 Body.insert*("End") —— 无锚点定位/哈希校验）
        → N 张编辑卡片，逐条接受 / 拒绝（所有类型均以 Word 修订写入，按 kind
          分派接受 / 拒绝 / 对账 —— 见 RevisionService）
```

## 环境要求

- **Node.js ≥ 22**（含 npm）
- **Python ≥ 3.12**（Windows 下用 `py` 启动器）
- **Microsoft Word**：桌面版 Microsoft 365（修订功能需要 **WordApi 1.6**；
  不满足时任务窗格自动降级为「仅对话」，不抛异常）
- 操作系统：Windows 10 / 11（sideload 流程以 Windows Word 桌面版为准）

## 安装依赖

```bash
# 前端
cd frontend
npm install

# 后端
cd backend
py -m pip install -r requirements.txt

# 生成 manifest 图标（一次性）
cd frontend && npm run icons

# 生成测试文档（一次性，可选）
cd ..
py -m pip install python-docx
py tools/generate_test_documents.py
```

## HTTPS Certificate

Office Add-in 的任务窗格在 Word 中以 HTTPS iframe 加载，开发期使用
[office-addin-dev-certs](https://www.npmjs.com/package/office-addin-dev-certs) 提供的
可信 localhost 证书（首次执行会请求管理员权限并把证书装入本机受信存储）：

```bash
cd frontend
npm run dev-certs     # 安装 localhost 开发证书
```

证书就绪后 `npm run dev` 会自动以 `https://localhost:3000` 启动；
若证书未安装，Vite 会退回 HTTP 并打印警告（此时 Word 无法加载任务窗格）。

## 启动 Frontend

```bash
cd frontend
BACKEND_PORT=8100 npm run dev  # https://localhost:3000，/api 代理到 http://localhost:8100
```

Vite 已把 `/api/*` 反向代理到本地 FastAPI（默认 8100，`BACKEND_PORT` 环境变量可改），
避免 https 任务窗格访问 http 后端的混合内容问题 —— 任务窗格内后端地址保持留空即可。

## 启动 Backend

```bash
cd backend
# 1. 准备配置
copy .env.example .env      # 然后编辑 .env 填入 LLM_API_KEY 等

# 2. 启动（二选一；端口 8100 —— 本机 8000 常被其他服务占用）
py -m uvicorn app.main:app --reload --port 8100
py run.py
```

健康检查：`curl http://localhost:8100/api/v1/health`
（或直接在任务窗格「设置」里点「保存并测试」）。

### Skill 配置后台

后端启动后访问 [http://localhost:8100/admin/skills](http://localhost:8100/admin/skills)。
管理页支持新建、编辑、启用/停用和删除 Skill。每个 Skill 包含：

- 名称与说明；
- 触发词（逗号分隔；留空表示所有请求均加载）；
- 注入聊天及 Agent 系统提示词的工作指令；
- 启用状态。

配置默认保存在 `backend/data/skills.json`（已加入 `.gitignore`）。如需指定其他
位置，可设置环境变量 `SKILL_STORE_PATH`。管理接口为 `GET/POST /api/v1/skills`
及 `PUT/DELETE /api/v1/skills/{id}`。

## Word Add-in Sideload

```bash
cd frontend
npm run sideload         # 注册到 Word 开发目录（用户级注册表，持久）
npm run sideload:word    # 直接启动 Word 并加载插件（可选，也可手动开 Word）
```

**注意**：`sideload:word` 通过打开一个内嵌 webextension 引用的模板文档来激活插件。
不要给它追加 `-d <文档>` 参数——传了文档后工具只会原样复制该文档（不注入插件引用），
插件不会被加载。想打开自己的测试文档，等插件激活后在 Word 里正常打开即可。

然后：完全退出 Word → 重新打开 Word → 功能区「开始」选项卡 →
「AI 助手」分组 → 点击按钮打开任务窗格。

移除 sideload：

```bash
npm run unsideload       # office-addin-dev-settings unregister manifest.xml
```

### 任务窗格没出现时的排查（运行时日志）

Word 会**静默丢弃**解析失败的 manifest（功能区不出现按钮、也不弹错）。
开启运行时日志可以看到真实原因：

```bash
npx office-addin-dev-settings runtime-log --enable "%TEMP%\office-runtime.log"
# 复现问题后查看日志；关闭：
npx office-addin-dev-settings runtime-log --disable
```

本项目曾踩过的坑：VersionOverrides 里 Word 的 Host 元素必须是
`<Host xsi:type="Document">`（不是 `DocumentHost`）；
资源字符串必须分 `<bt:ShortStrings>`（≤125 字符）和 `<bt:LongStrings>`（≤250 字符），
不存在单一的 `<bt:Strings>`。

## Manifest 配置

`frontend/manifest.xml` 关键点：

- `Id`：插件唯一 GUID（换新插件身份时重新生成）
- `Resources`/`bt:Url`：任务窗格地址，当前为 `https://localhost:3000/taskpane.html`
  （部署时改为生产域名，须 HTTPS）
- 图标：`assets/icon-16/32/80.png`（`npm run icons` 生成的占位图可直接替换）
- **有意未声明 WordApi Requirements**：声明后旧版 Word 会直接拒绝加载插件；
  本项目选择运行时能力探测（`WordApi 1.6`）+ 优雅降级（仅对话），
  见 `frontend/src/services/word/WordService.ts`

## LLM 配置

`backend/.env`（任何 OpenAI 兼容接口均可）：

```ini
LLM_PROVIDER=openai                       # 仅作展示名；协议统一为 OpenAI Compatible
LLM_BASE_URL=https://api.openai.com/v1    # DeepSeek: https://api.deepseek.com/v1
LLM_API_KEY=sk-xxxx
LLM_MODEL=gpt-4o-mini

# 可选
LLM_TIMEOUT_SECONDS=120        # 对话流式超时
LLM_EDIT_TIMEOUT_SECONDS=60    # Edit 生成超时
LLM_TEMPERATURE=0.3
LLM_MAX_TOKENS=2048

# ---- Agent（批量编辑 / 文档问答，Microsoft Agent Framework）----
# 依赖：py -m pip install agent-framework-core agent-framework-openai
# 未安装时后端仍可启动，agent 接口返回 AGENT_NOT_INSTALLED
AGENT_MAX_PROPOSALS=20         # 单次任务最大提案数
AGENT_TIMEOUT_SECONDS=300      # Agent 运行总超时（秒）
AGENT_MAX_TOOL_CALLS=40        # 单次任务最大工具调用次数

# 生产环境必须收紧（逗号分隔，禁止 *）
BACKEND_CORS_ORIGINS=https://your-addin-origin.example.com
```

## 调试方法

- **前端**：浏览器 DevTools 直接 attach 任务窗格 iframe（Word 桌面版：
  右键任务窗格 → 检查 / 或 Edge DevTools）。开发模式下日志输出到 console。
- **前端单测**：`cd frontend && npm run test`（Diff / 定位 / 哈希 / 匹配算法 /
  LaTeX→OMML 转换链 / SSE 帧解析 / 意图路由 / 对账模式判定）
- **类型检查**：`npm run typecheck`
- **后端**：uvicorn 控制台直接看日志（request_id / conversation_id /
  latency / token usage；**不会**打印文档正文）。
  Swagger UI：`http://localhost:8100/docs`
- **SSE 快速验证**：
  `curl -N -X POST http://localhost:8100/api/v1/chat/stream -H "Content-Type: application/json" -d "{\"conversation_id\":\"t\",\"message\":\"你好\"}"`
- **Agent 接口快速验证**（SSE 双通道：token / proposal）：

  ```bash
  curl -N -X POST http://localhost:8100/api/v1/agent/stream -H "Content-Type: application/json" -d @- <<'EOF'
  {
    "conversation_id": "t",
    "instruction": "检查全文错别字，逐段提交修改提案。",
    "history": [],
    "snapshot": {
      "outline": [],
      "paragraphs": [
        {"id": "p1", "text": "他昨天去去了公园。"}
      ],
      "truncated": false
    }
  }
  EOF
  ```

  预期先出现 `event: proposal`（含 original_text / new_text / summary），
  随后 `event: token` 流式输出总结，最后 `event: done`。
- **后端单测**（纯函数，不需要 LLM / agent-framework）：
  `cd backend && py -m pytest tests/ -q`（先 `py -m pip install -r requirements-dev.txt`）
- **Office.js 报错定位**：`OfficeExtension.Error` 的 `code / debugInfo`，
  已由 `frontend/src/utils/errors.ts` 统一映射为 `WORD_API_ERROR`

## Revision POC 测试方法

> 这是全项目最关键的技术验证（§69 / §70），对应测试文档 `test-documents/04-existing-revisions.docx`。

**测试 1 —— AI 修订与人工修订隔离（§69）**

1. 用 Word 打开 `test-documents/04-existing-revisions.docx`
   （文档内已含人工修订 Revision A：一段「插入修订」+ 一段「删除修订」）
2. 选中干净的段落「系统采用传统架构，可以处理大量数据。」
3. 任务窗格 → 快捷命令「润色」→ 等待编辑卡片出现（Revision B 产生）
4. Word 正文此时应显示两类修订；点击编辑卡片「**接受**」
5. **预期**：AI 的修订被接受、卡片变为「已接受」；
   **人工 Revision A 仍完整存在**（审阅窗格可见）

**测试 2 —— 拒绝后精确还原（§70）**

1. 任意选一段纯文本，执行「专业化」
2. 卡片上点「拒绝」
3. **预期**：正文逐字符恢复为原文（文字、段落、格式均不变），
   临时 Content Control 被移除，修订列表回到操作前的状态

**测试 3 —— Word 原生审阅对账（§29）**

1. 产生一个 AI 修改（不点卡片上的按钮）
2. 切到 Word「审阅」→「接受所有修订」（或逐条接受 AI 修订）
3. 回到任务窗格（焦点切回即触发对账）
4. **预期**：卡片状态自动变为「已在 Word 中处理」，无报错

**测试 4 —— 乐观锁（§15）**

1. 选中一段文字，点击「润色」
2. 在 AI 返回前直接在 Word 里改动这段文字
3. **预期**：提示 `DOCUMENT_CHANGED`（目标内容已变化），文档不会被覆盖

**测试 5 —— Agent 批量编辑（§51）**

1. 用 Word 打开 `test-documents/07-chinese-document.docx`
2. 任务窗格 → 快捷命令「全文纠错」（或输入「检查全文的错别字」）
3. 聊天区先显示「已读取全文 N 段（M 字）」披露，随后流式输出总结
4. **预期**：有错的段落各生成一张**独立**编辑卡片（逐条接受 / 拒绝，
   隔离语义同测试 1），干净段落不产生卡片；运行中点「停止」→ 已收到的提案
   全部丢弃，文档无任何变化（§58）
5. 换 `02-heading-document.docx` 问「各章讲什么」→ 验证纯问答（只读工具 + 回答，零卡片）

**测试 6 —— Agent 格式 / 表格 / 公式 / 段落提案（四类结构化能力）**

1. 用 Word 打开任意含多段文字的文档（如 `07-chinese-document.docx`）
2. 任务窗格模式切到「**Agent**」，输入：
   「把第一段加粗并居中；在第二段后面插入 2 列 3 行表格（表头「项目、数量」，
   数据「苹果、5」「香蕉、8」）；在最后一段后面插入公式 E=mc^2」
3. **预期**：三张不同类型的卡片（「AI 格式 / AI 表格 / AI 公式」标签）：
   - 格式卡片：属性变更列表（如「加粗：关 → 开」「对齐：左对齐 → 居中」）
   - 表格卡片：迷你表格预览（≤5 数据行，超出显示「…」）
   - 公式卡片：KaTeX 预览 + LaTeX 原文
4. **格式**：接受 → 格式保留、卡片「已接受」；另开一次拒绝 → **段落格式逐项还原
   为应用前的值**（formatBefore 回写；Word 不把格式修改记为修订时走此路径，
   已跟踪时走 rejectAll）
5. **格式卡 §29 边界**：保持 pending → 切到 Word 窗口再切回任务窗格（触发对账）
   → 卡片**必须仍是「待处理」**（existence 模式：CC 存在即 pending，
   不得误判「已在 Word 中处理」）
6. **表格**：接受 → 表格保留（TableGrid 边框样式、表头行）；拒绝 → **整表删除**
   （跟踪与未跟踪两条路径都试：分别用「审阅→修订 开/关」状态各做一次）
7. **公式**：提案卡出现时文档必须保持不变；点接受后公式才写入，双击公式能进入
   Word 原生公式编辑器（OMML 对象）；点拒绝则始终不写入；
   「插入公式 \frac{a}{b}」分别试行内 / 独立成行；
   故意给非法 LaTeX（如 `\frac{`）→ 提示公式语法错误且**文档无任何变化**
8. **插入类原生审阅**：插入表格卡片保持 pending → 在 Word 审阅里手动接受该插入
   → 回焦点 → 卡片变「已在 Word 中处理」
9. **混合任务**：「全文纠错，并在最后一段后插一个总结表格」→ 文本 + 表格卡片
   都出现，汇总消息分类计数（「文本 2、表格 1」）
10. **降级路径**：低版本 Word（无 WordApi 1.3）→ 表格报
    `UNSUPPORTED_WORD_VERSION`（汇总附提示行），公式与格式不受影响；
    同一锚点连续两次插入 → 第二条因 `PENDING_EDIT_CONFLICT` 跳过

**测试 7 —— 空文档直接生成（锚点省略 = 文档末尾）**

1. 新建一个**空白** Word 文档，模式切到「**Agent**」
2. 输入「帮我画一个表格」
3. **预期**：AI **直接提交表格提案**（不再要求先输入文字）：卡片显示
   「文档末尾」位置 + 表格预览；接受 → 表格写入空文档
4. 再试「随便写两段人工智能简介」→ 「AI 段落」卡片（文本预览，≤3 段）；
   接受 → 多段文字以修订写入；拒绝 → 全部移除
5. 含内容文档的锚点回归：在已有段落上重复测试 6 → 仍走锚点段后插入
   （卡片显示「锚点段落之后」）
6. **表内段落锚点拒绝**：文档含已接受的表格后再要求「在 95% 那段后面
   插公式」→ 快照把单元格段落标注（表格内）；模型要么主动改插文档末尾，
   要么锚点被校验拒绝（错误回喂后改用文档末尾 / 正文段落）—— 绝不把
   内容插进单元格；应用失败时汇总消息带**具体原因**（不只列段落 id）

## 项目结构

```
frontend/          Vue 3 + Pinia + Vite 任务窗格
  src/services/word/    Office.js 全部封装（WordService/Selection/Document/
                        RangeLocator/PatchEngine/FormatEngine/InsertEngine/
                        FormulaOoxml/RevisionService/ContentControl）
  src/services/agent/   提案 → CapturedTarget 映射（ProposalMapper，纯函数）
  src/services/diff/    DiffEngine（diff-match-patch + EditOperation 折叠）
  src/services/api/     ChatApi(SSE) / EditApi / AgentApi(SSE 双通道)
  src/stores/           chat / edits / agent / document / settings
  src/commands/         QuickCommands / intent（§17 意图路由：batch > agent-tool > edit > chat）
  src/components/       Chat / Edit / Context / Settings
  tests/                vitest 单元测试（Diff §68 / 意图路由 / 提案映射 /
                        LaTeX→OMML / 段落 OOXML / SSE 帧 / 对账模式）
backend/           FastAPI（无数据库）
  app/api/              chat(SSE) / edit / agent(SSE) / health / config
  app/services/         Prompt / Context 渲染 / Edit JSON 校验 / agent_service
  app/llm/              LLMProvider → OpenAICompatibleProvider
  app/prompts/          chat.md / edit.md / agent.md（§36 规则 + 提案纪律）
  app/models/           chat / edit / agent（Pydantic 请求与事件模型）
  tests/                pytest 纯函数单测（提案校验 / 快照渲染 / 四类提案工具）
test-documents/    §67 的 8 个测试 docx（tools/generate_test_documents.py 生成）
tools/             测试文档生成脚本
```

## 常见问题（FAQ）

**Q1：任务窗格空白 / 加载失败？**
先确认 `npm run dev` 正常且证书已安装（`npm run dev-certs`）；
浏览器直接打开 `https://localhost:3000/taskpane.html` 应能渲染。
Word 桌面版需要完全退出后重开（sideload 注册的网络目录读取在启动时发生）。

**Q2：编辑卡片提示「当前 Word 版本不支持修订」？**
修订 API 需要 WordApi 1.6（较新的 Microsoft 365 桌面版）。
此时对话功能不受影响；升级 Office 或使用 Microsoft 365 最新版。

**Q3：提示 DOCUMENT_CHANGED？**
发送请求后目标文字被（你或协作者）改动了 —— 这是乐观锁在防止覆盖，
重新执行修改即可。不要试图绕过该检查。

**Q4：提示 RANGE_AMBIGUOUS / RANGE_NOT_FOUND？**
目标文本在文档中重复出现或已不存在。换用选区模式选中具体文字后重试。

**Q5：SSE 不出字、直接结束？**
看后端日志：多为 LLM_API_KEY / LLM_MODEL 未配置（`LLM_ERROR`），
或上游模型不可达 / 超时（`LLM_TIMEOUT`）。`/api/v1/health` 与 `/api/v1/config`
可快速验证连通性与模型名。

**Q6：能修改表格、图片、公式吗？**

分通道：

- **单段 Edit 通道**：仅支持普通文本（§38）——选区包含图片 / Content Control 等
  复杂对象时明确提示并拒绝，不会暴力修改 OOXML。
- **Agent 通道**：可以**插入**表格（≤20 行 × 8 列、纯文本单元格）与数学公式
  （LaTeX → Word 原生 OMML 公式，KaTeX 单行子集），并可修改**整段**字体 /
  段落格式（加粗、斜体、下划线、删除线、字体、字号、颜色、对齐）。
  均以修订 + 事务卡片写入，可逐条接受 / 拒绝。
  段内选区级格式化、表格样式定制（列宽 / 合并单元格）与已有表格 / 图片编辑
  仍在第一版范围外。

**Q7：聊天记录会保存吗？**
仅保存在当前任务窗格会话（Pinia 内存），关闭 Word 后消失（§41），后端无数据库。

**Q8：生产部署要注意什么？**
任务窗格与后端都必须 HTTPS；`BACKEND_CORS_ORIGINS` 必须收紧到 Add-in 实际
Origin（禁止 `*`）；manifest 中的 URL 换成生产域名并重新分发。

**Q9：`npm run dev` 报「Port 3000 is already in use」？**
manifest 中任务窗格地址固定为 `https://localhost:3000`（端口写死，`strictPort: true`）。
本机若已有其他 Add-in 开发服务占用 3000（可用 `netstat -ano | findstr :3000` 查到进程），
需要先停掉它；或临时换端口调试（`npx vite --port 3100`），但要让 Word 正常加载，
必须改 `vite.config.ts`、`manifest.xml` 两处端口并保持一致。

**Q10：在输入框打字能让 AI 修改文档吗？**

可以。输入区上方有**模式开关**（三档，记忆在本地）：

- **智能**（默认）：关键词路由（batch > agent-tool > edit > chat）——
  批量/全文指令（「全文纠错」「检查整篇文档的错别字」）→ Agent 通道（§51）；
  格式 / 插入类指令（「把这段加粗」「插入 3x2 表格」「插入公式 E=mc^2」）→
  Agent 通道（这些能力 Agent 独占）；含「修改/优化/改写/润色/缩写/扩写/重写/
  专业化/纠错/翻译并替换」关键词的非疑问指令 → 单段 Edit 通道；
  疑问句和普通对话 → Chat。
- **对话**：所有输入一律按聊天处理，绝不修改文档（即使打了「润色」）。
- **Agent**：所有输入进 Agent 通道（Microsoft Agent Framework）——读取全文快照 +
  当前光标/选区段落（模型据此理解「这段」的指代），**由模型自主决定**是直接回答
  还是通过 propose_edit / propose_format / insert_table / insert_formula /
  insert_paragraph 提交修改提案；提案仍逐条走完整安全管线后生成事务卡片。

疑问句在智能模式下不会被误路由（保守设计：批量修改绝不因误判而触发）。
快捷命令芯片自带类型（编辑/对话/Agent），不受模式开关影响。
**能力边界**：AI 可以修改**已有文字**，也可以在锚点段后或**文档末尾**插入
表格 / 公式 / 纯文字段落（插入类工具的锚点可省略 —— 空文档也能直接生成
内容，如「帮我画一个表格」「随便写两段简介」）。所有修改以 Word 修订写入、
可逐条接受/拒绝；模型不编造快照中不存在的内容，无把握时不提交修改。

**Q11：Agent 模式和普通修改有什么区别？会绕过安全机制吗？**
Agent（Microsoft Agent Framework）面向多目标任务：读全文快照、调用只读工具
（get_outline / read_paragraph / search_paragraphs）回答问题，并通过五类提案工具
（propose_edit / propose_format / insert_table / insert_formula / insert_paragraph）
提交批量提案。
**不会绕过任何安全机制**——LLM 永远碰不到 Word：每条提案由前端独立走
冲突检查与哈希乐观锁，N 条修改 = N 张独立事务卡片。文本、段落和公式均进入
Word 修订写入管线；公式与富段落使用完整 Flat OPC，避免裸 OOXML 片段导致的
临时文档损坏。所有卡片均可逐条接受/拒绝。
用户中途点「停止」时已收到的提案全部丢弃（§58：取消不允许产生文档修改）。
公式转换（LaTeX → OMML → Word Flat OPC 包）在前端纯函数完成，转换失败或用户拒绝时
**不触碰文档**。

**Q12：提示 AGENT_NOT_INSTALLED？**
后端未安装 Agent Framework。执行：
`cd backend && py -m pip install agent-framework-core agent-framework-openai`
后重启后端。未安装时对话与单段修改功能不受影响。

**Q13：AI 回复支持 Markdown 吗？**
支持。AI 回复按受限 Markdown 渲染（标题 / 列表 / 加粗斜体 / 行内代码 /
代码块 / 引用 / 分隔线 / 链接 / **GFM 表格** / **数学公式**）。公式用 KaTeX
渲染（`$行内$` 与 `$$块级$$`，LaTeX 语法），样式在 `ChatMessage.vue` 引入
`katex/dist/katex.min.css`。渲染器是自写的纯函数
（`frontend/src/utils/markdown.ts`）：块级语法在原文上解析、内容输出时整体
HTML 转义、只生成白名单标签、链接仅允许 http(s)；公式在转义前提取交给
KaTeX（其输出自带转义），无 XSS 面（含 `javascript:` 协议防护）。
用户消息按原文显示。

## 第三方许可

- **[mathml2omml](https://www.npmjs.com/package/mathml2omml)**（公式插入，
  MathML → OMML 转换）：**LGPL-3.0**。本项目以**未修改的 npm 依赖**形式引用
  （不 vendor 源码），仅为库的正常导入使用，未做任何修改。
- katex（公式渲染，MIT）、diff-match-patch（Diff，Apache-2.0）、
  Vue / Pinia / Vite / vitest（MIT）、FastAPI（MIT）、
  Microsoft Agent Framework（MIT）。

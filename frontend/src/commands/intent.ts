/**
 * 自由输入指令的意图路由（对应需求文档 §17）：
 *
 *   「如果用户指令包含：修改/优化/改写/润色/缩写/扩写/重写/专业化/
 *     翻译并替换/纠正，允许模型判断 intent = edit」
 *
 * 实现为客户端关键词启发式（第一版不引入额外的意图分类 LLM 调用），
 * 显式快捷按钮始终可用，避免完全依赖该启发式（§17 原文要求）。
 *
 * 疑问句不路由（「润色是什么意思？」「这段怎么改写更好？」是对话）。
 * 注意：「翻译」单独出现走对话（用户只想看译文）；「翻译并替换」才触发编辑。
 *
 * Agent 通道（§51）：批量范围标记（全文/整篇/批量…）+ 编辑或检查类动词
 * → isBatchIntent；格式 / 插入类请求（加粗、改字体、插入表格、插入公式…）
 * → isAgentToolIntent（Agent 的四类工具独占能力，Edit 通道无法完成）。
 * 路由优先级 batch > agent-tool > edit > chat。
 *
 * 模式开关覆盖启发式：用户可显式选择「对话」（仅聊天）或「Agent」
 * （任意指令进 Agent 通道，模型自主决定是否修改文档）—— 见 routeInput。
 */

/** §17 编辑关键词（含快捷命令中的「纠错」变体） */
const EDIT_KEYWORDS: readonly string[] = [
  "修改",
  "优化",
  "改写",
  "润色",
  "缩写",
  "扩写",
  "重写",
  "专业化",
  "翻译并替换",
  "纠正",
  "纠错",
];

/** 疑问标记：出现即视为对话意图 */
const QUESTION_MARKERS: readonly string[] = ["什么", "怎么", "如何", "为什么", "哪儿", "哪里", "吗", "呢"];

/** 批量范围标记（§17 扩展 / §51）：任务面向全文或多段 */
const BATCH_SCOPE_MARKERS: readonly string[] = [
  "全文",
  "整篇",
  "通篇",
  "批量",
  "全部",
  "所有",
  "各段",
  "每段",
  "每个段落",
];

/** 检查 / 排查类动词：与批量范围标记组合时指向 Agent 任务 */
const BATCH_CHECK_KEYWORDS: readonly string[] = [
  "检查",
  "找出",
  "排查",
  "哪些地方",
  "语病",
  "错别字",
  "错字",
  "统一",
  "改正",
  "修正",
];

/**
 * 判断自由输入是否应路由到 Edit 通道。
 * 空串、疑问句 → false；否则命中任一编辑关键词 → true。
 */
export function isEditIntent(text: string): boolean {
  const t = text.trim();
  if (!t) return false;
  // 疑问句优先：以 ？/? 结尾或含疑问词
  if (/[?？]$/.test(t)) return false;
  if (QUESTION_MARKERS.some((m) => t.includes(m))) return false;
  return EDIT_KEYWORDS.some((k) => t.includes(k));
}

/**
 * 判断自由输入是否应路由到 Agent 通道（批量编辑 / 文档问答，§51）。
 * 需同时满足：批量范围标记 +（编辑关键词或检查类动词），且非疑问句。
 * 例：「全文纠错」「检查整篇文档的错别字」「把全文的术语统一一下」。
 */
export function isBatchIntent(text: string): boolean {
  const t = text.trim();
  if (!t) return false;
  if (/[?？]$/.test(t)) return false;
  if (QUESTION_MARKERS.some((m) => t.includes(m))) return false;
  if (!BATCH_SCOPE_MARKERS.some((m) => t.includes(m))) return false;
  return (
    EDIT_KEYWORDS.some((k) => t.includes(k)) ||
    BATCH_CHECK_KEYWORDS.some((k) => t.includes(k))
  );
}

// ---- Agent 工具意图（格式 / 插入：Agent 通道独占能力） ----

/** 格式修改标记：propose_format 能力（字体/段落格式） */
const FORMAT_MARKERS: readonly string[] = [
  "加粗",
  "斜体",
  "倾斜",
  "下划线",
  "删除线",
  "字体",
  "字号",
  "颜色",
  "居中",
  "对齐",
  "行距",
  "格式",
];

/** 插入标记：insert_table / insert_formula / insert_paragraph 能力（短词覆盖其长组合） */
const INSERT_MARKERS: readonly string[] = [
  "插入表格",
  "画表格",
  "做表格",
  "生成表格",
  "建表格",
  "表格",
  "公式",
  "插入段落",
  "插入文字",
  "插入内容",
  "写一段",
  "写上",
  "补充一段",
  "新增段落",
];

/**
 * 判断自由输入是否为 Agent 工具意图（格式修改 / 插入表格 / 插入公式 /
 * 插入纯文字段落）。
 * 这些能力只有 Agent 通道具备（Edit 管线是纯文本替换）—— auto 模式下
 * 直接路由到 agent，让模型调 propose_format / insert_table /
 * insert_formula / insert_paragraph。
 * 疑问句 → false（「表格怎么插入？」是对话）。
 * 例：「把这段加粗」「把标题改成微软雅黑」「在第二段后插入 3x2 表格」
 * 「帮我写一段产品介绍」。
 */
export function isAgentToolIntent(text: string): boolean {
  const t = text.trim();
  if (!t) return false;
  if (/[?？]$/.test(t)) return false;
  if (QUESTION_MARKERS.some((m) => t.includes(m))) return false;
  return (
    FORMAT_MARKERS.some((m) => t.includes(m)) ||
    INSERT_MARKERS.some((m) => t.includes(m))
  );
}

// ---------------- 模式开关（对话 / 智能 / Agent） ----------------

/** 输入模式：对话（仅聊天）/ 智能（关键词路由，默认）/ Agent（模型自主决策） */
export type ChatMode = "auto" | "chat" | "agent";

/** 自由输入的最终通道 */
export type RouteTarget = "chat" | "edit" | "agent";

/**
 * 按模式路由自由输入：
 * - chat：一律对话（即使含「润色」等编辑关键词，用户已明确选择仅对话）；
 * - agent：一律 Agent 通道 —— 模型自主决定回答还是调用修改工具；
 * - auto（默认）：关键词启发式，优先级 batch > agent-tool > edit > chat
 *   （「把这段加粗并润色」→ agent，「润色这段话」→ edit，
 *   「表格怎么插入？」→ chat）。
 */
export function routeInput(text: string, mode: ChatMode): RouteTarget {
  if (mode === "chat") return "chat";
  if (mode === "agent") return "agent";
  if (isBatchIntent(text)) return "agent";
  if (isAgentToolIntent(text)) return "agent";
  if (isEditIntent(text)) return "edit";
  return "chat";
}

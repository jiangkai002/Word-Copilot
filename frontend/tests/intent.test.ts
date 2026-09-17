/**
 * §17 意图路由单元测试：自由输入包含修改类关键词 → Edit；疑问句 / 普通对话 → Chat；
 * 批量范围标记 + 编辑/检查动词 → Agent（§51，优先级最高）；
 * 格式 / 插入类请求 → Agent 工具意图（Agent 通道独占能力）。
 * 模式开关（routeInput）：chat 一律对话 / agent 一律 Agent / auto 启发式。
 */
import { describe, expect, it } from "vitest";
import { isAgentToolIntent, isBatchIntent, isEditIntent, routeInput } from "@/commands/intent";

describe("isEditIntent（§17 自由输入意图路由）", () => {
  it.each<[string, boolean]>([
    ["润色这段话", true],
    ["帮我扩写一下这段", true],
    ["把这段话改写得更正式", true],
    ["优化这一段", true],
    ["重写这段内容", true],
    ["纠正里面的错别字", true],
    ["纠错", true],
    ["专业化表达", true],
    ["翻译并替换成英文", true],
    ["修改这段话", true],
  ])("编辑指令「%s」→ %s", (text, expected) => {
    expect(isEditIntent(text)).toBe(expected);
  });

  it.each<[string, boolean]>([
    ["你好", false],
    ["这份文档讲了什么？", false],
    ["什么是润色", false],
    ["润色是什么意思？", false],
    ["这段话怎么改写更好？", false],
    ["为什么推荐微服务架构", false],
    ["帮我翻译这段话", false], // 「翻译」单独出现是对话（想看译文），「翻译并替换」才是编辑
    ["", false],
    ["   ", false],
  ])("对话/疑问「%s」→ %s", (text, expected) => {
    expect(isEditIntent(text)).toBe(expected);
  });

  it("结尾标点不影响疑问判断", () => {
    expect(isEditIntent("这段话怎么改写更好")).toBe(false); // 含「怎么」
    expect(isEditIntent("帮我改写这段，谢谢")).toBe(true);
  });
});

describe("isBatchIntent（§51 批量任务 → Agent 路由）", () => {
  it.each<[string, boolean]>([
    ["全文纠错", true],
    ["检查全文的错别字并修改", true],
    ["把全文的术语统一一下", true],
    ["通篇润色", true],
    ["批量修改所有段落", true],
    ["各段检查语病", true],
    ["全文里哪些地方有问题", true], // 陈述式排查指令（无疑问标记）
  ])("批量指令「%s」→ %s", (text, expected) => {
    expect(isBatchIntent(text)).toBe(expected);
  });

  it.each<[string, boolean]>([
    ["帮我润色这一段", false], // 无批量标记 → 单段 Edit
    ["总结全文", false], // 无编辑/检查动词 → Chat
    ["全文讲了什么？", false], // 疑问句
    ["全文哪些地方有语病？", false], // 疑问句（保守：不自动批量修改）
    ["你好", false],
    ["", false],
    ["   ", false],
  ])("非批量「%s」→ %s", (text, expected) => {
    expect(isBatchIntent(text)).toBe(expected);
  });

  it("优先级：批量指令同时命中 isEditIntent 时仍应路由 Agent", () => {
    expect(isBatchIntent("全文纠错")).toBe(true);
    expect(isEditIntent("全文纠错")).toBe(true); // 二者都命中，路由顺序由调用方决定
  });
});

describe("isAgentToolIntent（格式 / 插入请求 → Agent 工具路由）", () => {
  it.each<[string, boolean]>([
    ["把这段加粗", true],
    ["这一段设置成斜体", true],
    ["标题加下划线", true],
    ["把这些字加删除线", true],
    ["把正文改成微软雅黑字体", true],
    ["字号改成14", true],
    ["把这段文字颜色改成红色", true],
    ["这段居中", true],
    ["把这段改成右对齐", true],
    ["行距调成1.5", true],
    ["统一一下段落格式", true],
    ["在第二段后插入3x2表格", true],
    ["帮我画一个表格", true],
    ["做个两列的表格", true],
    ["生成一个两列的表格", true],
    ["在这段后面写公式 E=mc^2", true],
    ["插入公式", true],
    ["在段末插入数学公式", true],
    ["帮我写一段产品介绍", true], // insert_paragraph
    ["在文档末尾插入文字总结", true],
    ["随便写点什么", false], // 无标记词 → Agent 模式自会兜住，auto 下默认对话
    ["给我补充一段过渡文字", true],
    ["把这段后面新增段落说明", true],
    ["在文档末尾创建一级标题", true],
    ["添加一个二级小节标题", true],
    ["创建标题章节", true],
    ["新增章节‘实施方案’", true],
  ])("工具指令「%s」→ %s", (text, expected) => {
    expect(isAgentToolIntent(text)).toBe(expected);
  });

  it.each<[string, boolean]>([
    ["表格怎么插入？", false], // 疑问句 → 对话
    ["公式是什么意思", false],
    ["怎么把这段加粗", false], // 含「怎么」
    ["这份文档的公式对吗", false],
    ["标题章节怎么创建？", false],
    ["帮我润色这一段", false], // 纯文本编辑 → Edit 通道
    ["你好", false],
    ["", false],
    ["   ", false],
  ])("非工具指令「%s」→ %s", (text, expected) => {
    expect(isAgentToolIntent(text)).toBe(expected);
  });
});

describe("routeInput（模式开关）", () => {
  it("chat 模式：编辑关键词与批量指令也一律走对话", () => {
    expect(routeInput("帮我润色这一段", "chat")).toBe("chat");
    expect(routeInput("全文纠错", "chat")).toBe("chat");
    expect(routeInput("你好吗", "chat")).toBe("chat");
  });

  it("agent 模式：疑问句与普通输入也走 Agent（模型自主决策）", () => {
    expect(routeInput("这段话是什么意思？", "agent")).toBe("agent");
    expect(routeInput("帮我润色这一段", "agent")).toBe("agent");
    expect(routeInput("你好", "agent")).toBe("agent");
  });

  it("auto 模式：优先级 batch > agent-tool > edit > chat", () => {
    expect(routeInput("全文纠错", "auto")).toBe("agent");
    expect(routeInput("把这段加粗", "auto")).toBe("agent"); // agent-tool 命中
    expect(routeInput("创建标题章节", "auto")).toBe("agent");
    expect(routeInput("帮我润色这一段", "auto")).toBe("edit");
    expect(routeInput("润色是什么意思？", "auto")).toBe("chat");
    expect(routeInput("表格怎么插入？", "auto")).toBe("chat"); // 疑问句不路由
    expect(routeInput("你好", "auto")).toBe("chat");
    expect(routeInput("", "auto")).toBe("chat");
  });

  it("auto 模式：格式 + 编辑动词混合指令走 Agent（Edit 通道做不了格式）", () => {
    expect(routeInput("把这段加粗并润色", "auto")).toBe("agent");
    expect(routeInput("修改这段的字体和颜色", "auto")).toBe("agent");
  });
});

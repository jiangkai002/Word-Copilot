/**
 * Markdown 渲染器单元测试（utils/markdown）：
 * 安全（转义先行 / 链接白名单）+ 常用子集 + 流式半成品语法。
 */
import { describe, expect, it } from "vitest";
import { renderMarkdown } from "@/utils/markdown";

describe("renderMarkdown —— 安全", () => {
  it("HTML 全部转义（无 XSS 面）", () => {
    const html = renderMarkdown('<script>alert("x")</script>');
    expect(html).toContain("&lt;script&gt;");
    expect(html).not.toContain("<script");
  });

  it("链接仅允许 http(s)，javascript: 协议按原文渲染", () => {
    expect(renderMarkdown("[点我](https://example.com)")).toContain(
      '<a href="https://example.com" target="_blank" rel="noopener noreferrer">点我</a>',
    );
    const evil = renderMarkdown("[点我](javascript:alert(1))");
    expect(evil).not.toContain("<a ");
    expect(evil).toContain("[点我](javascript:alert(1))"); // 字面量
  });

  it("行内代码内的 HTML / 链接语法不生效", () => {
    const html = renderMarkdown("`<b>**不是加粗**</b>`");
    expect(html).toContain("<code>");
    expect(html).not.toContain("<strong>");
    expect(html).not.toContain("<b>");
  });
});

describe("renderMarkdown —— 块级", () => {
  it("标题 h1-h6", () => {
    expect(renderMarkdown("# 标题一")).toContain("<h1>标题一</h1>");
    expect(renderMarkdown("### 小节")).toContain("<h3>小节</h3>");
  });

  it("无序 / 有序列表（连续行合并）", () => {
    expect(renderMarkdown("- 甲\n- 乙")).toBe("<ul><li>甲</li><li>乙</li></ul>");
    expect(renderMarkdown("1. 第一步\n2. 第二步")).toBe(
      "<ol><li>第一步</li><li>第二步</li></ol>",
    );
  });

  it("代码块（闭合）", () => {
    const html = renderMarkdown("```python\nprint(1)\n```");
    expect(html).toContain("<pre><code");
    expect(html).toContain('data-lang="python"');
    expect(html).toContain("print(1)");
  });

  it("未闭合代码块（流式中）按代码块渲染", () => {
    const html = renderMarkdown("```\nlet x = 1;");
    expect(html).toContain("<pre><code");
    expect(html).toContain("let x = 1;");
  });

  it("引用块与分隔线", () => {
    expect(renderMarkdown("> 引用一行\n> 引用两行")).toContain("<blockquote>");
    expect(renderMarkdown("---")).toBe("<hr />");
  });

  it("段落内单换行 → <br>；空行分段", () => {
    expect(renderMarkdown("第一行\n第二行")).toBe("<p>第一行<br>第二行</p>");
    expect(renderMarkdown("第一段\n\n第二段")).toBe("<p>第一段</p><p>第二段</p>");
  });
});

describe("renderMarkdown —— 行内", () => {
  it("加粗 / 斜体 / 删除线", () => {
    expect(renderMarkdown("**加粗**")).toContain("<strong>加粗</strong>");
    expect(renderMarkdown("*斜体*")).toContain("<em>斜体</em>");
    expect(renderMarkdown("~~删除~~")).toContain("<del>删除</del>");
  });

  it("斜体不误伤数学表达式（词字符间的 * 不生效）", () => {
    const html = renderMarkdown("计算 3*4*5 的结果");
    expect(html).not.toContain("<em>");
  });

  it("行内代码优先于加粗（代码里的 ** 保持字面）", () => {
    const html = renderMarkdown("`a **b** c`");
    expect(html).toContain("<code>a **b** c</code>");
    expect(html).not.toContain("<strong>");
  });

  it("中文标点旁的斜体可生效（^|[非词字符] 边界）", () => {
    const html = renderMarkdown("。*强调*。");
    expect(html).toContain("<em>强调</em>");
  });
});

describe("renderMarkdown —— 边界", () => {
  it("空串 / 纯空白返回空串", () => {
    expect(renderMarkdown("")).toBe("");
    expect(renderMarkdown("   \n  ")).toBe("");
  });

  it("CRLF 归一为 LF", () => {
    expect(renderMarkdown("第一行\r\n第二行")).toBe("<p>第一行<br>第二行</p>");
  });

  it("流式半成品加粗按字面渲染（闭合后自然成型）", () => {
    expect(renderMarkdown("正在生成**加粗")).not.toContain("<strong>");
    expect(renderMarkdown("正在生成**加粗**")).toContain("<strong>加粗</strong>");
  });
});

describe("renderMarkdown —— 表格", () => {
  it("GFM 表格：表头 / 分隔行 / 表体", () => {
    const html = renderMarkdown("| 错误 | 修正 |\n|---|---|\n| 去去了 | 去了 |");
    expect(html).toContain("<table>");
    expect(html).toContain("<th>错误</th>");
    expect(html).toContain("<th>修正</th>");
    expect(html).toContain("<td>去了</td>");
  });

  it("列对齐（:--- / ---: / :---:）", () => {
    const html = renderMarkdown("| A | B | C |\n|:---|---:|:---:|\n| 1 | 2 | 3 |");
    expect(html).toContain('text-align:left');
    expect(html).toContain('text-align:right');
    expect(html).toContain('text-align:center');
  });

  it("单元格支持行内语法（加粗 / 行内代码）", () => {
    const html = renderMarkdown("| 说明 |\n|---|\n| **加粗** `code` |");
    expect(html).toContain("<strong>加粗</strong>");
    expect(html).toContain("<code>code</code>");
  });

  it("单元格内 HTML 被转义", () => {
    const html = renderMarkdown("| A |\n|---|\n| <b>x</b> |");
    expect(html).toContain("&lt;b&gt;");
    expect(html).not.toContain("<b>");
  });

  it("表体列数不齐：补空 / 截断", () => {
    const html = renderMarkdown("| A | B |\n|---|---|\n| 1 |");
    expect(html).toContain("<td>1</td><td></td>");
  });

  it("非表格（无分隔行）不误判", () => {
    const html = renderMarkdown("普通段落 a | b 的说明");
    expect(html).not.toContain("<table>");
  });

  it("以 | 开头但不成表格的行按段落渲染（不死循环）", () => {
    // 单列无分隔行 / 无法解析的表格行：必须终止并按普通段落输出
    const html = renderMarkdown("| 只有我 |");
    expect(html).toContain("<p>");
    expect(html).toContain("只有我");
    expect(html).not.toContain("<table>");
  });

  it("正文行 + --- 分隔线不误判为单列表格", () => {
    const html = renderMarkdown("一句话\n---");
    expect(html).not.toContain("<table>");
    expect(html).toContain("<hr />");
  });
});

describe("renderMarkdown —— 公式（KaTeX）", () => {
  it("行内公式 $…$ 渲染为 KaTeX 输出", () => {
    const html = renderMarkdown("质能方程 $E=mc^2$ 很有名");
    expect(html).toContain('class="katex"');
    expect(html).not.toContain("$E");
  });

  it("多行块级公式 $$…$$（displayMode）", () => {
    const html = renderMarkdown("$$\n\\int_0^1 x\\,dx = \\frac{1}{2}\n$$");
    expect(html).toContain("katex-display");
    expect(html).toContain('class="katex"');
  });

  it("整行 $$…$$ 也按块级公式渲染", () => {
    const html = renderMarkdown("$$x=1$$");
    expect(html).toContain("katex-display");
  });

  it("未闭合 $$（流式中）按公式渲染到底", () => {
    const html = renderMarkdown("$$\n\\frac{1}{2}");
    expect(html).toContain('class="katex"');
  });

  it("公式内容不被 HTML 转义破坏（< 完整进入 KaTeX，不产生裸标签）", () => {
    const html = renderMarkdown("当 $a<b$ 时成立");
    expect(html).toContain('class="katex"');
    // KaTeX 自行转义输出（<mo>&lt;</mo>）；关键是没有裸 <b> 标签产生
    expect(html).toContain("<mo>&lt;</mo>");
    expect(html).not.toContain("<b>");
  });

  it("公式内的 * 不被斜体化（转义前置提取）", () => {
    const html = renderMarkdown("卷积 $a*b*c$ 运算");
    expect(html).toContain('class="katex"');
    expect(html).not.toContain("<em>");
  });

  it("货币 $5 和 $20 不误判为公式（内容带空格）", () => {
    const html = renderMarkdown("价格 $5 和 $20 之间");
    expect(html).not.toContain('class="katex"');
    expect(html).toContain("$5");
  });

  it("行内代码里的 $ 不触发公式", () => {
    const html = renderMarkdown("`$x$` 是代码");
    expect(html).toContain("<code>$x$</code>");
    expect(html).not.toContain('class="katex"');
  });

  it("KaTeX 解析失败不抛异常（流式安全）", () => {
    expect(() => renderMarkdown("$\\unexpectedcmd{x}$")).not.toThrow();
  });
});

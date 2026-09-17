/**
 * InsertEngine 纯函数测试：纯文字段落 → OOXML 转换。
 *
 * 钉死的契约：
 * - 每行一个 <w:p>（<w:r><w:t xml:space="preserve">）—— 换行分段语义
 * - 文本 XML 转义（& < > " '）—— 防注入 / 防破坏 OOXML 结构
 * - 完整 Flat OPC 包（insertOoxml 稳定导入，不传裸 w:p 片段）
 */
import { describe, expect, it } from "vitest";
import { containsInlineFormula, escapeXmlText, headingToOoxml, paragraphsToOoxml } from "@/services/word/InsertEngine";

describe("escapeXmlText", () => {
  it("五个 XML 保留字符全部转义", () => {
    expect(escapeXmlText(`a<b>&"c'"d`)).toBe(`a&lt;b&gt;&amp;&quot;c&apos;&quot;d`);
  });

  it("普通中文 / LaTeX 符号原样保留", () => {
    expect(escapeXmlText("系统采用传统架构。")).toBe("系统采用传统架构。");
    expect(escapeXmlText("E=mc^2")).toBe("E=mc^2");
  });
});

describe("paragraphsToOoxml", () => {
  it("单行 → Flat OPC 包中的单个 w:p（含 xml:space）", () => {
    const ooxml = paragraphsToOoxml(["新段落"]);
    expect(ooxml).toMatch(/^<pkg:package xmlns:pkg=/);
    expect(ooxml).toContain('pkg:name="/word/document.xml"');
    expect(ooxml).toContain('<w:body><w:p><w:r><w:t xml:space="preserve">新段落</w:t></w:r></w:p></w:body>');
    expect(ooxml).not.toContain("<semantics");
  });

  it("多行 → 多个 w:p 依序拼接（换行分段语义）", () => {
    const ooxml = paragraphsToOoxml(["第一段", "第二段", "第三段"]);
    expect(ooxml.match(/<w:p>/g)).toHaveLength(3);
    // 顺序保持
    const texts = [...ooxml.matchAll(/<w:t xml:space="preserve">([^<]*)<\/w:t>/g)].map((m) => m[1]);
    expect(texts).toEqual(["第一段", "第二段", "第三段"]);
  });

  it("含 XML 保留字符的文本被转义（结构不被破坏）", () => {
    const ooxml = paragraphsToOoxml(['含 <标签> & "引号" 的文本']);
    expect(ooxml).toContain("含 &lt;标签&gt; &amp; &quot;引号&quot; 的文本");
    // 转义后结构仍平衡（w:t 内不含裸 <）
    expect(ooxml).not.toMatch(/<w:t[^>]*>[^<]*<[^/]/);
  });

  it("文字中的 $LaTeX$ 转为同段原生行内 OMML", () => {
    const lines = ["质能方程 $E=mc^2$ 表明质量与能量可以相互转化。"];
    const ooxml = paragraphsToOoxml(lines);
    expect(containsInlineFormula(lines)).toBe(true);
    expect(ooxml).toContain('<w:t xml:space="preserve">质能方程 </w:t>');
    expect(ooxml).toContain("<m:oMath");
    expect(ooxml).toContain('<w:t xml:space="preserve"> 表明质量与能量可以相互转化。</w:t>');
    expect(ooxml).not.toContain("$E=mc^2$");
  });

  it("文字中的非法公式在写入 Word 前失败", () => {
    expect(() => paragraphsToOoxml(["公式 $\\frac{$ 无效"])).toThrow();
  });
});

describe("headingToOoxml", () => {
  it.each([1, 2, 3, 9])("生成 Word 内置 Heading%s 样式", (level) => {
    const ooxml = headingToOoxml("第一章 <系统概述>", level);
    expect(ooxml).toContain(`<w:pStyle w:val="Heading${level}"/>`);
    expect(ooxml).toContain(`<w:outlineLvl w:val="${level - 1}"/>`);
    expect(ooxml).toContain("第一章 &lt;系统概述&gt;");
  });

  it("拒绝非法级别、空标题和多行标题", () => {
    expect(() => headingToOoxml("标题", 0)).toThrow("1-9");
    expect(() => headingToOoxml("   ", 1)).toThrow("非空的单行");
    expect(() => headingToOoxml("第一章\n正文", 1)).toThrow("非空的单行");
  });
});

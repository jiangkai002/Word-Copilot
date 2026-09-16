/**
 * FormulaOoxml 单元测试：LaTeX → OMML 转换链（纯函数，不触碰 Word 文档）。
 *
 * 钉形转换链关键结构：display 居中段落属性（w:jc）、行内无居中、
 * 分数 <m:f> / 上标 <m:sSup>、无 KaTeX/semantics/annotation 残留、
 * 以完整 Flat OPC 包承载 /word/document.xml、解析失败抛 FORMULA_INVALID。
 *
 * 结构铁律（真机验证）：**不生成 <m:oMathPara>** —— 段落被 Content Control
 * 包住后 Word 保存挂死（msoXXXX.tmp 损坏对话框的根源）；display 公式用
 * 居中 w:p + 行内 m:oMath 表达，m:oMath 是行内级内容（与 w:r 同级），
 * SDT 包裹 / 修订跟踪 / 保存全部正常。
 */
import { describe, expect, it } from "vitest";
import { latexToOoxml } from "@/services/word/FormulaOoxml";
import { CopilotError } from "@/utils/errors";

describe("latexToOoxml", () => {
  it("display=true：居中段落属性 + 行内 oMath（独立成行居中语义）", () => {
    const ooxml = latexToOoxml("E=mc^2", true);
    expect(ooxml).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>');
    expect(ooxml).toContain("<m:oMath");
    expect(ooxml).toContain("</m:oMath></w:p>");
  });

  it("任何模式都不生成 oMathPara（CC 包裹后 Word 保存挂死 —— 禁用）", () => {
    expect(latexToOoxml("E=mc^2", true)).not.toContain("oMathPara");
    expect(latexToOoxml("E=mc^2", false)).not.toContain("oMathPara");
  });

  it("display=false：行内公式，无居中段落属性", () => {
    const ooxml = latexToOoxml("E=mc^2", false);
    expect(ooxml).not.toContain("<w:pPr>");
    expect(ooxml).toContain("<m:oMath");
  });

  it("输出为 Word.js 官方格式的最小 Flat OPC 包（可直接 insertOoxml）", () => {
    const ooxml = latexToOoxml("a+b", true);
    expect(ooxml.startsWith("<pkg:package xmlns:pkg=")).toBe(true);
    expect(ooxml).toContain('pkg:name="/_rels/.rels"');
    expect(ooxml).toContain('Target="word/document.xml"');
    expect(ooxml).toContain('pkg:name="/word/document.xml"');
    expect(ooxml).toContain("<pkg:xmlData><w:document");
    expect(ooxml).toContain("<w:body><w:p>");
    expect(ooxml).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(ooxml.endsWith("</pkg:package>")).toBe(true);
  });

  it("\\frac{a}{b} 转换出 OMML 分数结构 <m:f>", () => {
    const ooxml = latexToOoxml("\\frac{a}{b}", true);
    expect(ooxml).toContain("<m:f");
    expect(ooxml).toContain("<m:num");
    expect(ooxml).toContain("<m:den");
  });

  it("x^{2} 转换出 OMML 上标结构 <m:sSup>", () => {
    const ooxml = latexToOoxml("x^{2}", false);
    expect(ooxml).toContain("<m:sSup");
  });

  it("无 KaTeX span / semantics / annotation 残留", () => {
    const ooxml = latexToOoxml("\\frac{a}{b}", true);
    expect(ooxml).not.toContain("katex");
    expect(ooxml).not.toContain("semantics");
    expect(ooxml).not.toContain("annotation");
    // annotation 携带的 LaTeX 源码回显也不得出现在输出中
    expect(ooxml).not.toContain("frac{a}{b}");
  });

  it("希腊字母 α 保留为 Unicode 字符（非实体引用）", () => {
    const ooxml = latexToOoxml("\\alpha", false);
    expect(ooxml).toContain("α");
  });

  it("非法 LaTeX 抛 FORMULA_INVALID（不产生任何 OOXML）", () => {
    expect(() => latexToOoxml("\\frac{", true)).toThrow(CopilotError);
    try {
      latexToOoxml("\\frac{", true);
    } catch (err) {
      expect((err as CopilotError).code).toBe("FORMULA_INVALID");
    }
  });
});

/**
 * LaTeX → Word OMML 公式转换（纯函数，vitest 可测）。
 *
 * 转换链：KaTeX(MathML 输出) → 剥 <semantics>/<annotation> → mathml2omml
 * (MathML → OMML) → 包进 <w:p> → 包装为 Word insertOoxml 要求的 Flat OPC 包。
 *
 * - KaTeX 输出的 MathML 带 <semantics> 包装与 <annotation>（LaTeX 源码回显），
 *   OMML 不认识这些元素 —— 按标签名容错剥离（非位置解析，结构漂移友好）。
 * - Word.js 的 insertOoxml 不能稳定导入裸 <w:p> 片段；桌面版 Word 会先生成
 *   mso*.tmp，再因它不是完整文档包而报「内容有问题」。因此最终输出严格使用
 *   Microsoft 官方示例采用的 pkg:package + /_rels/.rels + /word/document.xml。
 * - **不使用 <m:oMathPara>**：display 公式用居中 w:p + 行内 <m:oMath> 表达。
 *   实测（Word COM 真机验证）：段落被 Content Control 包住后，含 oMathPara 的
 *   裸片段会触发「无法打开文档 msoXXXX.tmp，内容有问题」；完整 Flat OPC
 *   下的行内 m:oMath 已通过 TrackAll + Content Control 保存 / 重开真机验证。
 *   视觉效果等价：公式独立成行且居中（w:jc 代替 oMathPara 的居中语义）。
 * - display=true：居中独立段落；display=false：普通段落（行内公式语义）。
 * - 解析 / 转换失败抛 CopilotError（FORMULA_INVALID / FORMULA_CONVERSION_ERROR），
 *   调用方在 Word.run 之外调用 —— 纯失败不触碰文档。
 */
import katex from "katex";
import { mml2omml } from "mathml2omml";
import { CopilotError } from "@/utils/errors";
import { bodyContentToFlatOpc } from "./FlatOpc";

const MATH_BLOCK_RE = /<math[\s\S]*?<\/math>/;
const ANNOTATION_RE = /<annotation[\s\S]*?<\/annotation>/g;

/** LaTeX → 可嵌入 w:p 的行内 OMML（m:oMath）。 */
export function latexToOmml(latex: string, display = false): string {
  let mathml: string;
  try {
    const html = katex.renderToString(latex, {
      output: "mathml",
      displayMode: display,
      throwOnError: true,
      strict: "ignore",
    });
    const match = MATH_BLOCK_RE.exec(html);
    if (!match) {
      throw new CopilotError(
        "FORMULA_CONVERSION_ERROR",
        "公式渲染输出不含 MathML 结构，无法转换为 Word 公式。",
      );
    }
    mathml = match[0];
  } catch (err) {
    if (err instanceof CopilotError) throw err;
    const message = err instanceof Error ? err.message : String(err);
    throw new CopilotError("FORMULA_INVALID", `LaTeX 语法无法解析：${message}`);
  }

  const cleaned = mathml.replace(ANNOTATION_RE, "").replace(/<\/?semantics>/g, "");
  if (!/<math[\s>]/.test(cleaned)) {
    throw new CopilotError("FORMULA_CONVERSION_ERROR", "公式 MathML 结构清洗失败。");
  }

  let omml: string;
  try {
    omml = mml2omml(cleaned);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    throw new CopilotError("FORMULA_CONVERSION_ERROR", `公式无法转换为 Word 格式（OMML）：${message}`);
  }
  if (!/<m:oMath[\s>]/.test(omml)) {
    throw new CopilotError("FORMULA_CONVERSION_ERROR", "公式 OMML 转换结果不完整。");
  }

  return omml;
}

/** LaTeX → 可直接传给 Range/Body.insertOoxml 的 Flat OPC OMML 文档包。 */
export function latexToOoxml(latex: string, display: boolean): string {
  const omml = latexToOmml(latex, display);
  const pPr = display ? '<w:pPr><w:jc w:val="center"/></w:pPr>' : "";
  const paragraph = `<w:p>${pPr}${omml}</w:p>`;
  return bodyContentToFlatOpc(paragraph);
}

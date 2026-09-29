/**
 * 文本规范化与哈希（对应需求文档 §49）。
 *
 * normalizeText 只统一 Word 的行尾（CRLF / CR / 软换行 VT），
 * 不做 trim —— 否则定位可能出错。
 */
import { sha256Hex } from "./crypto";

/** Word 行尾形式统一为 "\n"：\r\n、\r（段落标记）、\v（软换行） */
export function normalizeText(text: string): string {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\v/g, "\n");
}

/** 计算 textHash：SHA-256(normalizeText(text)) 的十六进制 */
export async function textHash(text: string): Promise<string> {
  return sha256Hex(normalizeText(text));
}

/** 规范化后是否相等（用于乐观锁比较之外的文本比较） */
export function normalizeEqual(a: string, b: string): boolean {
  return normalizeText(a) === normalizeText(b);
}

/**
 * 去掉结尾的 Word 结构标记（搜索 / 锚点哈希时用）。
 * \x07 是 Word 表格单元格结束标记，部分宿主会把它附在 Range.text 末尾。
 */
export function stripTrailingMarks(text: string): string {
  return text.replace(/[\r\v\n\x07]+$/, "");
}

/** 去掉开头的段落标记 / 软换行 */
export function stripLeadingMarks(text: string): string {
  return text.replace(/^[\r\v\n]+/, "");
}

/** 把规范化文本(\n)转回 Word 文本形式(\r)——插入 Word 时使用 */
export function toWordText(text: string): string {
  return text.replace(/\n/g, "\r");
}

/**
 * “接受全部修订后”文本的比对（PatchEngine 校验用）。
 *
 * Range.getReviewedText 与 Range.text 对结构标记的渲染不一致：
 * 部分宿主的 getReviewedText 会在控件内容前附加 '<' 等结构标记，
 * 并把原生公式中的字母转为 Unicode 数学字母。只清理范围边界的标记，
 * 保留正文中的比较符，再折叠数学字母后比较。
 */
export function reviewedTextEquals(final: string, expected: string): boolean {
  if (final === expected) return true;
  // 宿主添加的结构标记仅出现在范围边界；正文中的 <、>（如 α_i > 0）必须保留。
  const stripMarks = (s: string): string => {
    let result = s.replace(/\x07+$/, "");
    if (!expected.startsWith("<")) result = result.replace(/^<+/, "");
    if (!expected.endsWith(">")) result = result.replace(/>+$/, "");
    return result;
  };
  // Word 会把原生公式内的 ASCII 字母/数字转成 Unicode 数学字母，
  // 例如 d_i → 𝑑_𝑖。只折叠数学字母区，避免 NFKC 改写正文中的其他字符。
  const normalizeMathLetters = (s: string): string =>
    s.replace(/[\u{1D400}-\u{1D7FF}]/gu, (char) => char.normalize("NFKC"));
  return normalizeMathLetters(stripMarks(final)) === normalizeMathLetters(expected);
}

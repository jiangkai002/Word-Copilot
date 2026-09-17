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
 * 部分宿主的 getReviewedText 会包含控件 / 表格单元格等结构标记字符
 * （实测 Word 桌面版会在控件内容前附加字面 '<'），而 Range.text 不包含。
 * 若期望文本本身不含标记字符，则剥离两侧标记字符后再比较，
 * 避免把正确应用误判为失败而回滚。
 */
export function reviewedTextEquals(final: string, expected: string): boolean {
  if (final === expected) return true;
  if (/[<>]/.test(expected)) return false; // 期望文本含 '<'/'>' 时不容忍（防掩盖真实差异）
  const stripMarks = (s: string): string => s.replace(/[<>\x07]/g, "");
  return stripMarks(final) === stripMarks(expected);
}

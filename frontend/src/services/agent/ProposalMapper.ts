/**
 * 提案 → CapturedTarget 映射（§51，纯函数，供单测）。
 *
 * Agent 提案（六类判别联合）：
 * - text / format：paragraph_id 指向目标段落（改文字 / 改格式都作用于该段）
 * - insert-table / insert-formula / insert-paragraph / insert-heading：anchor_paragraph_id
 *   指向锚点段落（内容插其后）；anchor 为 null / 缺省 = 文档末尾（无目标段）
 *
 * 各类提案均从快照段落重新构造与 captureEditTarget 段落模式完全一致的
 * CapturedTarget：
 * - text / textHash / locator.originalText 取快照段落文本（前端为权威原文）
 * - locator.paragraphId = 段落 uniqueLocalId（RangeLocator tier-1）
 * - prefix / suffix = 快照前后段落文本（edit 请求上下文 / 消歧参考）
 *
 * 提案应用前的回显校验（text 类：normalizeEqual(proposal.original_text,
 * target.text)）由 edits store 的 applyProposal 执行 —— 因此本函数必须使用
 * 快照段落文本而非 proposal.original_text，该校验才有意义。
 */
import type { AgentProposal } from "@/models/Api";
import type { SnapshotParagraph } from "@/services/word/DocumentService";
import type { CapturedTarget } from "@/services/word/SelectionService";
import { textHash } from "@/utils/text";

/**
 * 提案指向的段落 id：text / format 为目标段，insert-* 为锚点段；
 * insert-* 的文档末尾模式（anchor 省略 / null）返回 null。
 */
export function proposalTargetId(proposal: AgentProposal): string | null {
  if ("paragraph_id" in proposal) return proposal.paragraph_id;
  return proposal.anchor_paragraph_id ?? null;
}

export async function proposalToCapturedTarget(
  proposal: AgentProposal,
  paragraphs: readonly SnapshotParagraph[],
): Promise<CapturedTarget | null> {
  const targetId = proposalTargetId(proposal);
  if (targetId === null) return null; // 文档末尾模式 —— 由调用方按无锚点处理
  const index = paragraphs.findIndex((p) => p.id === targetId);
  if (index < 0) return null;
  const paragraph = paragraphs[index];
  const hash = await textHash(paragraph.text);
  return {
    kind: "paragraph",
    text: paragraph.text,
    textHash: hash,
    prefix: paragraphs[index - 1]?.text ?? "",
    suffix: paragraphs[index + 1]?.text ?? "",
    locator: {
      paragraphId: paragraph.id,
      originalText: paragraph.text,
      textHash: hash,
    },
  };
}

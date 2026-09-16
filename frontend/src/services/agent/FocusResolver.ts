/**
 * Agent 模式焦点解析（§51，纯函数）：
 * 把 SelectionService 捕获的光标 / 选区焦点解析为随请求发送的 focus 载荷。
 *
 * 段落 id / 文本一律以快照为权威（与 ProposalMapper 同一原则）：
 * - 光标段 uniqueLocalId 命中快照 id → 直接采用；
 * - 否则按段落文本唯一匹配回退（快照 id 缺失时前端用 para-{i} 兜底）；
 * - 都无法解析（如光标段被 50k 截断排除）→ 返回 null，不发送焦点。
 */
import type { AgentFocusPayload } from "@/models/Api";
import type { SnapshotParagraph } from "@/services/word/DocumentService";

export interface AgentFocusInput {
  paragraphId: string | null;
  paragraphText: string;
  selectedText: string | null;
}

export function resolveFocus(
  focus: AgentFocusInput,
  paragraphs: readonly SnapshotParagraph[],
): AgentFocusPayload | null {
  let match: SnapshotParagraph | undefined = focus.paragraphId
    ? paragraphs.find((p) => p.id === focus.paragraphId)
    : undefined;
  if (!match) {
    const hits = paragraphs.filter((p) => p.text === focus.paragraphText);
    if (hits.length === 1) match = hits[0]; // 重复段落无法消歧 → 不发焦点
  }
  if (!match) return null;
  return {
    paragraph_id: match.id,
    paragraph_text: match.text,
    selected_text: focus.selectedText,
  };
}

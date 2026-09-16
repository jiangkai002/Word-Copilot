/**
 * Agent Store（§51 / §58 / §60）：批量编辑 + 文档问答的运行编排。
 *
 * 流程：文档快照 → 隐私披露（§60，聊天区明示读取范围）→ SSE
 * （token 流式气泡 + proposal 收集）→ 流结束（done / error 均适用）
 * → 逐条经 editsStore.applyProposal 走独立确定性管线（§15 乐观锁 +
 * Track Changes 事务卡片，逐条接受 / 拒绝）。
 *
 * §58：用户停止（AbortController）不允许产生 Word 修改 —— 已收到的
 * 提案直接丢弃并说明；流自身出错（超时 / LLM 错误）时，已收到的提案
 * 仍逐条应用（每条独立校验、以修订写入、可逐条拒绝）并在汇总中注明。
 *
 * 运行状态复用 chatStore.streaming / abortController，驱动输入区禁用
 * 与「停止」按钮（chat.stop() 与 agent 停止走同一个 AbortController）。
 */
import { defineStore } from "pinia";
import { ref } from "vue";
import type { AgentFocusPayload, AgentProposal, ProposalKind } from "@/models/Api";
import type { DocumentSnapshot } from "@/services/word/DocumentService";
import { agentApi } from "@/services/api/AgentApi";
import { proposalTargetId } from "@/services/agent/ProposalMapper";
import { resolveFocus } from "@/services/agent/FocusResolver";
import { documentService } from "@/services/word/DocumentService";
import { selectionService } from "@/services/word/SelectionService";
import { toCopilotError } from "@/utils/errors";
import { logger } from "@/utils/logger";
import { useChatStore } from "./chat";
import { useDocumentStore } from "./document";
import { useEditStore } from "./edits";

const KIND_LABELS: Record<ProposalKind, string> = {
  text: "文本",
  format: "格式",
  "insert-table": "表格",
  "insert-formula": "公式",
  "insert-paragraph": "段落",
};

export const useAgentStore = defineStore("agent", () => {
  /** Agent 任务进行中（含提案应用阶段） */
  const working = ref(false);

  /**
   * 运行 Agent 任务（文档问答 / 修改提案；模型自主决策）。
   * 入口：ChatInput 按 mode 路由（智能模式的 isBatchIntent 或 Agent 模式的一切指令）、
   * 或 QuickCommands（kind === "agent"）。
   */
  async function runAgent(instruction: string): Promise<void> {
    const chatStore = useChatStore();
    const documentStore = useDocumentStore();
    const editsStore = useEditStore();

    if (working.value || chatStore.streaming || editsStore.working) {
      chatStore.addSystem("正在处理上一个任务，请稍候。");
      return;
    }
    if (!documentStore.hostReady) {
      chatStore.addSystem("Word 宿主未就绪，无法读取文档。", "WORD_API_ERROR");
      return;
    }

    // ---- 1. 快照（§51：agent 读取全文 —— §60 要求发送前披露）----
    let snapshot: DocumentSnapshot;
    try {
      snapshot = await documentService.getDocumentSnapshot();
    } catch (err) {
      const ce = toCopilotError(err);
      chatStore.addSystem(ce.message, ce.code);
      return;
    }
    if (snapshot.paragraphs.length === 0 && snapshot.truncated) {
      // 单段超长导致快照为空：无内容可参考也无法提案
      chatStore.addSystem("文档单段超长，无法生成快照，无法执行该任务。");
      return;
    }

    // ---- 2. 焦点（Agent 模式「这段」的指代；尽力而为，失败不阻断）----
    let focus: AgentFocusPayload | null = null;
    if (snapshot.paragraphs.length > 0) {
      try {
        const raw = await selectionService.captureAgentFocus();
        if (raw) focus = resolveFocus(raw, snapshot.paragraphs);
      } catch {
        /* 焦点获取失败仅失去指代信息，任务照常 */
      }
    }

    const charCount = snapshot.paragraphs.reduce((sum, p) => sum + p.text.length, 0);
    if (snapshot.paragraphs.length === 0) {
      // 空文档：无内容可读，但仍可在文档末尾插入新内容（表格 / 公式 / 段落）
      chatStore.addSystem("文档为空：AI 没有可参考的文档内容，可直接对话，也可让 AI 在文档末尾写入新内容（表格 / 公式 / 文字段落）。");
    } else {
      chatStore.addSystem(
        `已读取全文 ${snapshot.paragraphs.length} 段（约 ${charCount} 字）` +
          (snapshot.truncated ? "，超长已截断，仅处理前一部分" : "") +
          "。",
      );
    }

    // ---- 2. SSE 流：token 进流式气泡，proposal 收集 ----
    const history = chatStore.recentHistory;
    chatStore.pushMessage({ type: "user", content: instruction });
    const assistant = chatStore.pushMessage({ type: "assistant", content: "", streaming: true });
    chatStore.streaming = true;
    const controller = new AbortController();
    chatStore.abortController = controller;
    working.value = true;

    const proposals: AgentProposal[] = [];
    let streamError = false;
    let aborted = false;
    /** done 帧带回的提案总数 —— 与本地解析数不一致 = 有帧被丢弃（不可静默） */
    let sentProposalCount: number | null = null;

    try {
      await agentApi.streamAgent(
        {
          conversation_id: chatStore.conversationId,
          instruction,
          history,
          // wire 转换：inTable → in_table（快照本地对象仍保留 camelCase 供 apply 管线用）
          snapshot: {
            outline: snapshot.outline,
            truncated: snapshot.truncated,
            paragraphs: snapshot.paragraphs.map((p) => ({
              id: p.id,
              text: p.text,
              style: p.style,
              level: p.level,
              in_table: p.inTable ?? false,
            })),
          },
          focus,
        },
        {
          onToken: (t) => {
            assistant.content += t;
          },
          onProposal: (p) => {
            proposals.push(p);
          },
          onError: (code, message) => {
            streamError = true;
            assistant.errorCode = code;
            assistant.content = assistant.content || message;
          },
          onDone: (event) => {
            if (typeof event.proposal_count === "number") sentProposalCount = event.proposal_count;
          },
        },
        controller.signal,
      );
    } catch (err) {
      const ce = toCopilotError(err);
      if (ce.code === "ABORTED") {
        aborted = true;
        assistant.content = assistant.content || "（已停止生成）";
      } else {
        streamError = true;
        assistant.errorCode = ce.code;
        assistant.content = assistant.content || ce.message;
        logger.error("agent 失败：", ce.code, ce.message);
      }
    } finally {
      assistant.streaming = false;
      chatStore.streaming = false;
      chatStore.abortController = null;
      working.value = false;
    }

    // ---- 3. 应用提案 ----
    if (sentProposalCount !== null && sentProposalCount > proposals.length) {
      // 帧被丢弃（畸形 / 前端不认识的提案类型）—— 死寂是最差的失败形态，必须明示
      chatStore.addSystem(
        `后端提交了 ${sentProposalCount} 条修改建议，任务窗格只识别出 ${proposals.length} 条` +
          "（其余提案帧格式不识别，被丢弃）。任务窗格可能加载了旧版代码 —— 请关闭并重新打开任务窗格后重试。",
      );
    }
    if (proposals.length === 0) return;

    if (aborted) {
      // §58：用户取消不产生 Word 修改 —— 收到的提案直接丢弃
      chatStore.addSystem(`已停止：收到 ${proposals.length} 条修改建议，均未写入文档。`);
      return;
    }
    if (!documentStore.revisionSupported) {
      // §28 降级：问答可用，修改无法写入
      chatStore.addSystem(
        `当前 Word 版本不支持修订（需 WordApi 1.6）：已收到 ${proposals.length} 条修改建议，` +
          "无法写入文档，仅对话可用。",
        "UNSUPPORTED_WORD_VERSION",
      );
      return;
    }

    working.value = true;
    try {
      let applied = 0;
      const appliedByKind = new Map<ProposalKind, number>();
      const skipped: string[] = [];
      const failed: { label: string; message: string }[] = [];
      let unsupportedHint = false;
      for (const proposal of proposals) {
        const result = await editsStore.applyProposal(instruction, proposal, snapshot.paragraphs);
        if (result.ok) {
          applied += 1;
          appliedByKind.set(proposal.kind, (appliedByKind.get(proposal.kind) ?? 0) + 1);
        } else if (result.code === "NO_CHANGE" || result.code === "PENDING_EDIT_CONFLICT") {
          skipped.push(proposalTargetId(proposal) ?? "文档末尾");
        } else {
          // 失败必须带原因 —— 只列段落 id 的「N 条失败」无法定位问题
          const label = proposalTargetId(proposal) ?? "文档末尾";
          failed.push({ label, message: result.message });
          if (result.code === "UNSUPPORTED_WORD_VERSION") unsupportedHint = true;
        }
      }
      const kindDetail = [...appliedByKind.entries()]
        .map(([kind, count]) => `${KIND_LABELS[kind]} ${count}`)
        .join("、");
      const parts = [
        `共 ${proposals.length} 条建议：${applied} 条已写入待审` +
          (kindDetail ? `（${kindDetail}）` : ""),
      ];
      if (skipped.length > 0) parts.push(`${skipped.length} 条跳过（无变化或已有待处理修改）`);
      if (failed.length > 0) {
        const detail = failed
          .map((f) => {
            const where = f.label === "文档末尾" ? "" : `段落 ${f.label}：`;
            const message = f.message.length > 60 ? `${f.message.slice(0, 60)}…` : f.message;
            return `${where}${message}`;
          })
          .join("；");
        parts.push(`${failed.length} 条失败（${detail}）`);
      }
      if (unsupportedHint) parts.push("部分建议需要更高版本的 Word 支持");
      if (streamError) parts.push("（任务中途出错，以上为中断前收到的提案）");
      chatStore.addSystem(parts.join("，") + "。");
    } finally {
      working.value = false;
    }
  }

  /** 停止（§58）：与 chat 停止按钮共用 AbortController */
  function stop(): void {
    const chatStore = useChatStore();
    chatStore.abortController?.abort();
  }

  return { working, runAgent, stop };
});

/**
 * Edit Store（对应需求文档 §23 / §56 / §57 / §78–§80）：
 * AI 编辑事务的生命周期：创建 → pending → accept / reject / regenerate / 对账。
 *
 * 关键约束：
 * - §57：同一范围内已有 pending 事务时拒绝新修改（PENDING_EDIT_CONFLICT）
 * - §56：重新生成复用已保存的 RangeLocator，光标移动不影响找回目标
 * - §29：用户可能用 Word 原生审阅处理修订 —— reconcile 对账而非报错
 */
import { defineStore } from "pinia";
import { ref, computed } from "vue";
import type { EditTransaction } from "@/models/EditTransaction";
import type {
  FormulaInsertPlan,
  FormatChanges,
  FormatPlan,
  HeadingInsertPlan,
  ParagraphInsertPlan,
  TableInsertPlan,
} from "@/models/EditPlan";
import type { EditPlan } from "@/models/EditPlan";
import type {
  AgentFormatProposalEvent,
  AgentFormulaProposalEvent,
  AgentHeadingProposalEvent,
  AgentParagraphProposalEvent,
  AgentProposal,
  AgentTableProposalEvent,
  AgentTextProposalEvent,
  EditRequestPayload,
} from "@/models/Api";
import type { CapturedTarget } from "@/services/word/SelectionService";
import type { SnapshotParagraph } from "@/services/word/DocumentService";
import type { TransactionRef, TransactionOutcome } from "@/services/word/RevisionService";
import { editApi } from "@/services/api/EditApi";
import { proposalToCapturedTarget } from "@/services/agent/ProposalMapper";
import { diffEngine } from "@/services/diff/DiffEngine";
import { formatEngine } from "@/services/word/FormatEngine";
import { headingToOoxml, insertEngine, paragraphsToOoxml } from "@/services/word/InsertEngine";
import { latexToOoxml } from "@/services/word/FormulaOoxml";
import { patchEngine, type PatchOutcome } from "@/services/word/PatchEngine";
import { revisionService } from "@/services/word/RevisionService";
import { WordService } from "@/services/word/WordService";
import { editControlTag, nextTransactionId } from "@/utils/id";
import { CopilotError, toCopilotError } from "@/utils/errors";
import { normalizeEqual } from "@/utils/text";
import { logger } from "@/utils/logger";
import { useChatStore } from "./chat";
import { useDocumentStore } from "./document";

/** 应用一条已生成修改的结果（agent 批量 / 单段编辑共用） */
export type ApplyEditResult =
  | { ok: true }
  | { ok: false; code: string; message: string };

type InsertEngineCompat = {
  applyFormulaInsertPlan?: (plan: FormulaInsertPlan) => Promise<PatchOutcome>;
  commitFormulaInsertPlan?: (plan: FormulaInsertPlan) => Promise<PatchOutcome>;
  applyParagraphInsertPlan?: (plan: ParagraphInsertPlan) => Promise<PatchOutcome>;
  commitParagraphInsertPlan?: (plan: ParagraphInsertPlan) => Promise<PatchOutcome>;
  applyHeadingInsertPlan?: (plan: HeadingInsertPlan) => Promise<PatchOutcome>;
  commitHeadingInsertPlan?: (plan: HeadingInsertPlan) => Promise<PatchOutcome>;
};

/** 新旧前端模块在 Vite HMR / Office 缓存窗口内共存时兼容两代方法名。 */
function applyFormulaInsertCompat(plan: FormulaInsertPlan): Promise<PatchOutcome> {
  const engine = insertEngine as unknown as InsertEngineCompat;
  const apply = engine.applyFormulaInsertPlan ?? engine.commitFormulaInsertPlan;
  if (!apply) {
    throw new CopilotError("FRONTEND_VERSION_MISMATCH", "InsertEngine 公式插入方法不可用");
  }
  return apply.call(engine, plan);
}

function applyParagraphInsertCompat(plan: ParagraphInsertPlan): Promise<PatchOutcome> {
  const engine = insertEngine as unknown as InsertEngineCompat;
  const apply = engine.applyParagraphInsertPlan ?? engine.commitParagraphInsertPlan;
  if (!apply) {
    throw new CopilotError("FRONTEND_VERSION_MISMATCH", "InsertEngine 段落插入方法不可用");
  }
  return apply.call(engine, plan);
}

function applyHeadingInsertCompat(plan: HeadingInsertPlan): Promise<PatchOutcome> {
  const engine = insertEngine as unknown as InsertEngineCompat;
  const apply = engine.applyHeadingInsertPlan ?? engine.commitHeadingInsertPlan;
  if (!apply) {
    throw new CopilotError("FRONTEND_VERSION_MISMATCH", "InsertEngine 标题插入方法不可用");
  }
  return apply.call(engine, plan);
}

export const useEditStore = defineStore("edits", () => {
  /** id → 事务（插入顺序即 order） */
  const transactions = ref<Record<string, EditTransaction>>({});
  const order = ref<string[]>([]);
  /** 是否有编辑流程进行中（请求 / diff / patch 任一阶段） */
  const working = ref(false);

  const pendingList = computed<EditTransaction[]>(() =>
    order.value
      .map((id) => transactions.value[id])
      .filter((tx): tx is EditTransaction => !!tx && tx.status === "pending"),
  );

  function getTransaction(id: string): EditTransaction | undefined {
    return transactions.value[id];
  }

  function registerTransaction(tx: EditTransaction): void {
    transactions.value[tx.id] = tx;
    order.value.push(tx.id);
  }

  /**
   * §57：同一范围待处理冲突检测。
   * - 原文哈希一致 → 同一目标文本（即使文本在别处重复出现也视为冲突，保守拒绝）
   * - pending 为段落级事务且新目标位于同段落 → 范围必然重叠
   * - locator 为 null（文档末尾插入）→ 无目标段落，跳过检测
   */
  function findPendingConflict(locator: EditTransaction["locator"]): EditTransaction | null {
    if (locator === null) return null;
    for (const tx of pendingList.value) {
      if (tx.locator && tx.locator.textHash === locator.textHash) return tx;
      if (
        tx.locator &&
        tx.targetKind === "paragraph" &&
        tx.locator.paragraphId &&
        tx.locator.paragraphId === locator.paragraphId
      ) {
        return tx;
      }
    }
    return null;
  }

  // ---------------- 编辑主流程（§56） ----------------

  /**
   * 入口：捕获当前目标（选区优先，否则光标段落）并执行 AI 修改。
   */
  async function sendEdit(instruction: string): Promise<void> {
    const chatStore = useChatStore();
    const documentStore = useDocumentStore();

    if (working.value) {
      chatStore.addSystem("正在处理上一个修改，请稍候。");
      return;
    }
    if (!documentStore.hostReady) {
      chatStore.addSystem("Word 宿主未就绪，无法修改文档。", "WORD_API_ERROR");
      return;
    }
    if (!documentStore.revisionSupported) {
      // §28：修订能力缺失时优雅降级 —— 不抛异常，提示后返回
      chatStore.addSystem(
        "当前 Word 版本不支持修订（需 WordApi 1.6 / Microsoft 365 较新版本），仅支持对话，无法执行修改。",
        "UNSUPPORTED_WORD_VERSION",
      );
      return;
    }

    let target: CapturedTarget;
    try {
      target = await documentStore.captureEditTarget();
    } catch (err) {
      const ce = toCopilotError(err);
      chatStore.addSystem(ce.message, ce.code);
      return;
    }

    await applyEditFlow(instruction, target, { regenerateCount: 0, avoidTexts: [] });
  }

  interface EditFlowOptions {
    regenerateCount: number;
    avoidTexts: string[];
    regenerationFeedback?: string;
  }

  /**
   * 请求 → 客户端 Diff → PatchEngine 应用 → 登记事务 → 展示编辑卡片。
   */
  async function applyEditFlow(
    instruction: string,
    target: CapturedTarget,
    options: EditFlowOptions,
  ): Promise<void> {
    const chatStore = useChatStore();
    working.value = true;
    try {
      // §57 冲突检查
      const conflict = findPendingConflict(target.locator);
      if (conflict) {
        chatStore.addSystem(
          `该内容已有待处理修改（${conflict.id}），请先接受或拒绝后再执行新的修改。`,
          "PENDING_EDIT_CONFLICT",
        );
        return;
      }

      // ---- 1. 请求后端（返回 original_text / new_text）----
      const payload: EditRequestPayload = {
        conversation_id: chatStore.conversationId,
        instruction,
        target: { type: target.kind, text: target.text, text_hash: target.textHash },
        context: {
          previous: target.prefix || undefined,
          next: target.suffix || undefined,
        },
        regenerate_count: options.regenerateCount,
        avoid_texts: options.avoidTexts.length > 0 ? options.avoidTexts : undefined,
        regeneration_feedback: options.regenerationFeedback?.trim() || undefined,
      };
      const response = await editApi.requestEdit(payload);

      // ---- 2. 校验：后端必须原样回显原文，否则 diff 坐标系无效 ----
      if (!normalizeEqual(response.original_text, target.text)) {
        throw new CopilotError(
          "INVALID_EDIT_RESPONSE",
          "模型未原样返回原文，无法安全计算修改差异，已放弃本次修改。",
        );
      }

      if (response.warnings && response.warnings.length > 0) {
        chatStore.addSystem(`提示：${response.warnings.join("；")}`);
      }

      if (normalizeEqual(response.new_text, response.original_text)) {
        chatStore.addSystem("模型认为该内容无需修改。");
        return;
      }

      // ---- 3. 客户端 Diff → PatchEngine 应用 → 登记事务（与 agent 提案共用）----
      const result = await applyGeneratedEdit(instruction, target, response, {
        regenerateCount: options.regenerateCount,
      });
      if (!result.ok) {
        chatStore.addSystem(result.message);
      }
    } catch (err) {
      const ce = toCopilotError(err);
      logger.error("编辑流程失败：", ce.code, ce.message);
      chatStore.addSystem(`修改失败：${ce.message}`, ce.code);
    } finally {
      working.value = false;
    }
  }

  /**
   * 应用半程（applyEditFlow 与 applyProposal 共用）：
   * 客户端字符级 Diff（§19）→ PatchEngine 应用（§15 乐观锁 / §25 修订 /
   * 失败自动回滚）→ 登记事务卡片（§23）。
   * 调用方保证：original_text 已与 target.text 校验一致、new_text 有差异、
   * 无 pending 冲突。working 状态由调用方管理。
   */
  async function applyGeneratedEdit(
    instruction: string,
    target: CapturedTarget,
    edit: { original_text: string; new_text: string; summary: string },
    options: { regenerateCount?: number } = {},
  ): Promise<ApplyEditResult> {
    const chatStore = useChatStore();
    const operations = diffEngine.computeEditOperations(edit.original_text, edit.new_text);
    if (operations.length === 0) {
      return { ok: false, code: "NO_CHANGE", message: "未计算出有效差异，内容无变化。" };
    }

    const plan: EditPlan = {
      id: nextTransactionId(),
      contentControlTag: editControlTag(),
      target: target.locator,
      originalText: edit.original_text,
      newText: edit.new_text,
      operations,
      summary: edit.summary,
    };
    const outcome = await patchEngine.applyEditPlan(plan);

    const tx: EditTransaction = {
      id: plan.id,
      contentControlTag: outcome.contentControlTag,
      createdAt: Date.now(),
      status: "pending",
      kind: "text",
      originalText: plan.originalText,
      newText: plan.newText,
      summary: plan.summary,
      instruction,
      changeCount: outcome.changeCount ?? outcome.appliedOps,
      locator: target.locator,
      targetKind: target.kind,
      regenerateCount: options.regenerateCount ?? 0,
    };
    registerTransaction(tx);
    chatStore.addEditMessage(tx.id);
    return { ok: true };
  }

  /**
   * Agent 提案应用（§51）：按 kind 分派，逐条独立校验（§15 乐观锁兜底）——
   * 单条失败不影响其余提案。返回结果由 agent store 汇总展示（不在此处发系统消息）。
   */
  async function applyProposal(
    instruction: string,
    proposal: AgentProposal,
    paragraphs: readonly SnapshotParagraph[],
  ): Promise<ApplyEditResult> {
    working.value = true;
    try {
      switch (proposal.kind) {
        case "text":
          return await applyTextProposal(instruction, proposal, paragraphs);
        case "format":
          return await applyFormatProposal(instruction, proposal, paragraphs);
        case "insert-table":
          return await applyTableProposal(instruction, proposal, paragraphs);
        case "insert-formula":
          return await applyFormulaProposal(instruction, proposal, paragraphs);
        case "insert-paragraph":
          return await applyParagraphProposal(instruction, proposal, paragraphs);
        case "insert-heading":
          return await applyHeadingProposal(instruction, proposal, paragraphs);
      }
      return { ok: false, code: "UNKNOWN_PROPOSAL_KIND", message: `未知提案类型：${(proposal as { kind?: string }).kind}` };
    } catch (err) {
      const ce = toCopilotError(err);
      logger.error("提案应用失败：", ce.code, ce.message);
      return { ok: false, code: ce.code, message: ce.message };
    } finally {
      working.value = false;
    }
  }

  /** 文本提案（原 applyProposal 主体）：快照段落 → 冲突检查 → 回显校验 → 编辑管线 */
  async function applyTextProposal(
    instruction: string,
    proposal: AgentTextProposalEvent,
    paragraphs: readonly SnapshotParagraph[],
  ): Promise<ApplyEditResult> {
    const target = await proposalToCapturedTarget(proposal, paragraphs);
    if (!target) {
      return {
        ok: false,
        code: "PROPOSAL_TARGET_NOT_FOUND",
        message: `段落 ${proposal.paragraph_id} 不在快照中`,
      };
    }
    // §57 冲突：同段已有 pending 事务
    const conflict = findPendingConflict(target.locator);
    if (conflict) {
      return { ok: false, code: "PENDING_EDIT_CONFLICT", message: "该内容已有待处理修改" };
    }
    // 回显校验：后端 original_text 必须与快照段落一致（防御性，§19）
    if (!normalizeEqual(proposal.original_text, target.text)) {
      return { ok: false, code: "INVALID_EDIT_RESPONSE", message: "提案原文与快照不一致" };
    }
    if (normalizeEqual(proposal.new_text, proposal.original_text)) {
      return { ok: false, code: "NO_CHANGE", message: "提案内容无变化" };
    }
    return applyGeneratedEdit(instruction, target, proposal);
  }

  /** 格式提案：快照段落 → 冲突检查 → FormatPlan → FormatEngine → 登记事务 */
  async function applyFormatProposal(
    instruction: string,
    proposal: AgentFormatProposalEvent,
    paragraphs: readonly SnapshotParagraph[],
  ): Promise<ApplyEditResult> {
    const chatStore = useChatStore();
    const target = await proposalToCapturedTarget(proposal, paragraphs);
    if (!target) {
      return {
        ok: false,
        code: "PROPOSAL_TARGET_NOT_FOUND",
        message: `段落 ${proposal.paragraph_id} 不在快照中`,
      };
    }
    const conflict = findPendingConflict(target.locator);
    if (conflict) {
      return { ok: false, code: "PENDING_EDIT_CONFLICT", message: "该内容已有待处理修改" };
    }
    const changes = formatChangesFromProposal(proposal);
    if (Object.keys(changes).length === 0) {
      return { ok: false, code: "NO_CHANGE", message: "提案未包含任何格式修改" };
    }

    const plan: FormatPlan = {
      id: nextTransactionId(),
      contentControlTag: editControlTag(),
      target: target.locator,
      changes,
      summary: proposal.summary,
    };
    const outcome = await formatEngine.applyFormatPlan(plan);

    const tx: EditTransaction = {
      id: plan.id,
      contentControlTag: outcome.contentControlTag,
      createdAt: Date.now(),
      status: "pending",
      kind: "format",
      originalText: target.text,
      newText: target.text, // 格式修改不改文字
      summary: plan.summary,
      instruction,
      changeCount: outcome.changeCount ?? 0, // Word 可能不记录格式修订 —— 0 合法
      locator: target.locator,
      targetKind: "paragraph",
      regenerateCount: 0,
      formatChanges: plan.changes,
      formatBefore: outcome.beforeValues,
    };
    registerTransaction(tx);
    chatStore.addEditMessage(tx.id);
    return { ok: true };
  }

  /** 表格提案：1.3 门 → 锚点定位（null = 文档末尾，跳过）→ 冲突检查 → InsertEngine → 登记事务 */
  async function applyTableProposal(
    instruction: string,
    proposal: AgentTableProposalEvent,
    paragraphs: readonly SnapshotParagraph[],
  ): Promise<ApplyEditResult> {
    const chatStore = useChatStore();
    if (!WordService.tableInsertSupported) {
      return {
        ok: false,
        code: "UNSUPPORTED_WORD_VERSION",
        message: "当前 Word 版本不支持插入表格（需 WordApi 1.3）",
      };
    }
    const anchorEnd = proposal.anchor_paragraph_id == null; // 文档末尾模式
    const target = anchorEnd ? null : await proposalToCapturedTarget(proposal, paragraphs);
    if (!anchorEnd && !target) {
      return {
        ok: false,
        code: "PROPOSAL_TARGET_NOT_FOUND",
        message: `锚点段落 ${proposal.anchor_paragraph_id} 不在快照中`,
      };
    }
    const conflict = findPendingConflict(target?.locator ?? null);
    if (conflict) {
      return { ok: false, code: "PENDING_EDIT_CONFLICT", message: "该锚点已有待处理修改" };
    }

    const plan: TableInsertPlan = {
      id: nextTransactionId(),
      contentControlTag: editControlTag(),
      anchor: target?.locator ?? null, // null = 文档末尾
      values: proposal.values,
      header: proposal.header,
      summary: proposal.summary,
    };
    const outcome = await insertEngine.applyTableInsertPlan(plan);

    const tx: EditTransaction = {
      id: plan.id,
      contentControlTag: outcome.contentControlTag,
      createdAt: Date.now(),
      status: "pending",
      kind: "insert-table",
      originalText: target?.text ?? "", // 锚点段文本（文档末尾为空）
      newText: "",
      summary: plan.summary,
      instruction,
      changeCount: outcome.changeCount ?? 0,
      locator: target?.locator ?? null,
      targetKind: "paragraph",
      regenerateCount: 0,
      tableValues: proposal.values,
      tableHeader: proposal.header,
      anchorId: proposal.anchor_paragraph_id ?? null,
    };
    registerTransaction(tx);
    chatStore.addEditMessage(tx.id);
    return { ok: true };
  }

  /** 公式提案：先登记审核卡，再以 Word 修订事务写入原生公式。 */
  async function applyFormulaProposal(
    instruction: string,
    proposal: AgentFormulaProposalEvent,
    paragraphs: readonly SnapshotParagraph[],
  ): Promise<ApplyEditResult> {
    const chatStore = useChatStore();
    const anchorEnd = proposal.anchor_paragraph_id == null; // 文档末尾模式
    const target = anchorEnd ? null : await proposalToCapturedTarget(proposal, paragraphs);
    if (!anchorEnd && !target) {
      return {
        ok: false,
        code: "PROPOSAL_TARGET_NOT_FOUND",
        message: `锚点段落 ${proposal.anchor_paragraph_id} 不在快照中`,
      };
    }
    const conflict = findPendingConflict(target?.locator ?? null);
    if (conflict) {
      return { ok: false, code: "PENDING_EDIT_CONFLICT", message: "该锚点已有待处理修改" };
    }

    // 纯函数预检：非法 LaTeX 在生成卡片前失败，不触碰 Word 文档。
    latexToOoxml(proposal.latex, proposal.display);

    const plan: FormulaInsertPlan = {
      id: nextTransactionId(),
      contentControlTag: editControlTag(),
      anchor: target?.locator ?? null, // null = 文档末尾
      latex: proposal.latex,
      display: proposal.display,
      summary: proposal.summary,
    };

    const tx: EditTransaction = {
      id: plan.id,
      contentControlTag: plan.contentControlTag,
      createdAt: Date.now(),
      status: "pending",
      kind: "insert-formula",
      originalText: target?.text ?? "", // 锚点段文本（文档末尾为空）
      newText: "",
      summary: plan.summary,
      instruction,
      changeCount: 0,
      locator: target?.locator ?? null,
      targetKind: "paragraph",
      regenerateCount: 0,
      latex: proposal.latex,
      displayFormula: proposal.display,
      anchorId: proposal.anchor_paragraph_id ?? null,
    };
    registerTransaction(tx);
    chatStore.addEditMessage(tx.id);
    try {
      const outcome = await applyFormulaInsertCompat(plan);
      tx.contentControlTag = outcome.contentControlTag;
      tx.changeCount = outcome.changeCount ?? 0;
      return { ok: true };
    } catch (err) {
      const ce = toCopilotError(err);
      tx.status = "invalid";
      tx.note = `写入 Word 修订失败：${ce.message}`;
      throw err;
    }
  }

  /**
   * 段落提案先登记审核卡，再以 Word 修订事务写入；普通文字和行内公式一致。
   * 即使 Word API 后续失败，卡片也会保留并明确标记失败，不再出现无卡片写入。
   */
  async function applyParagraphProposal(
    instruction: string,
    proposal: AgentParagraphProposalEvent,
    paragraphs: readonly SnapshotParagraph[],
  ): Promise<ApplyEditResult> {
    const chatStore = useChatStore();
    const lines = proposal.paragraph_text
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line.length > 0);
    if (lines.length === 0) {
      return { ok: false, code: "NO_CHANGE", message: "提案内容为空" };
    }
    const anchorEnd = proposal.anchor_paragraph_id == null; // 文档末尾模式
    const target = anchorEnd ? null : await proposalToCapturedTarget(proposal, paragraphs);
    if (!anchorEnd && !target) {
      return {
        ok: false,
        code: "PROPOSAL_TARGET_NOT_FOUND",
        message: `锚点段落 ${proposal.anchor_paragraph_id} 不在快照中`,
      };
    }
    const conflict = findPendingConflict(target?.locator ?? null);
    if (conflict) {
      return { ok: false, code: "PENDING_EDIT_CONFLICT", message: "该锚点已有待处理修改" };
    }

    const plan: ParagraphInsertPlan = {
      id: nextTransactionId(),
      contentControlTag: editControlTag(),
      anchor: target?.locator ?? null, // null = 文档末尾
      lines,
      summary: proposal.summary,
    };
    // 纯函数预检：XML 转义及行内公式转换失败时不触碰 Word。
    paragraphsToOoxml(lines);

    const tx: EditTransaction = {
      id: plan.id,
      contentControlTag: plan.contentControlTag,
      createdAt: Date.now(),
      status: "pending",
      kind: "insert-paragraph",
      originalText: target?.text ?? "", // 锚点段文本（文档末尾为空）
      newText: lines.join("\n"),
      summary: plan.summary,
      instruction,
      changeCount: 0,
      locator: target?.locator ?? null,
      targetKind: "paragraph",
      regenerateCount: 0,
      paragraphLines: lines,
      anchorId: proposal.anchor_paragraph_id ?? null,
    };
    registerTransaction(tx);
    chatStore.addEditMessage(tx.id);
    try {
      const outcome = await applyParagraphInsertCompat(plan);
      tx.contentControlTag = outcome.contentControlTag;
      tx.changeCount = outcome.changeCount ?? 0;
      return { ok: true };
    } catch (err) {
      const ce = toCopilotError(err);
      tx.status = "invalid";
      tx.note = `写入 Word 修订失败：${ce.message}`;
      throw err;
    }
  }

  /** 标题提案：使用 Word 内置 Heading1～Heading9，进入导航窗格和自动目录。 */
  async function applyHeadingProposal(
    instruction: string,
    proposal: AgentHeadingProposalEvent,
    paragraphs: readonly SnapshotParagraph[],
  ): Promise<ApplyEditResult> {
    const chatStore = useChatStore();
    const text = proposal.heading_text.trim();
    if (!text || /[\r\n\v]/.test(text)) {
      return { ok: false, code: "INVALID_EDIT_RESPONSE", message: "标题必须是非空的单行文字" };
    }
    if (!Number.isInteger(proposal.level) || proposal.level < 1 || proposal.level > 9) {
      return { ok: false, code: "INVALID_EDIT_RESPONSE", message: "标题级别必须是 1-9 的整数" };
    }
    const anchorEnd = proposal.anchor_paragraph_id == null;
    const target = anchorEnd ? null : await proposalToCapturedTarget(proposal, paragraphs);
    if (!anchorEnd && !target) {
      return {
        ok: false,
        code: "PROPOSAL_TARGET_NOT_FOUND",
        message: `锚点段落 ${proposal.anchor_paragraph_id} 不在快照中`,
      };
    }
    const conflict = findPendingConflict(target?.locator ?? null);
    if (conflict) {
      return { ok: false, code: "PENDING_EDIT_CONFLICT", message: "该锚点已有待处理修改" };
    }

    const plan: HeadingInsertPlan = {
      id: nextTransactionId(),
      contentControlTag: editControlTag(),
      anchor: target?.locator ?? null,
      text,
      level: proposal.level as HeadingInsertPlan["level"],
      summary: proposal.summary,
    };
    headingToOoxml(plan.text, plan.level); // 写入前纯函数预检

    const tx: EditTransaction = {
      id: plan.id,
      contentControlTag: plan.contentControlTag,
      createdAt: Date.now(),
      status: "pending",
      kind: "insert-heading",
      originalText: target?.text ?? "",
      newText: text,
      summary: plan.summary,
      instruction,
      changeCount: 0,
      locator: target?.locator ?? null,
      targetKind: "paragraph",
      regenerateCount: 0,
      headingText: text,
      headingLevel: plan.level,
      anchorId: proposal.anchor_paragraph_id ?? null,
    };
    registerTransaction(tx);
    chatStore.addEditMessage(tx.id);
    try {
      const outcome = await applyHeadingInsertCompat(plan);
      tx.contentControlTag = outcome.contentControlTag;
      tx.changeCount = outcome.changeCount ?? 0;
      return { ok: true };
    } catch (err) {
      const ce = toCopilotError(err);
      tx.status = "invalid";
      tx.note = `写入 Word 修订失败：${ce.message}`;
      throw err;
    }
  }

  // ---------------- 事务操作（§78 / §79 / §80） ----------------

  /** 接受事务（仅作用于该事务的 Word 修订 / Content Control，§55） */
  async function accept(id: string): Promise<void> {
    const tx = transactions.value[id];
    if (!tx || tx.status !== "pending" || working.value) return;
    const chatStore = useChatStore();
    working.value = true;
    try {
      const outcome = await revisionService.acceptTransaction(transactionRef(tx));
      applyOutcome(tx, outcome);
      if (outcome === "not_found") {
        chatStore.addSystem(`${tx.id} 的修订内容未找到，可能已被手动清理。`);
      }
    } catch (err) {
      const ce = toCopilotError(err);
      chatStore.addSystem(`接受修改失败：${ce.message}`, ce.code);
    } finally {
      working.value = false;
    }
  }

  /** 拒绝事务（恢复原文 / 撤销插入） */
  async function reject(id: string): Promise<void> {
    const tx = transactions.value[id];
    if (!tx || tx.status !== "pending" || working.value) return;
    const chatStore = useChatStore();
    working.value = true;
    try {
      const outcome = await revisionService.rejectTransaction(transactionRef(tx));
      applyOutcome(tx, outcome);
      if (outcome === "not_found") {
        chatStore.addSystem(`${tx.id} 的修订内容未找到，可能已被手动清理。`);
      }
    } catch (err) {
      const ce = toCopilotError(err);
      chatStore.addSystem(`拒绝修改失败：${ce.message}`, ce.code);
    } finally {
      working.value = false;
    }
  }

  /**
   * 重新生成（§80）：拒绝当前修订（恢复原文）→
   * 携带 avoid_texts 重新请求 → 生成新的替代版本。
   * 复用已保存的 locator，光标移动不影响。
   * 仅文本事务支持（格式 / 插入类无「重新生成」语义，UI 亦隐藏按钮）。
   */
  async function regenerate(id: string, feedback = ""): Promise<void> {
    const tx = transactions.value[id];
    const chatStore = useChatStore();
    if (!tx) return;
    if (tx.kind !== "text") {
      chatStore.addSystem("该类型修改不支持重新生成。");
      return;
    }
    if (tx.status !== "pending") {
      chatStore.addSystem("该修改已处理完毕，无法重新生成。");
      return;
    }
    if (working.value) {
      chatStore.addSystem("正在处理修改，请稍候。");
      return;
    }

    const instruction = tx.instruction;
    const locator = tx.locator;
    if (!locator) return; // text 类事务恒有 locator（防御性）
    const olderAvoidTexts = collectAvoidTexts(locator);

    working.value = true;
    try {
      // 先回滚当前修订（恢复原文），旧事务标记为已拒绝
      const outcome = await revisionService.rejectTransaction(transactionRef(tx));
      applyOutcome(tx, outcome, "已重新生成，回滚此前的修改。");
    } catch (err) {
      working.value = false;
      const ce = toCopilotError(err);
      chatStore.addSystem(`回滚旧修改失败：${ce.message}`, ce.code);
      return;
    }

    // §56：复用 locator，不依赖当前光标
    const target: CapturedTarget = {
      kind: tx.targetKind,
      text: tx.originalText,
      textHash: locator.textHash,
      prefix: locator.prefix ?? "",
      suffix: locator.suffix ?? "",
      locator,
    };
    await applyEditFlow(instruction, target, {
      regenerateCount: tx.regenerateCount + 1,
      avoidTexts: [tx.newText, ...olderAvoidTexts],
      regenerationFeedback: feedback,
    });
  }

  /**
   * 对账（§29）：用户可能用 Word 原生审阅处理了修订。
   * 刷新 pending 事务状态；已处理的事务不报错，只更新状态。
   * （未跟踪的格式 / 插入事务走 existence 模式 —— CC 存在即 pending。）
   */
  async function reconcile(): Promise<void> {
    const pending = pendingList.value;
    if (pending.length === 0) return;
    try {
      const outcomes = await revisionService.reconcile(pending.map(transactionRef));
      for (const tx of pending) {
        const outcome = outcomes.get(tx.contentControlTag);
        if (outcome) applyOutcome(tx, outcome);
      }
    } catch (err) {
      logger.debug("对账失败（忽略，下次再试）：", toCopilotError(err).message);
    }
  }

  function clear(): void {
    transactions.value = {};
    order.value = [];
  }

  // ---------------- 内部 ----------------

  /** 事务 → RevisionService 引用（按 kind 分派接受 / 拒绝 / 对账语义） */
  function transactionRef(tx: EditTransaction): TransactionRef {
    return {
      tag: tx.contentControlTag,
      kind: tx.kind,
      appliedChangeCount: tx.changeCount,
      formatBefore: tx.formatBefore,
    };
  }

  /** 格式提案（snake_case / null 宽松）→ FormatChanges（camelCase / 仅含修改项） */
  function formatChangesFromProposal(p: AgentFormatProposalEvent): FormatChanges {
    const changes: FormatChanges = {};
    if (p.bold !== null && p.bold !== undefined) changes.bold = p.bold;
    if (p.italic !== null && p.italic !== undefined) changes.italic = p.italic;
    if (p.underline !== null && p.underline !== undefined) changes.underline = p.underline;
    if (p.strikethrough !== null && p.strikethrough !== undefined) changes.strikethrough = p.strikethrough;
    if (p.font_name !== null && p.font_name !== undefined) changes.fontName = p.font_name;
    if (p.font_size !== null && p.font_size !== undefined) changes.fontSize = p.font_size;
    if (p.color !== null && p.color !== undefined) changes.color = p.color;
    if (p.alignment !== null && p.alignment !== undefined) changes.alignment = p.alignment;
    if (p.paragraph_style !== null && p.paragraph_style !== undefined) changes.paragraphStyle = p.paragraph_style;
    return changes;
  }

  function applyOutcome(tx: EditTransaction, outcome: TransactionOutcome, note?: string): void {
    switch (outcome) {
      case "accepted":
        tx.status = "accepted";
        break;
      case "rejected":
        tx.status = "rejected";
        if (note) tx.note = note;
        break;
      case "handled_externally":
        tx.status = "invalid";
        tx.note = "已在 Word 审阅中处理";
        break;
      case "not_found":
        tx.status = "invalid";
        tx.note = "修订内容未找到（可能已被手动清理）";
        break;
      case "pending":
        break;
    }
  }

  /** 收集同一目标（同原文哈希）历史上生成过的文本，供重新生成避开 */
  function collectAvoidTexts(locator: NonNullable<EditTransaction["locator"]>): string[] {
    return order.value
      .map((id) => transactions.value[id])
      .filter(
        (tx): tx is EditTransaction =>
          !!tx && !!tx.locator && tx.locator.textHash === locator.textHash && tx.status !== "pending",
      )
      .map((tx) => tx.newText)
      .filter((text) => text.trim().length > 0);
  }

  return {
    transactions,
    order,
    working,
    pendingList,
    getTransaction,
    sendEdit,
    applyProposal,
    accept,
    reject,
    regenerate,
    reconcile,
    clear,
  };
});

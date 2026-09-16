<script setup lang="ts">
import { computed } from "vue";
import katex from "katex";
import { useEditStore } from "@/stores/edits";
import { useChatStore } from "@/stores/chat";
import { TRANSACTION_STATUS_LABELS } from "@/models/EditTransaction";
import type { EditTransaction, TransactionKind } from "@/models/EditTransaction";
import type { FormatChanges } from "@/models/EditPlan";

const props = defineProps<{ transactionId: string }>();

const editsStore = useEditStore();
const chatStore = useChatStore();

const tx = computed<EditTransaction | undefined>(() => editsStore.getTransaction(props.transactionId));

const busy = computed(() => editsStore.working);

const KIND_LABELS: Record<TransactionKind, string> = {
  text: "AI 修改",
  format: "AI 格式",
  "insert-table": "AI 表格",
  "insert-formula": "AI 公式",
  "insert-paragraph": "AI 段落",
};

const kindLabel = computed(() => (tx.value ? KIND_LABELS[tx.value.kind] : ""));

/** 插入类：有锚点 = 锚点段之后；null = 文档末尾（空文档场景） */
const targetLabel = computed(() => {
  if (!tx.value) return "";
  if (tx.value.kind === "text") return tx.value.targetKind === "selection" ? "选区" : "段落";
  if (tx.value.kind === "format") return "段落";
  return tx.value.anchorId == null ? "文档末尾" : "锚点段落之后";
});

const timeLabel = computed(() => {
  if (!tx.value) return "";
  const d = new Date(tx.value.createdAt);
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
});

// ---------------- 格式卡片：属性变更列表 ----------------

interface FormatRow {
  key: string;
  label: string;
  before: string;
  after: string;
  hasBefore: boolean;
}

const FORMAT_FIELD_LABELS: Record<keyof FormatChanges, string> = {
  bold: "加粗",
  italic: "斜体",
  underline: "下划线",
  strikethrough: "删除线",
  fontName: "字体",
  fontSize: "字号",
  color: "颜色",
  alignment: "对齐",
};

const onOff = (v: boolean): string => (v ? "开" : "关");

const ALIGNMENT_LABELS: Record<string, string> = {
  left: "左对齐",
  center: "居中",
  right: "右对齐",
  justify: "两端对齐",
  Left: "左对齐",
  Centered: "居中",
  Right: "右对齐",
  Justified: "两端对齐",
};

const formatRows = computed<FormatRow[]>(() => {
  const changes = tx.value?.formatChanges;
  const before = tx.value?.formatBefore;
  if (!changes) return [];
  const rows: FormatRow[] = [];
  for (const key of Object.keys(changes) as (keyof FormatChanges)[]) {
    const value = changes[key];
    if (value === undefined) continue;
    switch (key) {
      case "bold":
      case "italic":
      case "strikethrough": {
        // Object.keys 收窄不到具体类型 —— 此处值恒为 boolean
        const b = before?.[key] as boolean | undefined;
        rows.push({
          key,
          label: FORMAT_FIELD_LABELS[key],
          before: b === undefined ? "" : onOff(b),
          after: onOff(value as boolean),
          hasBefore: b !== undefined,
        });
        break;
      }
      case "underline": {
        // before 存 Word 枚举字符串（"None"/"Single"/…）
        const b = before?.underline;
        const bLabel = b === undefined ? "" : b === "None" || b === "" ? "无" : "有";
        rows.push({
          key,
          label: FORMAT_FIELD_LABELS[key],
          before: bLabel,
          after: onOff(value as boolean),
          hasBefore: b !== undefined,
        });
        break;
      }
      case "fontName":
      case "color":
        rows.push({
          key,
          label: FORMAT_FIELD_LABELS[key],
          before: before?.[key] ?? "",
          after: String(value),
          hasBefore: before?.[key] !== undefined,
        });
        break;
      case "fontSize": {
        const b = before?.fontSize;
        rows.push({
          key,
          label: FORMAT_FIELD_LABELS[key],
          before: b === undefined ? "" : `${b} 磅`,
          after: `${value as number} 磅`,
          hasBefore: b !== undefined,
        });
        break;
      }
      case "alignment":
        rows.push({
          key,
          label: FORMAT_FIELD_LABELS[key],
          before: before?.alignment ? (ALIGNMENT_LABELS[before.alignment] ?? before.alignment) : "",
          after: ALIGNMENT_LABELS[value as string] ?? String(value),
          hasBefore: before?.alignment !== undefined,
        });
        break;
    }
  }
  return rows;
});

// ---------------- 表格卡片：迷你预览（≤5 数据行） ----------------

const MAX_PREVIEW_ROWS = 5;

const tablePreview = computed(() => {
  const values = tx.value?.tableValues;
  if (!tx.value || tx.value.kind !== "insert-table" || !values || values.length === 0) return null;
  const header = tx.value.tableHeader === true;
  const dataRows = header ? values.slice(1) : values;
  return {
    header,
    headerRow: header ? values[0] : [],
    rows: dataRows.slice(0, MAX_PREVIEW_ROWS),
    totalDataRows: dataRows.length,
    truncated: dataRows.length > MAX_PREVIEW_ROWS,
  };
});

// ---------------- 公式卡片：KaTeX 预览 ----------------

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** 与 markdown.ts renderMath 同信任级：失败兜底转义字面量（插入前的 latex 已通过严格渲染） */
const formulaHtml = computed(() => {
  const latex = tx.value?.latex;
  if (!tx.value || tx.value.kind !== "insert-formula" || latex === undefined) return "";
  try {
    return katex.renderToString(latex, {
      displayMode: tx.value.displayFormula === true,
      throwOnError: false,
      strict: "ignore",
    });
  } catch {
    return `<code>${escapeHtml(latex)}</code>`;
  }
});

// ---------------- 段落卡片：文本预览（≤3 段，每段截断） ----------------

const MAX_PREVIEW_PARAGRAPH_LINES = 3;
const MAX_PREVIEW_LINE_LENGTH = 80;

const paragraphPreview = computed(() => {
  const lines = tx.value?.paragraphLines;
  if (!tx.value || tx.value.kind !== "insert-paragraph" || !lines || lines.length === 0) return null;
  const shown = lines.slice(0, MAX_PREVIEW_PARAGRAPH_LINES).map((line) =>
    line.length > MAX_PREVIEW_LINE_LENGTH ? `${line.slice(0, MAX_PREVIEW_LINE_LENGTH)}…` : line,
  );
  return {
    lines: shown,
    total: lines.length,
    truncated: lines.length > MAX_PREVIEW_PARAGRAPH_LINES,
  };
});
</script>

<template>
  <div v-if="tx" class="edit-card" :class="tx.status">
    <div class="card-head">
      <span class="card-tag">{{ kindLabel }}</span>
      <span class="card-id">{{ tx.id }}</span>
      <span class="card-status" :class="tx.status">{{ TRANSACTION_STATUS_LABELS[tx.status] }}</span>
      <span class="card-time">{{ timeLabel }}</span>
    </div>

    <div class="card-summary">{{ tx.summary || (tx.kind === "text" ? "已完成修改" : "已完成操作") }}</div>

    <!-- 格式卡片：属性变更列表 -->
    <div v-if="tx.kind === 'format' && formatRows.length > 0" class="format-list">
      <div v-for="row in formatRows" :key="row.key" class="format-row">
        <span class="format-label">{{ row.label }}</span>
        <span class="format-values">
          <template v-if="row.hasBefore">{{ row.before }} → {{ row.after }}</template>
          <template v-else>{{ row.after }}</template>
        </span>
      </div>
    </div>

    <!-- 表格卡片：迷你预览 -->
    <div v-if="tablePreview" class="table-preview">
      <table>
        <thead v-if="tablePreview.header && tablePreview.headerRow.length > 0">
          <tr>
            <th v-for="(cell, i) in tablePreview.headerRow" :key="i">{{ cell }}</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(row, r) in tablePreview.rows" :key="r">
            <td v-for="(cell, i) in row" :key="i">{{ cell }}</td>
          </tr>
        </tbody>
      </table>
      <div v-if="tablePreview.truncated" class="table-more">
        …（共 {{ tablePreview.totalDataRows }} 行，插入后查看完整表格）
      </div>
    </div>

    <!-- 公式卡片：KaTeX 预览 + LaTeX 原文 -->
    <div v-if="tx.kind === 'insert-formula' && formulaHtml" class="formula-preview">
      <!-- eslint-disable-next-line vue/no-v-html — katex 输出与 markdown.ts 同信任级 -->
      <div class="formula-math" :class="{ display: tx.displayFormula }" v-html="formulaHtml"></div>
      <div class="formula-latex">{{ tx.latex }}</div>
    </div>

    <!-- 段落卡片：文本预览（插值自动转义） -->
    <div v-if="paragraphPreview" class="paragraph-preview">
      <p v-for="(line, i) in paragraphPreview.lines" :key="i">{{ line }}</p>
      <div v-if="paragraphPreview.truncated" class="paragraph-more">
        …（共 {{ paragraphPreview.total }} 段，插入后查看完整内容）
      </div>
    </div>

    <div class="card-meta">
      <template v-if="tx.kind === 'text'">{{ tx.changeCount }} 处修订 · {{ targetLabel }}</template>
      <template v-else-if="tx.kind === 'format'">
        格式修改 · {{ targetLabel }}
        <template v-if="tx.changeCount > 0"> · {{ tx.changeCount }} 处修订</template>
      </template>
      <template v-else>插入内容 · {{ targetLabel }}<template v-if="tx.changeCount > 0"> · {{ tx.changeCount }} 处修订</template></template>
      <template v-if="tx.kind === 'text' && tx.regenerateCount > 0"> · 第 {{ tx.regenerateCount + 1 }} 版</template>
    </div>

    <p v-if="tx.note" class="card-note">{{ tx.note }}</p>

    <div v-if="tx.status === 'pending'" class="card-actions">
      <button class="btn primary" :disabled="busy" @click="editsStore.accept(tx.id)">接受</button>
      <button class="btn danger" :disabled="busy" @click="editsStore.reject(tx.id)">拒绝</button>
      <button v-if="tx.kind === 'text'" class="btn" :disabled="busy" @click="editsStore.regenerate(tx.id)">重新生成</button>
    </div>
    <div v-else class="card-done-hint">可在 Word「审阅」中查看修订历史</div>
  </div>
</template>

<style scoped>
.edit-card {
  margin: 10px 0;
  padding: 8px 10px;
  border: 1px solid var(--border-strong);
  border-left: 3px solid var(--accent);
  border-radius: var(--radius);
  background: var(--bg);
}

.edit-card.accepted {
  border-left-color: var(--success);
}

.edit-card.rejected,
.edit-card.invalid {
  border-left-color: var(--text-muted);
  opacity: 0.75;
}

.card-head {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 11px;
  color: var(--text-muted);
}

.card-tag {
  font-weight: 600;
  color: var(--accent);
}

.card-id {
  font-family: Consolas, "Courier New", monospace;
}

.card-status {
  margin-left: auto;
  font-weight: 600;
}

.card-status.pending {
  color: var(--accent);
}

.card-status.accepted {
  color: var(--success);
}

.card-status.rejected {
  color: var(--text-muted);
}

.card-status.invalid {
  color: var(--text-muted);
}

.card-time {
  font-variant-numeric: tabular-nums;
}

.card-summary {
  margin-top: 4px;
  font-size: 13px;
  color: var(--text);
  word-break: break-word;
}

.card-meta {
  margin-top: 2px;
  font-size: 12px;
  color: var(--text-secondary);
}

.card-note {
  margin: 4px 0 0;
  font-size: 12px;
  color: var(--text-muted);
}

.card-actions {
  display: flex;
  gap: 6px;
  margin-top: 8px;
}

.card-done-hint {
  margin-top: 6px;
  font-size: 11.5px;
  color: var(--text-muted);
}

/* ---- 格式卡片 ---- */

.format-list {
  margin-top: 6px;
  display: flex;
  flex-direction: column;
  gap: 2px;
  padding: 6px 8px;
  border-radius: var(--radius);
  background: var(--bg-secondary, rgba(127, 127, 127, 0.08));
}

.format-row {
  display: flex;
  gap: 8px;
  font-size: 12px;
  line-height: 1.6;
}

.format-label {
  min-width: 48px;
  color: var(--text-muted);
}

.format-values {
  color: var(--text);
  word-break: break-word;
}

/* ---- 表格卡片 ---- */

.table-preview {
  margin-top: 6px;
  overflow-x: auto;
}

.table-preview table {
  border-collapse: collapse;
  font-size: 12px;
  max-width: 100%;
}

.table-preview th,
.table-preview td {
  border: 1px solid var(--border-strong);
  padding: 3px 8px;
  text-align: left;
  white-space: nowrap;
  max-width: 200px;
  overflow: hidden;
  text-overflow: ellipsis;
}

.table-preview th {
  background: var(--bg-secondary, rgba(127, 127, 127, 0.12));
  font-weight: 600;
}

.table-more {
  margin-top: 2px;
  font-size: 11.5px;
  color: var(--text-muted);
}

/* ---- 公式卡片 ---- */

.formula-preview {
  margin-top: 6px;
  padding: 6px 8px;
  border-radius: var(--radius);
  background: var(--bg-secondary, rgba(127, 127, 127, 0.08));
}

.formula-math {
  font-size: 14px;
  overflow-x: auto;
}

.formula-math :deep(.katex-display) {
  margin: 4px 0;
}

.formula-latex {
  margin-top: 4px;
  font-family: Consolas, "Courier New", monospace;
  font-size: 11px;
  color: var(--text-muted);
  word-break: break-all;
}

/* ---- 段落卡片 ---- */

.paragraph-preview {
  margin-top: 6px;
  padding: 6px 8px;
  border-radius: var(--radius);
  background: var(--bg-secondary, rgba(127, 127, 127, 0.08));
}

.paragraph-preview p {
  margin: 2px 0;
  font-size: 12.5px;
  line-height: 1.6;
  color: var(--text);
  word-break: break-word;
}

.paragraph-more {
  margin-top: 2px;
  font-size: 11.5px;
  color: var(--text-muted);
}
</style>

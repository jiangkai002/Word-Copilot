/**
 * 编辑计划模型（对应需求文档 §14 RangeLocator / §20 DiffOperation / §22 EditPlan）。
 */

/** 三级定位信息：paragraphId → 原文精确匹配 → prefix + 原文 + suffix */
export interface RangeLocator {
  /** 第一优先级：段落 uniqueLocalId（同会话内定位） */
  paragraphId?: string;
  originalText: string;
  textHash: string;
  /** 同段落内目标之前的文字（用于消歧义 / 三级定位） */
  prefix?: string;
  /** 同段落内目标之后的文字 */
  suffix?: string;
}

/** DiffEngine 输出的字符级差异操作（§20） */
export type DiffOperation =
  | { type: "equal"; text: string }
  | { type: "delete"; text: string }
  | { type: "insert"; text: string };

export type EditOperationType = "insert" | "delete" | "replace";

export interface EditOperation {
  type: EditOperationType;
  /** 规范化后原文（\n 为段落分隔）中的起始偏移 */
  start: number;
  /** 规范化后原文中的结束偏移（exclusive） */
  end: number;
  /** delete/replace：要删除/替换的原文 */
  oldText?: string;
  /** insert/replace：要插入的新文本 */
  newText?: string;
}

export interface EditPlan {
  id: string;
  /** Content Control tag: word_ai_edit:{uuid} */
  contentControlTag: string;
  target: RangeLocator;
  originalText: string;
  newText: string;
  operations: EditOperation[];
  summary?: string;
}

// ---- 格式 / 插入计划（Agent 三项格式能力）----

/** 格式修改请求（camelCase；只含要修改的字段，undefined = 不改） */
export interface FormatChanges {
  bold?: boolean;
  italic?: boolean;
  /** true = 单下划线；false = 无下划线 */
  underline?: boolean;
  strikethrough?: boolean;
  fontName?: string;
  fontSize?: number;
  /** #RRGGBB */
  color?: string;
  alignment?: "left" | "center" | "right" | "justify";
}

/**
 * 应用前原值快照（拒绝时回写用）。混合格式读出 null → 存 undefined（回滚跳过）。
 * underline / alignment 存 Word 枚举字符串，可原样写回（不做有损布尔往返）。
 */
export interface FormatBeforeValues {
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  underline?: string;
  fontName?: string;
  fontSize?: number;
  color?: string;
  alignment?: string;
}

/** 格式修改计划（整段；目标段定位与乐观锁校验同 EditPlan） */
export interface FormatPlan {
  id: string;
  contentControlTag: string;
  /** 目标段定位（locator.textHash 校验段文本未变） */
  target: RangeLocator;
  /** 仅含要修改的字段 */
  changes: FormatChanges;
  summary?: string;
}

/** 表格插入计划：插入到锚点段落之后；anchor = null 时插入到文档末尾 */
export interface TableInsertPlan {
  id: string;
  contentControlTag: string;
  /**
   * 锚点段定位（hash 校验锚点未变 → 插入位置仍正确）；
   * null = 文档末尾（Body.insertTable("End")，空文档可用）
   */
  anchor: RangeLocator | null;
  values: string[][];
  header: boolean;
  summary?: string;
}

/** 公式插入计划：LaTeX → OMML，插入到锚点段落之后 / 文档末尾 */
export interface FormulaInsertPlan {
  id: string;
  contentControlTag: string;
  /** null = 文档末尾（Body.insertOoxml("End")） */
  anchor: RangeLocator | null;
  latex: string;
  /** true = 独立成行居中（w:jc center）；false = 行内（不居中） */
  display: boolean;
  summary?: string;
}

/**
 * 段落插入计划：paragraph_text 按换行拆为多个段落，可用 $LaTeX$ 标注行内公式，
 * 插入到锚点段落之后 / 文档末尾（Body.insertOoxml("End") 一次写入多个 w:p）。
 */
export interface ParagraphInsertPlan {
  id: string;
  contentControlTag: string;
  /** null = 文档末尾 */
  anchor: RangeLocator | null;
  /** 每行一个段落的文本（后端已限 1-20 行、共 ≤2000 字符） */
  lines: string[];
  summary?: string;
}

/** Word 内置标题段落插入计划：Heading1～Heading9，可进入导航窗格 / 自动目录。 */
export interface HeadingInsertPlan {
  id: string;
  contentControlTag: string;
  /** null = 文档末尾 */
  anchor: RangeLocator | null;
  text: string;
  level: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  summary?: string;
}

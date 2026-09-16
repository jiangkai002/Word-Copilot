/**
 * 快捷命令（对应需求文档 §33 / §51）。
 * 解释 / 总结 → Chat；润色 / 专业化 / 缩写 / 扩写 / 纠错 / 翻译并替换 → Edit；
 * 全文纠错 → Agent（批量编辑）。
 */
export interface QuickCommand {
  id: string;
  label: string;
  kind: "chat" | "edit" | "agent";
  instruction: string;
}

export const QUICK_COMMANDS: QuickCommand[] = [
  {
    id: "fix-all",
    label: "全文纠错",
    kind: "agent",
    instruction: "检查全文，找出所有错别字、语病和标点问题，逐段提交修改提案。",
  },
  {
    id: "polish",
    label: "润色",
    kind: "edit",
    instruction: "润色这段文字，使表达更流畅、更书面化，保持原意不变。",
  },
  {
    id: "professional",
    label: "专业化",
    kind: "edit",
    instruction: "把这段文字改写得更专业、更正式，使用规范的技术表达，保持原意。",
  },
  {
    id: "shorten",
    label: "缩写",
    kind: "edit",
    instruction: "压缩这段文字，保留核心信息，明显减少篇幅。",
  },
  {
    id: "expand",
    label: "扩写",
    kind: "edit",
    instruction: "扩写这段文字，补充合理的细节与论证，使内容更充实，不编造事实。",
  },
  {
    id: "fix",
    label: "纠错",
    kind: "edit",
    instruction: "找出并修正这段文字中的错别字、语病和标点错误，不改变原意与结构。",
  },
  {
    id: "translate",
    label: "翻译并替换",
    kind: "edit",
    instruction: "把这段文字翻译成英文，保持原意与语气。",
  },
  {
    id: "summarize",
    label: "总结",
    kind: "chat",
    instruction: "请总结当前内容的主要信息。",
  },
  {
    id: "explain",
    label: "解释",
    kind: "chat",
    instruction: "请解释当前内容的含义，说明它在文中的作用。",
  },
];

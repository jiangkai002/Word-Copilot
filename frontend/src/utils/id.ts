/** 生成 UUID（crypto.randomUUID，安全上下文可用；带兜底实现） */
export function uuid(): string {
  const c = globalThis.crypto;
  if (c && typeof c.randomUUID === "function") {
    return c.randomUUID();
  }
  // 兜底：符合 8-4-4-4-12 格式（非加密强随机，仅用于标识）
  const hex = (n: number) =>
    Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16)).join("");
  return `${hex(8)}-${hex(4)}-${hex(4)}-${hex(4)}-${hex(12)}`;
}

let txCounter = 0;

/** 事务展示编号：AI_EDIT_001、AI_EDIT_002 …… */
export function nextTransactionId(): string {
  txCounter += 1;
  return `AI_EDIT_${String(txCounter).padStart(3, "0")}`;
}

/** Content Control tag（§24）：word_ai_edit:{UUID} */
export function editControlTag(): string {
  return `word_ai_edit:${uuid()}`;
}

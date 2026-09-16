/**
 * SHA-256（Web Crypto）。任务窗格运行在 https 安全上下文，crypto.subtle 可用；
 * Node 22（测试环境）同样内置 globalThis.crypto。
 */
export async function sha256Hex(input: string): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    throw new Error("当前环境不支持 Web Crypto（crypto.subtle），无法计算 textHash。");
  }
  const data = new TextEncoder().encode(input);
  const digest = await subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

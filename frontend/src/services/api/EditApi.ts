/**
 * EditApi（对应需求文档 §18）：POST /api/v1/edit
 * 后端只返回 original_text / new_text，字符级 Diff 在客户端完成（§19）。
 */
import type { EditRequestPayload, EditResponsePayload } from "@/models/Api";
import { CopilotError } from "@/utils/errors";
import { resolveUrl, throwApiError } from "./ApiClient";

export class EditApi {
  async requestEdit(
    payload: EditRequestPayload,
    signal?: AbortSignal,
  ): Promise<EditResponsePayload> {
    let response: Response;
    try {
      response = await fetch(resolveUrl("/api/v1/edit"), {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify(payload),
        signal,
      });
    } catch (err) {
      if (err instanceof DOMException && err.name === "AbortError") throw err;
      throw new CopilotError("NETWORK_ERROR", "无法连接后端（/api/v1/edit）");
    }
    if (!response.ok) {
      await throwApiError(response);
    }
    const body = (await response.json()) as EditResponsePayload;
    if (!body || typeof body.new_text !== "string" || typeof body.original_text !== "string") {
      throw new CopilotError("INVALID_EDIT_RESPONSE", "后端返回结构不完整");
    }
    return body;
  }
}

export const editApi = new EditApi();

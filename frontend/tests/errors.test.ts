import { describe, expect, it } from "vitest";

import { toCopilotError } from "@/utils/errors";

describe("toCopilotError", () => {
  it("把 service 方法缺失识别为前端版本错配，而不是网络异常", () => {
    const error = toCopilotError(
      new TypeError("insertEngine.applyFormulaInsertPlan is not a function"),
    );

    expect(error.code).toBe("FRONTEND_VERSION_MISMATCH");
    expect(error.message).toContain("请重新加载任务窗格");
  });

  it("仍把 fetch 连接失败识别为网络异常", () => {
    const error = toCopilotError(new TypeError("Failed to fetch"));

    expect(error.code).toBe("NETWORK_ERROR");
  });

  it("不再把普通 TypeError 误报为后端未启动", () => {
    const error = toCopilotError(new TypeError("unexpected client failure"));

    expect(error.code).toBe("UNKNOWN");
  });
});

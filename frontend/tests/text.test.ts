/**
 * 文本规范化 / 哈希 单元测试（对应需求文档 §49）。
 */
import { describe, expect, it } from "vitest";
import {
  normalizeText,
  normalizeEqual,
  reviewedTextEquals,
  stripTrailingMarks,
  stripLeadingMarks,
  toWordText,
  textHash,
} from "@/utils/text";

describe("normalizeText（Word 行尾统一为 \n）", () => {
  it("CRLF → LF", () => {
    expect(normalizeText("第一段\r\n第二段")).toBe("第一段\n第二段");
  });

  it("CR（段落标记）→ LF", () => {
    expect(normalizeText("第一段\r第二段")).toBe("第一段\n第二段");
  });

  it("VT（软换行）→ LF", () => {
    expect(normalizeText("行内\v换行")).toBe("行内\n换行");
  });

  it("三种混合", () => {
    expect(normalizeText("a\r\nb\rc\vd")).toBe("a\nb\nc\nd");
  });

  it("不做 trim（保留前后空格 —— 否则定位出错）", () => {
    expect(normalizeText(" 空格保留 ")).toBe(" 空格保留 ");
  });

  it("无行尾原样返回", () => {
    expect(normalizeText("普通文本")).toBe("普通文本");
  });
});

describe("toWordText（\n → \r，插入 Word 时使用）", () => {
  it("段落分隔转换", () => {
    expect(toWordText("第一段\n第二段")).toBe("第一段\r第二段");
  });

  it("与 normalizeText 互逆（除 CRLF/VT 之外无信息损失）", () => {
    const original = "第一段\n第二段\n第三段";
    expect(normalizeText(toWordText(original))).toBe(original);
  });
});

describe("stripTrailingMarks / stripLeadingMarks", () => {
  it("去掉结尾段落标记", () => {
    expect(stripTrailingMarks("一段文字\r\r")).toBe("一段文字");
  });

  it("去掉结尾软换行 / 换行", () => {
    expect(stripTrailingMarks("文字\v")).toBe("文字");
    expect(stripTrailingMarks("文字\n")).toBe("文字");
  });

  it("去掉 Word 表格单元格结束标记", () => {
    expect(stripTrailingMarks("单元格\r\x07")).toBe("单元格");
  });

  it("正文中间的不受影响", () => {
    expect(stripTrailingMarks("中\r间\r")).toBe("中\r间");
  });

  it("去掉开头标记", () => {
    expect(stripLeadingMarks("\r\n开头")).toBe("开头");
  });
});

describe("textHash（SHA-256）", () => {
  it("已知向量：SHA-256(\"abc\")", async () => {
    expect(await textHash("abc")).toBe(
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
    );
  });

  it("已知向量：SHA-256(\"\")", async () => {
    expect(await textHash("")).toBe(
      "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
    );
  });

  it("不同行尾形态哈希一致（乐观锁兼容 Word 视图差异）", async () => {
    const a = await textHash("第一段\r\n第二段");
    const b = await textHash("第一段\r第二段");
    const c = await textHash("第一段\n第二段");
    expect(a).toBe(b);
    expect(b).toBe(c);
  });

  it("内容变化哈希变化", async () => {
    expect(await textHash("旧文本")).not.toBe(await textHash("新文本"));
  });
});

describe("normalizeEqual", () => {
  it("行尾差异视为相等", () => {
    expect(normalizeEqual("A\r\nB", "A\nB")).toBe(true);
  });

  it("内容差异不相等", () => {
    expect(normalizeEqual("A", "B")).toBe(false);
  });

  it("首尾空格参与比较", () => {
    expect(normalizeEqual(" A", "A")).toBe(false);
  });
});

describe("reviewedTextEquals（接受全部修订后文本比对）", () => {
  it("完全一致", () => {
    expect(reviewedTextEquals("相同文本", "相同文本")).toBe(true);
  });

  it("真实回归：getReviewedText 输出含结构标记前缀“<<”（2026-09-17 Word 桌面版实测）", () => {
    const expected = "步骤1.2 统一BIM坐标系";
    expect(reviewedTextEquals("<<" + expected, expected)).toBe(true);
  });

  it("标记字符出现在中间/结尾同样容忍（剥离两侧标记后比较）", () => {
    expect(reviewedTextEquals("<a>", "a")).toBe(true);
    expect(reviewedTextEquals("a\x07", "a")).toBe(true);
  });

  it("期望文本本身含 '<' 或 '>' 时不容忍（防止掩盖真实差异）", () => {
    expect(reviewedTextEquals("a < b", "a < b")).toBe(true); // 完全相等仍通过
    expect(reviewedTextEquals("a > b", "a < b")).toBe(false); // 含标记字符 → 严格比较
  });

  it("真实内容差异不相等（不能因剥离标记而误判成功）", () => {
    expect(reviewedTextEquals("<<完全不同的文本", "步骤1.2 统一BIM坐标系")).toBe(false);
  });
});

/**
 * DocumentService.scanParagraphs 表内段落探测（inTable）测试。
 *
 * Body.paragraphs 自 WordApi 1.3 起包含表格单元格内的段落 —— 快照必须
 * 标注 inTable，供后端拒绝「插入到表内段之后」的锚点（内容会写进单元格）。
 * 探测走 parentTableCellOrNullObject null-object 模式（不在表内时
 * isNullObject === true）；宿主不支持 1.3 或探测抛错时降级为 false。
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { documentService } from "@/services/word/DocumentService";

interface MockParagraph {
  uniqueLocalId?: string;
  text: string;
  styleBuiltIn: string;
  style: string;
  inTable: boolean;
}

/** 打桩 Office.js：Word.run 执行回调，段落集合在首次 sync 后填充 */
function stubWord(
  paragraphs: MockParagraph[],
  options: { supported13?: boolean; probeThrows?: boolean } = {},
): void {
  const { supported13 = true, probeThrows = false } = options;
  const items = paragraphs.map((p) => ({
    uniqueLocalId: p.uniqueLocalId,
    text: p.text,
    styleBuiltIn: p.styleBuiltIn,
    style: p.style,
    parentTableCellOrNullObject: {
      isNullObject: !p.inTable,
      load: () => {
        if (probeThrows) throw new Error("InvalidArgument: parentTableCellOrNullObject");
      },
    },
  }));
  let synced = false;
  const paras = { items: [] as typeof items, load: () => {} };
  const ctx = {
    document: { body: { paragraphs: paras } },
    sync: async () => {
      if (!synced) {
        paras.items = items;
        synced = true;
      }
    },
  };
  const globals = globalThis as unknown as Record<string, unknown>;
  globals.Word = { run: async (batch: (c: unknown) => Promise<unknown>) => batch(ctx) };
  globals.Office = {
    context: {
      requirements: { isSetSupported: (_set: string, version: string) => version === "1.3" && supported13 },
    },
  };
}

afterEach(() => {
  const globals = globalThis as unknown as Record<string, unknown>;
  delete globals.Word;
  delete globals.Office;
  vi.restoreAllMocks();
});

describe("scanParagraphs 表内段落探测", () => {
  it("表内段落标 inTable=true，正文段落 false", async () => {
    stubWord([
      { uniqueLocalId: "p1", text: "正文段落。", styleBuiltIn: "Normal", style: "Normal", inTable: false },
      { uniqueLocalId: "cell1", text: "95%", styleBuiltIn: "Normal", style: "Normal", inTable: true },
    ]);
    const snapshot = await documentService.getDocumentSnapshot();
    expect(snapshot.paragraphs).toHaveLength(2);
    expect(snapshot.paragraphs[0].inTable).toBe(false);
    expect(snapshot.paragraphs[1].inTable).toBe(true);
  });

  it("宿主不支持 WordApi 1.3 → 全部降级 false，不抛错", async () => {
    stubWord(
      [
        { uniqueLocalId: "p1", text: "正文段落。", styleBuiltIn: "Normal", style: "Normal", inTable: true },
        { uniqueLocalId: "p2", text: "另一段。", styleBuiltIn: "Normal", style: "Normal", inTable: true },
      ],
      { supported13: false },
    );
    const snapshot = await documentService.getDocumentSnapshot();
    expect(snapshot.paragraphs.every((p) => p.inTable === false)).toBe(true);
  });

  it("探测抛错（load 失败）→ 降级 false，扫描不中断", async () => {
    stubWord(
      [
        { uniqueLocalId: "p1", text: "正文段落。", styleBuiltIn: "Normal", style: "Normal", inTable: true },
      ],
      { probeThrows: true },
    );
    const snapshot = await documentService.getDocumentSnapshot();
    expect(snapshot.paragraphs).toHaveLength(1);
    expect(snapshot.paragraphs[0].inTable).toBe(false);
  });

  it("缺失 uniqueLocalId 时回退序号 id 且 inTable 照常填充", async () => {
    stubWord([{ text: "无 id 段落。", styleBuiltIn: "Normal", style: "Normal", inTable: true }]);
    const snapshot = await documentService.getDocumentSnapshot();
    expect(snapshot.paragraphs[0].id).toBe("para-0");
    expect(snapshot.paragraphs[0].inTable).toBe(true);
  });
});

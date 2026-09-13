import { describe, it, expect } from "vitest";
import * as JSZip from "jszip";
import {
  readExcel,
  columnLettersToIndex,
  excelSerialToDate,
  parseSharedStrings,
} from "./read-excel";
import { buildWorkbook, generateTree } from "./generate-excel";
import { dateToExcelSerial } from "./util";
import {
  writeEquation,
  skipCell,
  createDateCell,
  createBackgroundCell,
} from "./util";
import type { IPage } from "./index";

/**
 * Builds an in-memory .xlsx buffer from the same pipeline the writer uses, so the
 * reader is exercised against genuine library output without touching disk.
 */
const toXlsxBuffer = async (pages: IPage[]): Promise<Buffer> => {
  const tree = generateTree(buildWorkbook(pages));
  // jszip's default export interop under CJS/vitest
  const Zip = (JSZip as any).default || (JSZip as any);
  const zip = new Zip();
  Object.entries(tree).forEach(([name, content]) => zip.file(name, content as string));
  return zip.generateAsync({ type: "nodebuffer" });
};

describe("read-excel helpers", () => {
  it("columnLettersToIndex inverts Excel column naming", () => {
    expect(columnLettersToIndex("A")).toBe(0);
    expect(columnLettersToIndex("Z")).toBe(25);
    expect(columnLettersToIndex("AA")).toBe(26);
    expect(columnLettersToIndex("AB")).toBe(27);
  });

  it("parseSharedStrings decodes entities and rich-text runs", () => {
    const xml = `<sst><si><t>a &amp; b</t></si><si><r><t>Hello </t></r><r><t>World</t></r></si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(["a & b", "Hello World"]);
  });

  it("excelSerialToDate inverts the writer's dateToExcelSerial (calendar day)", () => {
    for (const iso of ["2024-01-01", "2024-01-15", "2000-02-29", "2024-12-31", "1990-07-04"]) {
      const original = new Date(iso + "T00:00:00");
      const back = excelSerialToDate(dateToExcelSerial(original));
      expect([back.getFullYear(), back.getMonth(), back.getDate()]).toEqual([
        original.getFullYear(),
        original.getMonth(),
        original.getDate(),
      ]);
    }
  });
});

describe("readExcel round-trip", () => {
  it("reads strings and numbers back in order", async () => {
    const buf = await toXlsxBuffer([
      { title: "Data", content: [["name", "age"], ["alice", 30], ["bob", 25]] },
    ]);
    const wb = await readExcel(buf);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].title).toBe("Data");
    expect(wb.sheets[0].rows).toEqual([
      ["name", "age"],
      ["alice", 30],
      ["bob", 25],
    ]);
  });

  it("preserves multiple sheets and their names", async () => {
    const buf = await toXlsxBuffer([
      { title: "First", content: [["x"]] },
      { title: "Second", content: [["y"]] },
    ]);
    const wb = await readExcel(buf);
    expect(wb.sheets.map((s) => s.title)).toEqual(["First", "Second"]);
    expect(wb.sheets[1].rows[0][0]).toBe("y");
  });

  it("represents skipped cells as null gaps", async () => {
    const buf = await toXlsxBuffer([
      { title: "Gaps", content: [[1, skipCell(1), 2]] },
    ]);
    const wb = await readExcel(buf);
    expect(wb.sheets[0].rows[0]).toEqual([1, null, 2]);
  });

  it("round-trips date cells to Date objects", async () => {
    const buf = await toXlsxBuffer([
      { title: "Dates", content: [["start", createDateCell(new Date("2024-06-15T00:00:00"))]] },
    ]);
    const wb = await readExcel(buf);
    const value = wb.sheets[0].rows[0][1];
    expect(value).toBeInstanceOf(Date);
    const d = value as Date;
    expect([d.getFullYear(), d.getMonth(), d.getDate()]).toEqual([2024, 5, 15]);
  });

  it("returns date serials as numbers when cellDates is disabled", async () => {
    const buf = await toXlsxBuffer([
      { title: "Dates", content: [[createDateCell(new Date("2024-06-15T00:00:00"))]] },
    ]);
    const wb = await readExcel(buf, { cellDates: false });
    expect(typeof wb.sheets[0].rows[0][0]).toBe("number");
  });

  it("exposes formulas and preserves styled string values", async () => {
    const buf = await toXlsxBuffer([
      {
        title: "Calc",
        content: [
          [1, 2, writeEquation("SUM(A1,B1)")],
          [createBackgroundCell("styled", "#FFFF00")],
        ],
      },
    ]);
    const wb = await readExcel(buf);
    // Formula cell: no cached value is emitted, so the value surfaces the formula text.
    expect(wb.sheets[0].rows[0][0]).toBe(1);
    expect(wb.sheets[0].rows[0][1]).toBe(2);
    expect(wb.sheets[0].formulas?.[0][2]).toBe("SUM(A1,B1)");
    // Styled string keeps its text value.
    expect(wb.sheets[0].rows[1][0]).toBe("styled");
  });

  it("accepts a raw ArrayBuffer (browser-style input)", async () => {
    const buf = await toXlsxBuffer([{ title: "Buf", content: [["hi", 7]] }]);
    // Copy into a standalone ArrayBuffer, as a browser File/Blob read would yield.
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const wb = await readExcel(ab as ArrayBuffer);
    expect(wb.sheets[0].rows[0]).toEqual(["hi", 7]);
  });

  it("decodes XML-escaped string content", async () => {
    const buf = await toXlsxBuffer([
      { title: "Esc", content: [["a & b < c > d"]] },
    ]);
    const wb = await readExcel(buf);
    expect(wb.sheets[0].rows[0][0]).toBe("a & b < c > d");
  });
});

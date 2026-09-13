import { describe, it, expect } from "vitest";
import { parseCsv } from "./read-csv";

describe("parseCsv", () => {
  it("parses a simple grid", () => {
    expect(parseCsv("a,b,c\n1,2,3")).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
    ]);
  });

  it("handles quoted fields with embedded commas", () => {
    expect(parseCsv('a,"b,c",d')).toEqual([["a", "b,c", "d"]]);
  });

  it("handles escaped quotes inside quoted fields", () => {
    expect(parseCsv('"she said ""hi""",x')).toEqual([['she said "hi"', "x"]]);
  });

  it("handles embedded newlines inside quoted fields", () => {
    expect(parseCsv('"line1\nline2",b')).toEqual([["line1\nline2", "b"]]);
  });

  it("handles CRLF line endings", () => {
    expect(parseCsv("a,b\r\n1,2\r\n")).toEqual([
      ["a", "b"],
      ["1", "2"],
    ]);
  });

  it("keeps empty fields", () => {
    expect(parseCsv("a,,c")).toEqual([["a", "", "c"]]);
  });

  it("strips a UTF-8 BOM", () => {
    expect(parseCsv("﻿a,b")).toEqual([["a", "b"]]);
  });

  it("supports a custom delimiter", () => {
    expect(parseCsv("a;b;c", { delimiter: ";" })).toEqual([["a", "b", "c"]]);
  });

  it("drops the trailing empty row from a final newline", () => {
    expect(parseCsv("a\nb\n")).toEqual([["a"], ["b"]]);
  });

  it("keeps interior empty rows and the trailing empty row when disabled", () => {
    expect(parseCsv("a\n\n", { skipEmptyTrailingRow: false })).toEqual([["a"], [""]]);
  });
});

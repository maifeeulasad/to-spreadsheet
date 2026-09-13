/**
 * @fileoverview Spreadsheet reader (import) for .xlsx (OOXML) files
 * Parses an Excel workbook back into plain JavaScript values in both Node.js and
 * browser environments. This is the read counterpart to {@link generateExcel}.
 *
 * The parser is dependency-light: it unzips the package with JSZip (already a
 * dependency of the writer) and extracts the handful of parts we need with small
 * self-contained XML helpers. It reads workbooks produced by this library as well
 * as ordinary `.xlsx` files emitted by Excel and other tools.
 *
 * References (OOXML SpreadsheetML / ECMA-376 Part 1):
 * - Structure of a SpreadsheetML document:
 *   https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document
 * - Working with the shared string table (sst / <si>/<t>, cell t="s"):
 *   https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table
 * - Cell type ("t") and value ("v"): ECMA-376 Part 1, §18.3.1.4 (c) and §18.18.11 (ST_CellType)
 * - Number formats / built-in date format ids (14-22, 45-47, ...): ECMA-376 Part 1, §18.8.30 (numFmt)
 * - Excel 1900 date system & the intentional 1900 leap-year bug (serial 60):
 *   https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
 *
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

/**
 * Any value a parsed cell can hold.
 * - `string` for text (shared, inline or formula-string results)
 * - `number` for numeric cells
 * - `boolean` for boolean cells
 * - `Date` for cells carrying a date number format (when {@link IReadOptions.cellDates} is on)
 * - `null` for empty/blank cells
 */
type ReadCellValue = string | number | boolean | Date | null;

/**
 * A single parsed worksheet.
 * @interface IReadSheet
 */
interface IReadSheet {
  /** Sheet name as declared in the workbook. */
  title: string;
  /** Dense grid of cell values, row-major, gaps filled with `null`. */
  rows: ReadCellValue[][];
  /**
   * Sparse grid of formula strings (without the leading `=`), parallel to {@link rows}.
   * A cell is `null` when it carries no formula. Only present when the sheet has
   * at least one formula cell.
   */
  formulas?: (string | null)[][];
}

/**
 * A parsed workbook.
 * @interface IReadWorkbook
 */
interface IReadWorkbook {
  /** Worksheets in workbook (tab) order. */
  sheets: IReadSheet[];
}

/**
 * Options controlling how a workbook is parsed.
 * @interface IReadOptions
 */
interface IReadOptions {
  /**
   * Convert cells carrying a date number format into JavaScript `Date` objects.
   * When `false`, such cells come back as their raw Excel serial number.
   * @default true
   */
  cellDates?: boolean;
}

/**
 * Accepted input shapes for {@link readExcel}.
 * - `string` — a file path (Node.js only); read from disk.
 * - `Buffer` / `Uint8Array` / `ArrayBuffer` — raw file bytes (Node or browser).
 * - `Blob` — a browser `Blob`/`File`.
 */
type ReadExcelInput = string | Uint8Array | ArrayBuffer | Blob;

/**
 * Decodes the five predefined XML entities plus numeric character references.
 * `&amp;` is decoded last so that e.g. `&amp;lt;` round-trips to `&lt;`.
 * @internal
 */
const decodeXml = (text: string): string =>
  text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

/**
 * Reads a single XML attribute's value off a tag's attribute string.
 * Handles single- or double-quoted values and returns `undefined` when absent.
 * @internal
 */
const getAttr = (attrs: string, name: string): string | undefined => {
  const match = attrs.match(new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"|\\b${name}\\s*=\\s*'([^']*)'`));
  if (!match) return undefined;
  return decodeXml(match[1] !== undefined ? match[1] : match[2]);
};

/**
 * Converts an Excel column reference (letters of a cell ref like `AB` in `AB12`)
 * to a zero-based column index. Inverse of the writer's column naming.
 * @internal
 * @example columnLettersToIndex("A") // 0
 * @example columnLettersToIndex("AA") // 26
 */
const columnLettersToIndex = (letters: string): number => {
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64); // 'A' -> 1
  }
  return index - 1;
};

/**
 * Excel epoch used by the writer's `dateToExcelSerial`. Converting back accounts
 * for the historical 1900 leap-year bug the same way the writer introduced it.
 * @internal
 */
const excelSerialToDate = (serial: number): Date => {
  const excelEpoch = new Date(1900, 0, 1);
  // The writer adds +2 for dates on/after 1900-02-29 (the phantom leap day) and
  // +1 before it; subtract the matching offset. serial >= 60 corresponds to that
  // phantom day, mirroring the writer's `date >= 1900-02-29 ? 2 : 1`.
  const offset = serial >= 60 ? 2 : 1;
  const millis = excelEpoch.getTime() + (serial - offset) * 24 * 60 * 60 * 1000;
  return new Date(millis);
};

/**
 * Built-in Excel number-format ids that denote dates/times.
 * @see ECMA-376, §18.8.30 (numFmt) built-in formats.
 * @internal
 */
const BUILTIN_DATE_FORMAT_IDS = new Set<number>([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
]);

/**
 * Heuristic: does a custom format code describe a date/time?
 * Strips literal text/quoted sections and colour/condition tokens, then looks for
 * date/time field characters.
 * @internal
 */
const isDateFormatCode = (code: string): boolean => {
  const stripped = code
    .replace(/\[[^\]]*\]/g, "") // [Red], [$-409], conditions
    .replace(/"[^"]*"/g, "") // quoted literals
    .replace(/\\./g, ""); // escaped literals
  return /[yYdDhHsS]/.test(stripped) || /\bmm?m?m?\b/.test(stripped) && /[yYdDhH]/.test(code);
};

/**
 * Parses `xl/styles.xml` into a lookup of cell-format (`s` attribute) index →
 * whether that format renders a date. Used to decide when a numeric cell should
 * be reinterpreted as a `Date`.
 * @internal
 */
const parseDateStyleFlags = (stylesXml: string | undefined): boolean[] => {
  if (!stylesXml) return [];

  // Custom number formats: numFmtId -> formatCode
  const dateNumFmtIds = new Set<number>(BUILTIN_DATE_FORMAT_IDS);
  const numFmtRe = /<numFmt\b([^>]*)\/?>/g;
  let numFmtMatch: RegExpExecArray | null;
  while ((numFmtMatch = numFmtRe.exec(stylesXml)) !== null) {
    const id = Number(getAttr(numFmtMatch[1], "numFmtId"));
    const code = getAttr(numFmtMatch[1], "formatCode");
    if (!Number.isNaN(id) && code && isDateFormatCode(code)) {
      dateNumFmtIds.add(id);
    }
  }

  // cellXfs: ordered list of <xf numFmtId="..."> — index is the cell `s` value.
  const cellXfsBlock = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (!cellXfsBlock) return [];
  const flags: boolean[] = [];
  const xfRe = /<xf\b([^>]*?)(?:\/>|>[\s\S]*?<\/xf>)/g;
  let xfMatch: RegExpExecArray | null;
  while ((xfMatch = xfRe.exec(cellXfsBlock[1])) !== null) {
    const numFmtId = Number(getAttr(xfMatch[1], "numFmtId"));
    flags.push(!Number.isNaN(numFmtId) && dateNumFmtIds.has(numFmtId));
  }
  return flags;
};

/**
 * Parses `xl/sharedStrings.xml` into an array of decoded strings, indexed as the
 * `t="s"` cell values reference them. Concatenates rich-text runs within a `<si>`.
 * @internal
 */
const parseSharedStrings = (sharedStringsXml: string | undefined): string[] => {
  if (!sharedStringsXml) return [];
  const strings: string[] = [];
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
  let siMatch: RegExpExecArray | null;
  while ((siMatch = siRe.exec(sharedStringsXml)) !== null) {
    const inner = siMatch[1] || "";
    // Concatenate every <t> (a plain string is one <t>; rich text is many <r><t>).
    let text = "";
    const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g;
    let tMatch: RegExpExecArray | null;
    while ((tMatch = tRe.exec(inner)) !== null) {
      text += decodeXml(tMatch[1] || "");
    }
    strings.push(text);
  }
  return strings;
};

/**
 * Maps relationship ids (`r:id`) to their targets from `xl/_rels/workbook.xml.rels`,
 * normalised to package paths under `xl/`.
 * @internal
 */
const parseWorkbookRels = (relsXml: string | undefined): Map<string, string> => {
  const map = new Map<string, string>();
  if (!relsXml) return map;
  const relRe = /<Relationship\b([^>]*)\/?>/g;
  let match: RegExpExecArray | null;
  while ((match = relRe.exec(relsXml)) !== null) {
    const id = getAttr(match[1], "Id");
    let target = getAttr(match[1], "Target");
    if (!id || !target) continue;
    target = target.replace(/^\//, "").replace(/^xl\//, "");
    map.set(id, `xl/${target}`);
  }
  return map;
};

/**
 * Ordered worksheet descriptor read from `xl/workbook.xml`.
 * @internal
 */
interface ISheetRef {
  name: string;
  rId?: string;
}

/**
 * Parses `xl/workbook.xml` into the ordered list of sheets (name + relationship id).
 * @internal
 */
const parseWorkbookSheets = (workbookXml: string | undefined): ISheetRef[] => {
  if (!workbookXml) return [];
  const sheets: ISheetRef[] = [];
  const sheetRe = /<sheet\b([^>]*)\/?>/g;
  let match: RegExpExecArray | null;
  while ((match = sheetRe.exec(workbookXml)) !== null) {
    const name = getAttr(match[1], "name") ?? `Sheet${sheets.length + 1}`;
    const rId = getAttr(match[1], "r:id") ?? getAttr(match[1], "id");
    sheets.push({ name, rId });
  }
  return sheets;
};

/**
 * Parses one worksheet part into a value grid (and formula grid when present).
 * @internal
 */
const parseWorksheet = (
  sheetXml: string,
  sharedStrings: string[],
  dateStyleFlags: boolean[],
  cellDates: boolean
): { rows: ReadCellValue[][]; formulas?: (string | null)[][] } => {
  const rows: ReadCellValue[][] = [];
  const formulas: (string | null)[][] = [];
  let sawFormula = false;

  const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let rowMatch: RegExpExecArray | null;
  let sequentialRowIndex = -1;
  while ((rowMatch = rowRe.exec(sheetXml)) !== null) {
    sequentialRowIndex++;
    const rAttr = getAttr(rowMatch[1], "r");
    const rowIndex = rAttr ? Number(rAttr) - 1 : sequentialRowIndex;
    const rowInner = rowMatch[2];

    const rowValues: ReadCellValue[] = [];
    const rowFormulas: (string | null)[] = [];

    let sequentialColIndex = -1;
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cellMatch: RegExpExecArray | null;
    while ((cellMatch = cellRe.exec(rowInner)) !== null) {
      sequentialColIndex++;
      const attrs = cellMatch[1];
      const inner = cellMatch[2] || "";
      const ref = getAttr(attrs, "r");
      const colLetters = ref ? ref.replace(/[0-9]/g, "") : "";
      const colIndex = colLetters ? columnLettersToIndex(colLetters) : sequentialColIndex;
      const type = getAttr(attrs, "t") || "n";
      const styleIndex = Number(getAttr(attrs, "s") || "0");

      // Formula (writer emits <f aca="false">EXPR</f>).
      const fMatch = inner.match(/<f\b[^>]*>([\s\S]*?)<\/f>|<f\b[^>]*\/>/);
      const formula = fMatch ? decodeXml(fMatch[1] || "") : null;
      if (formula) sawFormula = true;

      // Value.
      const vMatch = inner.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
      const rawValue = vMatch ? decodeXml(vMatch[1]) : undefined;

      let value: ReadCellValue = null;
      if (type === "s") {
        // Shared string index.
        const idx = Number(rawValue);
        value = Number.isNaN(idx) ? "" : sharedStrings[idx] ?? "";
      } else if (type === "inlineStr") {
        const isMatch = inner.match(/<is\b[^>]*>([\s\S]*?)<\/is>/);
        let text = "";
        if (isMatch) {
          const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
          let tMatch: RegExpExecArray | null;
          while ((tMatch = tRe.exec(isMatch[1])) !== null) text += decodeXml(tMatch[1]);
        }
        value = text;
      } else if (type === "str") {
        // Formula string result.
        value = rawValue ?? (formula ?? "");
      } else if (type === "b") {
        value = rawValue === "1" || rawValue === "true";
      } else if (type === "e") {
        value = rawValue ?? null; // error text, e.g. #DIV/0!
      } else {
        // Numeric (t="n" or unspecified).
        if (rawValue === undefined || rawValue === "") {
          // A formula cell with no cached value: surface the formula text.
          value = formula !== null ? formula : null;
        } else {
          const num = Number(rawValue);
          if (Number.isNaN(num)) {
            value = rawValue;
          } else if (cellDates && dateStyleFlags[styleIndex]) {
            value = excelSerialToDate(num);
          } else {
            value = num;
          }
        }
      }

      rowValues[colIndex] = value;
      rowFormulas[colIndex] = formula;
    }

    rows[rowIndex] = rowValues;
    formulas[rowIndex] = rowFormulas;
  }

  // Normalise: fill row/column gaps with null so the grid is rectangular.
  const width = rows.reduce((max, row) => Math.max(max, row ? row.length : 0), 0);
  for (let r = 0; r < rows.length; r++) {
    if (!rows[r]) rows[r] = [];
    for (let c = 0; c < width; c++) {
      if (rows[r][c] === undefined) rows[r][c] = null;
      if (!formulas[r]) formulas[r] = [];
      if (formulas[r][c] === undefined) formulas[r][c] = null;
    }
  }

  return sawFormula ? { rows, formulas } : { rows };
};

/**
 * Normalises any supported input into something JSZip's `loadAsync` accepts.
 * Node file paths are read from disk; every other shape is passed through.
 * @internal
 */
const toZipInput = async (input: ReadExcelInput): Promise<Uint8Array | ArrayBuffer | Blob> => {
  if (typeof input === "string") {
    // Treat a string as a filesystem path (Node.js only).
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const fs = require("fs");
    return fs.readFileSync(input) as Uint8Array;
  }
  return input;
};

/**
 * Reads an `.xlsx` workbook and returns its sheets as plain JavaScript values.
 * Works in both Node.js and the browser.
 *
 * @param input - A file path (Node), raw bytes, or a browser `Blob`/`File`.
 * @param options - Parsing options; see {@link IReadOptions}.
 * @returns A promise resolving to the parsed {@link IReadWorkbook}.
 *
 * @example
 * // Node.js — read from disk
 * const wb = await readExcel("./report.xlsx");
 * console.log(wb.sheets[0].title, wb.sheets[0].rows);
 *
 * @example
 * // Browser — read a File from an <input type="file">
 * const wb = await readExcel(file);
 */
const readExcel = async (
  input: ReadExcelInput,
  options: IReadOptions = {}
): Promise<IReadWorkbook> => {
  const cellDates = options.cellDates !== false;

  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(await toZipInput(input));

  const readPart = async (path: string): Promise<string | undefined> => {
    const file = zip.file(path);
    return file ? await file.async("string") : undefined;
  };

  const [workbookXml, relsXml, sharedStringsXml, stylesXml] = await Promise.all([
    readPart("xl/workbook.xml"),
    readPart("xl/_rels/workbook.xml.rels"),
    readPart("xl/sharedStrings.xml"),
    readPart("xl/styles.xml"),
  ]);

  const sharedStrings = parseSharedStrings(sharedStringsXml);
  const dateStyleFlags = parseDateStyleFlags(stylesXml);
  const rels = parseWorkbookRels(relsXml);
  const sheetRefs = parseWorkbookSheets(workbookXml);

  const sheets: IReadSheet[] = [];
  for (let i = 0; i < sheetRefs.length; i++) {
    const ref = sheetRefs[i];
    // Resolve the worksheet part: prefer the relationship target, fall back to
    // the positional sheetN.xml the writer produces.
    let path = ref.rId ? rels.get(ref.rId) : undefined;
    if (!path || !zip.file(path)) path = `xl/worksheets/sheet${i + 1}.xml`;
    const sheetXml = await readPart(path);
    if (sheetXml === undefined) {
      sheets.push({ title: ref.name, rows: [] });
      continue;
    }
    const parsed = parseWorksheet(sheetXml, sharedStrings, dateStyleFlags, cellDates);
    sheets.push({ title: ref.name, ...parsed });
  }

  return { sheets };
};

export {
  readExcel,
  ReadCellValue,
  IReadSheet,
  IReadWorkbook,
  IReadOptions,
  ReadExcelInput,
  // exported for unit testing and advanced reuse
  columnLettersToIndex,
  excelSerialToDate,
  parseSharedStrings,
  parseWorksheet,
};

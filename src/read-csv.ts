/**
 * @fileoverview CSV reader (import)
 * A small, dependency-free RFC 4180-style CSV parser that works in both Node.js
 * and the browser. Handles quoted fields, escaped quotes (`""`), embedded commas
 * and newlines, and both `\n` and `\r\n` line endings.
 *
 * Reference:
 * - RFC 4180 - Common Format and MIME Type for CSV Files:
 *   https://www.rfc-editor.org/rfc/rfc4180
 *
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

/**
 * Options controlling CSV parsing.
 * @interface ICsvOptions
 */
interface ICsvOptions {
  /**
   * Field delimiter.
   * @default ","
   */
  delimiter?: string;
  /**
   * Drop a trailing empty row produced by a final newline.
   * @default true
   */
  skipEmptyTrailingRow?: boolean;
}

/**
 * Parses CSV text into a 2D array of strings (rows of fields).
 *
 * @param text - The raw CSV content.
 * @param options - Parsing options; see {@link ICsvOptions}.
 * @returns A row-major array of string fields.
 *
 * @example
 * parseCsv('a,b\n1,"2,3"') // [["a","b"],["1","2,3"]]
 */
const parseCsv = (text: string, options: ICsvOptions = {}): string[][] => {
  const delimiter = options.delimiter ?? ",";
  const skipEmptyTrailingRow = options.skipEmptyTrailingRow !== false;
  const delim = delimiter.charAt(0);

  // Strip a UTF-8 BOM if present.
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let i = 0;
  const length = text.length;

  const endField = () => {
    row.push(field);
    field = "";
  };
  const endRow = () => {
    endField();
    rows.push(row);
    row = [];
  };

  while (i < length) {
    const char = text[i];

    if (inQuotes) {
      if (char === '"') {
        if (text[i + 1] === '"') {
          field += '"'; // escaped quote
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += char;
      i++;
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (char === delim) {
      endField();
      i++;
      continue;
    }
    if (char === "\n") {
      endRow();
      i++;
      continue;
    }
    if (char === "\r") {
      // Swallow \r, and the following \n if present (\r\n).
      endRow();
      if (text[i + 1] === "\n") i += 2;
      else i++;
      continue;
    }
    field += char;
    i++;
  }

  // Flush the final field/row (file not ending in a newline).
  if (field.length > 0 || row.length > 0) {
    endRow();
  }

  if (
    skipEmptyTrailingRow &&
    rows.length > 0 &&
    rows[rows.length - 1].length === 1 &&
    rows[rows.length - 1][0] === ""
  ) {
    rows.pop();
  }

  return rows;
};

export { parseCsv, ICsvOptions };

/**
 * @fileoverview Core Excel generation functionality
 * Handles the conversion of data structures to Excel format and manages the creation
 * of Excel workbook files in both Node.js and browser environments
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { generateContentTypesXml } from "./content-types.xml";
import { generateRels } from "./_rels/.rels";
import { generateAppXml } from "./docProps/app.xml";
import { generateCoreXml } from "./docProps/core.xml";
import { generateWorkBookXmlRels } from "./xl/_rels/workbook.xml.rels";
import { generateSharedStrings } from "./xl/sharedStrings.xml";
import { generateStyleXml } from "./xl/styles.xml";
import { generateTheme1 } from "./xl/theme/theme1.xml";
import { generateWorkBookXml } from "./xl/workbook.xml";
import { generateSheetXml } from "./xl/worksheets/sheet.xml";
import { ICellType, IPage, ISheet, IWorkbook, IBorder } from "./index";
import { Equation, SkipCell, getBorderKey } from "./util";

/**
 * Generates the complete XML file structure for an Excel workbook
 * Analyzes all cells for styling information and creates the appropriate XML files
 * @param {IWorkbook} workbook - The workbook data structure
 * @returns {Object} Object containing all XML files needed for the Excel workbook
 * @internal
 */
const generateTree = (workbook: IWorkbook) => {
  // Collect all unique border styles from the workbook
  const borderStyles = new Map<string, IBorder>();
  borderStyles.set("none", {}); // Default no-border style
  
  // Check if workbook contains any date cells
  let hasDateCells = false;
  
  workbook.sheets.forEach(sheet => {
    sheet.rows.forEach(row => {
      row.cells.forEach(cell => {
        // Check for date cells
        if (cell.type === ICellType.date) {
          hasDateCells = true;
        }
        
        // Collect border styles
        if ('style' in cell && cell.style?.border) {
          const borderKey = getBorderKey(cell.style.border);
          borderStyles.set(borderKey, cell.style.border);
        }
      });
    });
  });

  return {
    "[Content_Types].xml": generateContentTypesXml(workbook),
    "_rels/.rels": generateRels(),
    "docProps/app.xml": generateAppXml(workbook),
    "docProps/core.xml": generateCoreXml({}),
    "xl/_rels/workbook.xml.rels": generateWorkBookXmlRels(workbook),
    "xl/sharedStrings.xml": generateSharedStrings(workbook),
    "xl/styles.xml": generateStyleXml(borderStyles, hasDateCells),
    "xl/theme/theme1.xml": generateTheme1(),
    "xl/workbook.xml": generateWorkBookXml(workbook),
    ...workbook.sheets.reduce((acc, sheet, idx) => ({ ...acc, [`xl/worksheets/sheet${idx + 1}.xml`]: generateSheetXml(sheet, borderStyles, hasDateCells) }), {})
  };
};

/**
 * Enum representing different execution environments
 * @enum {number}
 */
enum EnvironmentType {
  /** Node.js server environment - generates files using filesystem */
  NODE,
  /** Browser environment - generates files using JSZip and triggers download */
  BROWSER,
}

/**
 * Main function to generate Excel spreadsheet files
 * Converts input data to Excel format and outputs .xlsx file in the specified environment
 * @param {IPage[]} dump - Array of worksheet data
 * @param {EnvironmentType} environmentType - Target environment (Node.js or Browser)
 * @returns {Promise<void>} Promise that resolves when file generation is complete
 * @example
 * // Generate Excel file in Node.js
 * generateExcel(data, EnvironmentType.NODE);
 * 
 * // Generate Excel file in browser (triggers download)
 * generateExcel(data, EnvironmentType.BROWSER);
 */
const generateExcel = (dump: IPage[], environmentType: EnvironmentType = EnvironmentType.NODE): Promise<void> => {
  const strings: string[] = [];
  
  // Convert input pages to internal workbook structure
  const sheets: ISheet[] = dump.map(({ title, content }) => {
    const rows = content.map(row => {
      const cells: any[] = [];
      
      // Process each cell in the row
      row.forEach(content => {
        if (typeof content === 'number') {
          // Handle numeric values
          cells.push({ type: ICellType.number, value: content });
        } else if (typeof content === 'string') {
          // Handle string values - add to shared strings table
          const type = ICellType.string;
          let value = strings.indexOf(content);

          if (value === -1) {
            strings.push(content);
            value = strings.length - 1;
          }
          cells.push({ type: ICellType.string, value });
        } else if (content instanceof SkipCell) {
          // Handle skip cell instructions
          for (let i = 0; i < content.getSkipCell(); i++) {
            cells.push({ type: ICellType.skip, value: undefined });
          }
        } else if (content instanceof Equation) {
          // Handle Excel formulas
          cells.push({ type: ICellType.equation, value: content });
        } else if (content && typeof content === 'object' && 'type' in content) {
          // Handle pre-built cell objects with styling
          const cell = content as any;
          if (cell.type === ICellType.string && typeof cell.value === 'string') {
            // Convert string value to string index for styled string cells
            let stringIndex = strings.indexOf(cell.value);
            if (stringIndex === -1) {
              strings.push(cell.value);
              stringIndex = strings.length - 1;
            }
            cells.push({ ...cell, value: stringIndex });
          } else {
            cells.push(cell);
          }
        } else {
          // Default to skip cell for undefined/null values
          cells.push({ type: ICellType.skip });
        }
      });

      return { cells };
    });

    return { title, rows };
  });

  const workbook: IWorkbook = {
    sheets,
    strings,
    filename: "tem.xlsx"
  };

  if (environmentType === EnvironmentType.BROWSER) {
    return generateExcelWorkbookBrowser(workbook);
  } else {
    return generateExcelWorkbookNode(workbook);
  }
}

/**
 * Generates Excel workbook file in Node.js environment
 * Uses the filesystem and archiver library to create .xlsx files
 * @param {IWorkbook} workbook - Complete workbook data structure
 * @returns {Promise<void>} Promise that resolves when file is written
 * @internal
 */
const generateExcelWorkbookNode = (workbook: IWorkbook): Promise<void> => {
  return new Promise((resolve, reject) => {

    const fs = require('fs');
    const archiver = require('archiver');

    const output = fs.createWriteStream(`${__dirname}/${workbook.filename}.xlsx`);
    const archive = archiver("zip", {
      zlib: { level: 9 },
    });

    output.on("close", () => {
      console.debug(archive.pointer() + " total bytes");
      console.debug(
        "archiver has been finalized and the output file descriptor has closed."
      );
      resolve();  // Resolve the promise when the archive is completed and closed
    });

    output.on("end", () => {
      console.debug("Data has been drained");
    });

    archive.on("warning", (err: any) => {
      if (err.code === "ENOENT") {
        // log warning
      } else {
        // throw error, but also reject the promise
        reject(err);
      }
    });

    archive.on("error", (err: any) => {
      reject(err);  // Reject the promise on error
    });

    archive.pipe(output);

    // Add all generated XML files to the archive
    Object.entries(generateTree(workbook)).map(([filename, fileContent]) => {
      archive.append(fileContent, { name: filename });
    });

    archive.finalize();
  });
};

/**
 * Generates Excel workbook file in browser environment  
 * Uses JSZip library to create .xlsx files and triggers download
 * @param {IWorkbook} workbook - Complete workbook data structure
 * @returns {Promise<void>} Promise that resolves when download is triggered
 * @internal
 */
const generateExcelWorkbookBrowser = (workbook: IWorkbook): Promise<void> => {
  return new Promise((resolve, reject) => {
    try {

      const JSZip = require('jszip');
      const { saveAs } = require('file-saver');

      const zip = new JSZip();
      const tree = generateTree(workbook);

      // Add all XML files to the zip
      Object.entries(tree).forEach(([filename, fileContent]) => {
        zip.file(filename, fileContent);
      });

      // Generate blob and trigger download
      zip.generateAsync({ type: "blob" }).then((blob: Blob) => {
        saveAs(blob, `${workbook.filename}.xlsx`);
        resolve();
      });

    } catch (error) {
      reject(error);
    }
  });
};

/**
 * Export the main generation function and environment type enum
 * These are the primary exports used by consuming applications
 */
export { generateExcel, EnvironmentType };
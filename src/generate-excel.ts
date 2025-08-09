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

const generateTree = (workbook: IWorkbook) => {
  // Collect all unique border styles from the workbook
  const borderStyles = new Map<string, IBorder>();
  borderStyles.set("none", {}); // Default no-border style
  
  workbook.sheets.forEach(sheet => {
    sheet.rows.forEach(row => {
      row.cells.forEach(cell => {
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
    "xl/styles.xml": generateStyleXml(borderStyles),
    "xl/theme/theme1.xml": generateTheme1(),
    "xl/workbook.xml": generateWorkBookXml(workbook),
    ...workbook.sheets.reduce((acc, sheet, idx) => ({ ...acc, [`xl/worksheets/sheet${idx + 1}.xml`]: generateSheetXml(sheet, borderStyles) }), {})
  };
};

enum EnvironmentType {
  NODE,
  BROWSER,
}

const generateExcel = (dump: IPage[], environmentType: EnvironmentType = EnvironmentType.NODE): Promise<void> => {
  const strings: string[] = [];
  const sheets: ISheet[] = dump.map(({ title, content }) => {
    const rows = content.map(row => {
      const cells: any[] = [];
      
      row.forEach(content => {
        if (typeof content === 'number') {
          cells.push({ type: ICellType.number, value: content });
        } else if (typeof content === 'string') {
          const type = ICellType.string;
          let value = strings.indexOf(content);

          if (value === -1) {
            strings.push(content);
            value = strings.length - 1;
          }
          cells.push({ type: ICellType.string, value });
        } else if (content instanceof SkipCell) {
          for (let i = 0; i < content.getSkipCell(); i++) {
            cells.push({ type: ICellType.skip, value: undefined });
          }
        } else if (content instanceof Equation) {
          cells.push({ type: ICellType.equation, value: content });
        } else if (content && typeof content === 'object' && 'type' in content) {
          // Handle pre-built cell objects with styling
          const cell = content as any;
          if (cell.type === ICellType.string && typeof cell.value === 'string') {
            // Convert string value to string index
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

    Object.entries(generateTree(workbook)).map(([filename, fileContent]) => {
      archive.append(fileContent, { name: filename });
    });

    archive.finalize();
  });
};

const generateExcelWorkbookBrowser = (workbook: IWorkbook): Promise<void> => {
  return new Promise((resolve, reject) => {
    try {

      const JSZip = require('jszip');
      const { saveAs } = require('file-saver');

      const zip = new JSZip();
      const tree = generateTree(workbook);

      Object.entries(tree).forEach(([filename, fileContent]) => {
        zip.file(filename, fileContent);
      });

      zip.generateAsync({ type: "blob" }).then((blob: Blob) => {
        saveAs(blob, `${workbook.filename}.xlsx`);
        resolve();
      });

    } catch (error) {
      reject(error);
    }
  });
};

export { generateExcel, EnvironmentType };
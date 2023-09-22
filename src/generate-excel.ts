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
import { ICellType, IPage, ISheet, IWorkbook } from "./index"

const fs = require("fs");
const archiver = require("archiver");

const generateTree = (workbook: IWorkbook) => {
  return {
    "[Content_Types].xml": generateContentTypesXml(workbook),
    "_rels/.rels": generateRels(),
    "docProps/app.xml": generateAppXml(workbook),
    "docProps/core.xml": generateCoreXml({}),
    "xl/_rels/workbook.xml.rels": generateWorkBookXmlRels(workbook),
    "xl/sharedStrings.xml": generateSharedStrings(workbook),
    "xl/styles.xml": generateStyleXml(),
    "xl/theme/theme1.xml": generateTheme1(),
    "xl/workbook.xml": generateWorkBookXml(workbook),
    ...workbook.sheets.reduce((acc, sheet, idx) => ({ ...acc, [`xl/worksheets/sheet${idx + 1}.xml`]: generateSheetXml(sheet) }), {})
  };
};


const generateExcel = (dump: IPage[]): Promise<void> => {
  const strings: string[] = []
  const sheets: ISheet[] = dump.map(({ title, content }) => {
    const rows = content.map(row => {
      const cells = row.map(content => {
        const isString = typeof content === 'string';
        const type = isString ? ICellType.string : ICellType.number;
        let value = isString ? strings.indexOf(content) : content;

        if (isString && value === -1) {
          strings.push(content);
          value = strings.length - 1;
        }

        return { type, value };
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

  return generateExcelWorkbook(workbook)
}

const generateExcelWorkbook = (workbook: IWorkbook): Promise<void> => {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(`${__dirname}/${workbook.filename}.xlsx`);
    const archive = archiver("zip", {
      zlib: { level: 9 },
    });

    output.on("close", function () {
      console.debug(archive.pointer() + " total bytes");
      console.debug(
        "archiver has been finalized and the output file descriptor has closed."
      );
      resolve();  // Resolve the promise when the archive is completed and closed
    });

    output.on("end", function () {
      console.debug("Data has been drained");
    });

    archive.on("warning", function (err: any) {
      if (err.code === "ENOENT") {
        // log warning
      } else {
        // throw error, but also reject the promise
        reject(err);
      }
    });

    archive.on("error", function (err: any) {
      reject(err);  // Reject the promise on error
    });

    archive.pipe(output);

    Object.entries(generateTree(workbook)).map(([filename, fileContent]) => {
      archive.append(fileContent, { name: filename });
    });

    archive.finalize();
  });
};

export { generateExcel, generateExcelWorkbook };

import { generateRels } from "./_rels/.rels";
import { generateAppXml } from "./docProps/app.xml";
import { generateCoreXml } from "./docProps/core.xml";
import { generateContentTypesXml } from "./content-types.xml";
import { generateWorkBookXmlRels } from "./xl/_rels/workbook.xml.rels";
import { generateTheme1 } from "./xl/theme/theme1.xml";
import { generateSheet1Xml, ICellEntry } from "./xl/worksheets/sheet1.xml";
import { generateStyleXml } from "./xl/styles.xml";
import { generateWorkBookXml } from "./xl/workbook.xml";

const fs = require("fs");
const archiver = require("archiver");

const generateTree = (data: ICellEntry[][]) => {
  return {
    "[Content_Types].xml": generateContentTypesXml(),
    "_rels/.rels": generateRels(),
    "docProps/app.xml": generateAppXml(),
    "docProps/core.xml": generateCoreXml(),
    "xl/_rels/workbook.xml.rels": generateWorkBookXmlRels(),
    "xl/theme/theme1.xml": generateTheme1(),
    "xl/worksheets/sheet1.xml": generateSheet1Xml(data),
    "xl/styles.xml": generateStyleXml(),
    "xl/workbook.xml": generateWorkBookXml(),
  };
};

const generateExcel = (data: ICellEntry[][]) => {
  const output = fs.createWriteStream(__dirname + "/example.xlsx");
  const archive = archiver("zip", {
    zlib: { level: 9 },
  });

  output.on("close", function () {
    console.debug(archive.pointer() + " total bytes");
    console.debug(
      "archiver has been finalized and the output file descriptor has closed."
    );
  });

  output.on("end", function () {
    console.debug("Data has been drained");
  });

  archive.on("warning", function (err: any) {
    if (err.code === "ENOENT") {
      // log warning
    } else {
      // throw error
      throw err;
    }
  });

  archive.on("error", function (err: any) {
    throw err;
  });

  archive.pipe(output);

  Object.entries(generateTree(data)).map((value) => {
    archive.append(value[1], { name: value[0] });
  });

  archive.finalize();
};

export { generateExcel };

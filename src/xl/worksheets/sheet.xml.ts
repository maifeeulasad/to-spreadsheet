import { ISheet, ICellType } from "../..";
import { rowColumnToVbPosition, indexToVbIndex, calculateExtant } from '../../util'


const generateSheetXml = (
  sheet: ISheet
) => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr filterMode="false">
    <pageSetUpPr fitToPage="false"/>
  </sheetPr>
  <dimension ref="A1:${calculateExtant(sheet.rows)}"/>
  <sheetViews>
    <sheetView tabSelected="1" zoomScale="79" zoomScaleNormal="79" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="12.8"></sheetFormatPr>
  <cols>
    <col max="1025" min="1" style="0" width="11.52"/>
  </cols>
  <sheetData>
    ${(sheet.rows).map((row, indexCol) => `
      <row r="${indexToVbIndex(indexCol)}">
        ${(row.cells).map((cell, indexRow) => cell.type === ICellType.skip
    ? ""
    : `
          <c r="${rowColumnToVbPosition(indexRow, indexCol)}" t="${cell.type}">
            <v>${cell.value}</v>
          </c>
          `).join('\n')
    }
      </row>
      `).join('\n')
    }
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
`};

export { generateSheetXml };

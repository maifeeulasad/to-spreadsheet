import { ISheet, ICellType } from "../..";
import { rowColumnToVbPosition, indexToVbIndex, calculateExtant, Equation } from '../../util'


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
  ${sheet.rows.map((row, rowIndex) => {
    let rowContent = '';
    rowContent += `<row r="${indexToVbIndex(rowIndex)}">\n`;

    row.cells.forEach((cell, cellIndex) => {
      const cellType = cell.type;

      if (cellType !== ICellType.skip) {
        const cellPosition = rowColumnToVbPosition(cellIndex, rowIndex);
        const cellValue = cell.value || '';

        if (cell.type === ICellType.equation) {

          // todo: write now I'm restricting to use function which will only return number
          rowContent += `
              <c r="${cellPosition}" t="n">
                <f aca="false">${cell.value.getEquation()}</f>
              </c>\n`;

        } else {
          
          rowContent += `
              <c r="${cellPosition}" t="${cellType}">
                <v>${cellValue}</v>
              </c>\n`;
        }
      }
    });

    rowContent += '</row>\n';
    return rowContent;
  }).join('')
    }
</sheetData>

  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
`};

export { generateSheetXml };
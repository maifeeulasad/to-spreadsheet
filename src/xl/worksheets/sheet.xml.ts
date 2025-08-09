/**
 * @fileoverview Excel worksheet sheet.xml generation
 * Handles the generation of Excel worksheet XML including cell data, positioning, and styling
 * This file creates the individual sheet.xml files that contain all the actual cell data and formatting
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { ISheet, ICellType, IBorder } from "../..";
import { rowColumnToVbPosition, indexToVbIndex, calculateExtant, Equation, getBorderKey, dateToExcelSerial } from '../../util'

/**
 * Generates the complete XML content for an Excel worksheet
 * Creates worksheet XML with cell data, styling references, and proper Excel structure
 * @param {ISheet} sheet - Sheet data containing rows and cells
 * @param {Map<string, IBorder>} borderStyles - Map of unique border styles used in the workbook
 * @param {boolean} hasDateCells - Whether the workbook contains date cells requiring special formatting
 * @returns {string} Complete XML content for sheet.xml file
 * @internal
 */
const generateSheetXml = (
  sheet: ISheet,
  borderStyles: Map<string, IBorder>,
  hasDateCells: boolean = false
) => {
  /**
   * Create a reverse mapping from border style to index
   * This allows us to reference border styles by index in cell styling
   * @type {Map<string, number>}
   */
  const borderStyleToIndex = new Map<string, number>();
  Array.from(borderStyles.keys()).forEach((key, index) => {
    borderStyleToIndex.set(key, index);
  });

  /**
   * Complete XML template for Excel worksheet
   * Includes worksheet metadata, dimensions, views, and all cell data
   * Follows OpenXML specification for Excel worksheets
   */
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
    // Generate opening row tag with proper row index
    rowContent += `<row r="${indexToVbIndex(rowIndex)}">\n`;

    /**
     * Process each cell in the row
     * Handles different cell types (skip, equation, regular values)
     * Applies appropriate styling and positioning
     */
    row.cells.forEach((cell, cellIndex) => {
      const cellType = cell.type;

      // Skip cells marked as ICellType.skip (for merged cells, etc.)
      if (cellType !== ICellType.skip) {
        // Convert array indices to Excel position (e.g., A1, B2)
        const cellPosition = rowColumnToVbPosition(cellIndex, rowIndex);
        const cellValue = cell.value || '';
        
        /**
         * Determine the style index for this cell
         * Style index references the border style in styles.xml
         * For date cells with date formatting enabled, add offset to get date format styles
         * Default to 0 (no border) if no style is specified
         */
        let styleIndex = 0; // Default to no-border style
        let isDateCell = cell.type === ICellType.date;
        
        if ('style' in cell && cell.style?.border) {
          const borderKey = getBorderKey(cell.style.border);
          const baseBorderIndex = borderStyleToIndex.get(borderKey) || 0;
          
          // If this is a date cell and we have date formatting, use the date format styles
          if (isDateCell && hasDateCells) {
            styleIndex = baseBorderIndex + borderStyles.size; // Offset for date formats
          } else {
            styleIndex = baseBorderIndex;
          }
        } else if (isDateCell && hasDateCells) {
          // Date cell with no border but needs date formatting
          styleIndex = borderStyles.size; // First date format style (no border)
        }

        /**
         * Handle different cell types with appropriate XML formatting
         */
        if (cell.type === ICellType.equation) {
          // TODO: Currently restricting to functions that return numbers only
          rowContent += `
              <c r="${cellPosition}" t="n" s="${styleIndex}">
                <f aca="false">${cell.value.getEquation()}</f>
              </c>\n`;

        } else if (cell.type === ICellType.date) {
          /**
           * Handle date cells - convert to Excel serial number format
           * Excel stores dates as numbers (days since 1900-01-01)
           */
          const excelDateValue = dateToExcelSerial(cell.value);
          rowContent += `
              <c r="${cellPosition}" t="n" s="${styleIndex}">
                <v>${excelDateValue}</v>
              </c>\n`;

        } else {
          /**
           * Handle regular value cells (string, number)
           * Uses <v> tag for value content and appropriate cell type
           */
          rowContent += `
              <c r="${cellPosition}" t="${cellType}" s="${styleIndex}">
                <v>${cellValue}</v>
              </c>\n`;
        }
      }
    });

    // Close the row tag
    rowContent += '</row>\n';
    return rowContent;
  }).join('')
    }
</sheetData>

  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
`};

/**
 * Exports the main worksheet XML generation function
 * @name generateSheetXml
 * @function
 * @description Generates complete sheet.xml content for Excel worksheet with cell data and styling support
 * @see {@link generateSheetXml} - Main worksheet generation function
 */
export { generateSheetXml };
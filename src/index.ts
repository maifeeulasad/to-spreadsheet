/**
 * @fileoverview Main entry point for the to-spreadsheet library
 * This library provides functionality to generate Excel spreadsheets (.xlsx files) 
 * in both Node.js and browser environments with support for cell borders, equations, and various data types.
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { generateExcel, EnvironmentType } from "./generate-excel";
import { SkipCell, skipCell, Equation, writeEquation, createBorder, createAllBorders, createTopBorder, createBottomBorder, createLeftBorder, createRightBorder, createStyledCell, createBorderedCell } from "./util";

/**
 * Enum representing different cell types in Excel
 * @enum {string}
 */
enum ICellType {
  /** String cell type - contains text values */
  string = "s",
  /** Number cell type - contains numeric values */
  number = "n",
  /** Skip cell type - represents empty/skipped cells */
  skip = "skip",
  /** Equation cell type - contains Excel formulas */
  equation = "equation",
}

/**
 * Enum representing different border styles available in Excel
 * @enum {string}
 */
enum BorderStyle {
  /** No border */
  none = "none",
  /** Thin border line (default) */
  thin = "thin",
  /** Medium thickness border line */
  medium = "medium",
  /** Thick border line */
  thick = "thick",
  /** Double border line */
  double = "double",
  /** Dotted border line */
  dotted = "dotted",
  /** Dashed border line */
  dashed = "dashed",
}

/**
 * Interface representing border configuration for a cell
 * @interface IBorder
 */
interface IBorder {
  /** Top border style */
  top?: BorderStyle;
  /** Right border style */
  right?: BorderStyle;
  /** Bottom border style */
  bottom?: BorderStyle;
  /** Left border style */
  left?: BorderStyle;
  /** Border color in hex format (e.g., "#000000" for black) */
  color?: string;
}

/**
 * Interface representing styling options for a cell
 * @interface ICellStyle
 */
interface ICellStyle {
  /** Border configuration for the cell */
  border?: IBorder;
}

/**
 * Interface representing a string cell with optional styling
 * @interface ICellString
 */
interface ICellString {
  /** Cell type identifier */
  type: ICellType.string;
  /** Index reference to the shared strings table */
  value: number;
  /** Optional styling configuration */
  style?: ICellStyle;
}

/**
 * Interface representing a number cell with optional styling
 * @interface ICellNumber
 */
interface ICellNumber {
  /** Cell type identifier */
  type: ICellType.number;
  /** Numeric value of the cell */
  value: number;
  /** Optional styling configuration */
  style?: ICellStyle;
}

/**
 * Interface representing a skipped/empty cell
 * @interface ICellSkip
 */
interface ICellSkip {
  /** Cell type identifier */
  type: ICellType.skip;
}

/**
 * Interface representing a cell containing an Excel formula
 * @interface ICellEquation
 */
interface ICellEquation {
  /** Cell type identifier */
  type: ICellType.equation;
  /** Equation object containing the formula */
  value: Equation;
  /** Optional styling configuration */
  style?: ICellStyle;
}

/**
 * Union type representing any valid cell type
 * @typedef {ICellString | ICellNumber | ICellSkip | ICellEquation} ICell
 */
type ICell = ICellString | ICellNumber | ICellSkip | ICellEquation;

/**
 * Interface representing a row of cells in a worksheet
 * @interface IRows
 */
interface IRows {
  /** Array of cells in this row */
  cells: ICell[];
}

/**
 * Interface representing a worksheet/sheet in the workbook
 * @interface ISheet
 */
interface ISheet {
  /** Name/title of the sheet */
  title: string;
  /** Array of rows containing cell data */
  rows: IRows[];
}

/**
 * Interface representing the complete Excel workbook structure
 * @interface IWorkbook
 */
interface IWorkbook {
  /** Output filename for the Excel file */
  filename: string;
  /** Array of worksheets in the workbook */
  sheets: ISheet[];
  /** Shared strings table for text content optimization */
  strings: string[];
}

/**
 * Interface representing input data structure for generating Excel files
 * Provides a simplified way to define worksheet content
 * @interface IPage
 */
interface IPage {
  /** Title of the worksheet */
  title: string;
  /** 2D array representing rows and columns of data
   * Supports: strings, numbers, undefined (empty cells), SkipCell objects, 
   * Equation objects, and pre-styled ICell objects */
  content: (string | number | undefined | SkipCell | Equation | ICell)[][];
}

/**
 * Export all type definitions and interfaces for external use
 */
export { ICell, ISheet, IWorkbook, IRows, ICellType, IPage, BorderStyle, IBorder, ICellStyle }

/**
 * Sample data demonstrating various features of the library
 * Including basic data types, equations, skip cells, and border styling
 * @constant {IPage[]}
 */
const sampleData = [
  {
    title: 'Maifee1', content: [
      ['meaw', 'grrr'],
      ['woof', 'smack'],
      [1],
      [1, 2],
      [1, 2, 3, writeEquation('SUM(A5,C5)')],
    ]
  },
  { title: 'Maifee2', content: [[1], [1, skipCell(3), 2]] },
  { title: 'Maifee3', content: [['meaw', undefined, "meaw"], ["woof", 'woof']] },
  // Demonstration of border functionality
  { 
    title: 'BorderDemo', 
    content: [
      [
        // Header row with thick borders
        createBorderedCell('Product', createAllBorders(BorderStyle.thick, '#000000')),
        createBorderedCell('Price', createAllBorders(BorderStyle.thick, '#000000')),
        createBorderedCell('Total', createAllBorders(BorderStyle.thick, '#000000'))
      ],
      [
        // Data rows with various border styles
        createBorderedCell('Apple', createLeftBorder()),
        createBorderedCell(10, createTopBorder()),
        createBorderedCell(100, createRightBorder())
      ],
      [
        // Custom styling example
        createStyledCell('Custom', { 
          border: { 
            left: BorderStyle.double, 
            bottom: BorderStyle.thin,
            color: '#0000FF'
          } 
        }),
        200,
        'No Border Cell'
      ]
    ] 
  }
]

// generateExcel(sampleData) // <- check before releasing with `yarn test:compile`

/**
 * Export all public functions and utilities for external use
 * This includes the main generation function, sample data, environment types,
 * utility functions, and all border/styling helper functions
 */
export { generateExcel, sampleData, EnvironmentType, skipCell, writeEquation, createBorder, createAllBorders, createTopBorder, createBottomBorder, createLeftBorder, createRightBorder, createStyledCell, createBorderedCell };

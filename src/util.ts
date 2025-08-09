/**
 * @fileoverview Utility functions for Excel spreadsheet generation
 * Contains helper functions for cell positioning, border styling, equations, and data conversion
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { IRows, IBorder, BorderStyle, ICellType, ICellStyle, ICell, HorizontalAlignment, VerticalAlignment } from ".";

/**
 * Converts zero-based index to one-based index (Excel format)
 * @param {number} index - Zero-based index
 * @returns {number} One-based index
 * @example indexToVbIndex(0) // returns 1
 */
const indexToVbIndex = (index: number) => index + 1;

/**
 * Converts index to relation index with offset
 * Used for internal Excel relationship indexing
 * @param {number} index - Zero-based index
 * @returns {number} Relation index (index + 3)
 */
const indexToVbRelationIndex = (index: number) => indexToVbIndex(index) + 2;

/**
 * Converts numeric index to Excel column notation
 * Handles Excel's base-26 column naming system (A, B, C... Z, AA, AB...)
 * @param {number} index - One-based column index
 * @returns {string} Excel column notation
 * @example 
 * indexToRowIndex(1) // returns "A"
 * indexToRowIndex(26) // returns "Z"  
 * indexToRowIndex(27) // returns "AA"
 */
const indexToRowIndex = (index: number): string => {
    const base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
    let result = "";
    do {
        const remainder = index % 26;
        result = base[(remainder || 26) - 1] + result;
        index = Math.floor(index / 26);
    } while (index > 0);
    return result;
}

/**
 * Converts row and column indices to Excel cell position notation
 * @param {number} row - Zero-based row index
 * @param {number} col - Zero-based column index  
 * @returns {string} Excel cell position (e.g., "A1", "B2")
 * @example rowColumnToVbPosition(0, 0) // returns "A1"
 */
const rowColumnToVbPosition = (row: number, col: number): string => indexToRowIndex(indexToVbIndex(row)) + indexToVbIndex(col);

/**
 * Calculates the extent/range of data in Excel format
 * Used to determine the worksheet dimension attribute
 * @param {IRows[]} rows - Array of row data
 * @returns {string} Excel range notation for the data extent
 * @example calculateExtant(rows) // returns "C3" for 3x3 data
 */
const calculateExtant = (rows: IRows[]): string => rowColumnToVbPosition(
    Math.max(...rows.map(row => row.cells.length)) - 1,
    rows.length - 1
)

/**
 * Class representing a skip cell instruction
 * Used to skip a specified number of cells in a row
 * @class SkipCell
 */
class SkipCell {
    private skipCell: number;
    
    /**
     * Gets the number of cells to skip
     * @returns {number} Number of cells to skip
     */
    public getSkipCell = () => this.skipCell;
    
    /**
     * Creates a SkipCell instance
     * @param {number} skipCell - Number of cells to skip
     */
    constructor(skipCell: number) {
        this.skipCell = skipCell;
    }
}

/**
 * Factory function to create a SkipCell instance
 * @param {number} skipCell - Number of cells to skip
 * @returns {SkipCell} SkipCell instance
 * @example skipCell(3) // skips 3 cells in the row
 */
const skipCell = (skipCell: number) => new SkipCell(skipCell);

/**
 * Class representing an Excel equation/formula
 * @class Equation
 */
class Equation {
    private equation: string;
    
    /**
     * Gets the equation string
     * @returns {string} The Excel formula
     */
    public getEquation = () => this.equation;
    
    /**
     * Creates an Equation instance
     * @param {string} equation - Excel formula string
     */
    constructor(equation: string) {
        this.equation = equation;
    }
}

/**
 * Factory function to create an Equation instance
 * @param {string} equation - Excel formula string (e.g., "SUM(A1:A10)")
 * @returns {Equation} Equation instance
 * @example writeEquation('SUM(A1:B1)') // creates a sum formula
 */
const writeEquation = (equation: string) => new Equation(equation);

/**
 * Helper function that simply returns the provided border configuration
 * Useful for type validation and consistency
 * @param {IBorder} border - Border configuration object
 * @returns {IBorder} The same border configuration
 */
const createBorder = (border: IBorder): IBorder => border;

/**
 * Creates a border configuration with the same style on all sides
 * @param {BorderStyle} style - Border style to apply (defaults to thin)
 * @param {string} color - Border color in hex format (defaults to black)
 * @returns {IBorder} Border configuration with all sides styled
 * @example createAllBorders(BorderStyle.thick, '#FF0000') // thick red borders on all sides
 */
const createAllBorders = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  top: style,
  right: style,
  bottom: style,
  left: style,
  color
});

/**
 * Creates a border configuration with style only on the top side
 * @param {BorderStyle} style - Border style to apply (defaults to thin)
 * @param {string} color - Border color in hex format (defaults to black)
 * @returns {IBorder} Border configuration with top border only
 * @example createTopBorder(BorderStyle.double) // double line top border
 */
const createTopBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  top: style,
  color
});

/**
 * Creates a border configuration with style only on the bottom side
 * @param {BorderStyle} style - Border style to apply (defaults to thin)
 * @param {string} color - Border color in hex format (defaults to black)
 * @returns {IBorder} Border configuration with bottom border only
 * @example createBottomBorder(BorderStyle.dashed, '#0000FF') // blue dashed bottom border
 */
const createBottomBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  bottom: style,
  color
});

/**
 * Creates a border configuration with style only on the left side
 * @param {BorderStyle} style - Border style to apply (defaults to thin)
 * @param {string} color - Border color in hex format (defaults to black)
 * @returns {IBorder} Border configuration with left border only
 * @example createLeftBorder(BorderStyle.medium) // medium line left border
 */
const createLeftBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  left: style,
  color
});

/**
 * Creates a border configuration with style only on the right side
 * @param {BorderStyle} style - Border style to apply (defaults to thin)
 * @param {string} color - Border color in hex format (defaults to black)
 * @returns {IBorder} Border configuration with right border only
 * @example createRightBorder(BorderStyle.dotted) // dotted right border
 */
const createRightBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  right: style,
  color
});

/**
 * Generates a unique string key for a border configuration
 * Used internally for style mapping and deduplication
 * @param {IBorder} border - Border configuration object (optional)
 * @returns {string} Unique key representing the border configuration
 * @internal
 */
const getBorderKey = (border?: IBorder): string => {
  if (!border) return "none";
  
  const parts = [
    border.top || "none",
    border.right || "none", 
    border.bottom || "none",
    border.left || "none",
    border.color || "#000000"
  ];
  
  return parts.join("-");
};

/**
 * Generates a unique key string for complete cell style configuration
 * Used internally to identify unique style combinations for Excel styling
 * @param {ICellStyle} style - Complete style configuration object
 * @returns {string} Unique key representing the complete style configuration
 * @internal
 */
const getStyleKey = (style?: ICellStyle): string => {
  if (!style) return "default";
  
  const parts = [
    getBorderKey(style.border),
    style.backgroundColor || "no-bg",
    style.foregroundColor || "no-fg",
    style.horizontalAlignment || "no-halign",
    style.verticalAlignment || "no-valign"
  ];
  
  return parts.join("|");
};

/**
 * Creates a styled cell with custom styling options
 * @param {string | number} value - Cell value (string or number)
 * @param {ICellStyle} style - Optional styling configuration
 * @returns {ICell} Styled cell object
 * @example createStyledCell('Hello', { border: createAllBorders() })
 */
const createStyledCell = (value: string | number, style?: ICellStyle): ICell => {
  if (typeof value === 'string') {
    return {
      type: ICellType.string,
      value: value as any, // Will be converted to index later
      style
    } as any;
  } else {
    return {
      type: ICellType.number,
      value,
      style
    };
  }
};

/**
 * Creates a cell with border styling
 * Convenience function that combines cell creation with border styling
 * @param {string | number} value - Cell value (string or number)
 * @param {IBorder} border - Border configuration
 * @returns {ICell} Cell with border styling applied
 * @example createBorderedCell('Header', createAllBorders(BorderStyle.thick))
 */
const createBorderedCell = (
  value: string | number, 
  border: IBorder
): ICell => {
  return createStyledCell(value, { border });
};

/**
 * Converts JavaScript Date object to Excel date serial number
 * Excel uses the number of days since January 1, 1900 as its date format
 * Note: Excel incorrectly treats 1900 as a leap year, so we account for that
 * @param {Date} date - JavaScript Date object
 * @returns {number} Excel serial date number
 * @internal
 */
const dateToExcelSerial = (date: Date): number => {
  const excelEpoch = new Date(1900, 0, 1); // January 1, 1900
  const diffTime = date.getTime() - excelEpoch.getTime();
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  // Excel incorrectly treats 1900 as a leap year, so dates after Feb 28, 1900 need +1
  return diffDays + (date >= new Date(1900, 1, 29) ? 2 : 1);
};

/**
 * Creates a date cell with optional styling
 * Converts JavaScript Date to Excel-compatible date format
 * @param {Date} date - JavaScript Date object
 * @param {ICellStyle} style - Optional styling configuration
 * @returns {ICell} Date cell object
 * @example createDateCell(new Date(), { border: createAllBorders() })
 */
const createDateCell = (date: Date, style?: ICellStyle): ICell => {
  return {
    type: ICellType.date,
    value: date,
    style
  } as any;
};

/**
 * Creates a date cell with border styling
 * Convenience function that combines date cell creation with border styling
 * @param {Date} date - JavaScript Date object
 * @param {IBorder} border - Border configuration
 * @returns {ICell} Date cell with border styling applied
 * @example createBorderedDateCell(new Date(), createAllBorders(BorderStyle.thin))
 */
const createBorderedDateCell = (
  date: Date, 
  border: IBorder
): ICell => {
  return createDateCell(date, { border });
};

/**
 * Creates a cell with background color styling
 * Convenience function for applying background color to cells
 * @param {string | number} value - Cell value (string or number)
 * @param {string} backgroundColor - Background color in hex format (e.g., "#FFFF00")
 * @returns {ICell} Cell with background color styling applied
 * @example createBackgroundCell('Highlighted', '#FFFF00')
 */
const createBackgroundCell = (
  value: string | number,
  backgroundColor: string
): ICell => {
  return createStyledCell(value, { backgroundColor });
};

/**
 * Creates a cell with foreground (text) color styling
 * Convenience function for applying text color to cells
 * @param {string | number} value - Cell value (string or number)
 * @param {string} foregroundColor - Text color in hex format (e.g., "#FF0000")
 * @returns {ICell} Cell with text color styling applied
 * @example createForegroundCell('Red Text', '#FF0000')
 */
const createForegroundCell = (
  value: string | number,
  foregroundColor: string
): ICell => {
  return createStyledCell(value, { foregroundColor });
};

/**
 * Creates a cell with both background and foreground color styling
 * Convenience function for applying both background and text colors
 * @param {string | number} value - Cell value (string or number)
 * @param {string} backgroundColor - Background color in hex format
 * @param {string} foregroundColor - Text color in hex format
 * @returns {ICell} Cell with color styling applied
 * @example createColoredCell('Styled Text', '#FFFF00', '#FF0000')
 */
const createColoredCell = (
  value: string | number,
  backgroundColor: string,
  foregroundColor: string
): ICell => {
  return createStyledCell(value, { backgroundColor, foregroundColor });
};

/**
 * Creates a date cell with background color styling
 * Convenience function for applying background color to date cells
 * @param {Date} date - JavaScript Date object
 * @param {string} backgroundColor - Background color in hex format
 * @returns {ICell} Date cell with background color styling applied
 * @example createBackgroundDateCell(new Date(), '#FFFF00')
 */
const createBackgroundDateCell = (
  date: Date,
  backgroundColor: string
): ICell => {
  return createDateCell(date, { backgroundColor });
};

/**
 * Creates a cell with horizontal alignment
 * @param {string | number | Date} value - Cell value
 * @param {HorizontalAlignment} alignment - Horizontal alignment option
 * @returns {ICell} Cell with horizontal alignment styling
 * @example createHorizontallyAlignedCell('Center Me', HorizontalAlignment.center)
 */
const createHorizontallyAlignedCell = (
  value: string | number | Date,
  alignment: HorizontalAlignment
): ICell => {
  if (value instanceof Date) {
    return createDateCell(value, { horizontalAlignment: alignment });
  }
  return createStyledCell(value, { horizontalAlignment: alignment });
};

/**
 * Creates a cell with vertical alignment
 * @param {string | number | Date} value - Cell value
 * @param {VerticalAlignment} alignment - Vertical alignment option
 * @returns {ICell} Cell with vertical alignment styling
 * @example createVerticallyAlignedCell('Top Align', VerticalAlignment.top)
 */
const createVerticallyAlignedCell = (
  value: string | number | Date,
  alignment: VerticalAlignment
): ICell => {
  if (value instanceof Date) {
    return createDateCell(value, { verticalAlignment: alignment });
  }
  return createStyledCell(value, { verticalAlignment: alignment });
};

/**
 * Creates a cell with both horizontal and vertical alignment
 * @param {string | number | Date} value - Cell value
 * @param {HorizontalAlignment} horizontal - Horizontal alignment option
 * @param {VerticalAlignment} vertical - Vertical alignment option
 * @returns {ICell} Cell with both alignment styles
 * @example createAlignedCell('Center Both', HorizontalAlignment.center, VerticalAlignment.center)
 */
const createAlignedCell = (
  value: string | number | Date,
  horizontal: HorizontalAlignment,
  vertical: VerticalAlignment
): ICell => {
  if (value instanceof Date) {
    return createDateCell(value, { 
      horizontalAlignment: horizontal,
      verticalAlignment: vertical 
    });
  }
  return createStyledCell(value, { 
    horizontalAlignment: horizontal,
    verticalAlignment: vertical 
  });
};

/**
 * Creates a center-aligned cell (both horizontally and vertically)
 * Convenience function for the most common alignment case
 * @param {string | number | Date} value - Cell value
 * @returns {ICell} Cell centered both horizontally and vertically
 * @example createCenteredCell('Centered Text')
 */
const createCenteredCell = (
  value: string | number | Date
): ICell => {
  return createAlignedCell(value, HorizontalAlignment.center, VerticalAlignment.center);
};

/**
 * Export all utility functions and classes for external use
 * Includes positioning utilities, cell creation helpers, border functions, and internal utilities
 */
export { 
  indexToVbIndex, 
  indexToVbRelationIndex, 
  indexToRowIndex, 
  rowColumnToVbPosition, 
  calculateExtant, 
  SkipCell, 
  skipCell, 
  Equation, 
  writeEquation,
  createBorder,
  createAllBorders,
  createTopBorder,
  createBottomBorder,
  createLeftBorder,
  createRightBorder,
  getBorderKey,
  getStyleKey,
  createStyledCell,
  createBorderedCell,
  dateToExcelSerial,
  createDateCell,
  createBorderedDateCell,
  createBackgroundCell,
  createForegroundCell,
  createColoredCell,
  createBackgroundDateCell,
  createHorizontallyAlignedCell,
  createVerticallyAlignedCell,
  createAlignedCell,
  createCenteredCell
}
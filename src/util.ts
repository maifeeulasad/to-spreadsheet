import { IRows, IBorder, BorderStyle, ICellType, ICellStyle, ICell } from ".";

const indexToVbIndex = (index: number) => index + 1;

const indexToVbRelationIndex = (index: number) => indexToVbIndex(index) + 2;

// 1->a, 2->b, 26->z, 27->aa
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

const rowColumnToVbPosition = (row: number, col: number): string => indexToRowIndex(indexToVbIndex(row)) + indexToVbIndex(col);

const calculateExtant = (rows: IRows[]): string => rowColumnToVbPosition(
    Math.max(...rows.map(row => row.cells.length)) - 1,
    rows.length - 1
)

class SkipCell {
    private skipCell: number;
    public getSkipCell = () => this.skipCell;
    constructor(skipCell: number) {
        this.skipCell = skipCell;
    }
}

const skipCell = (skipCell: number) => new SkipCell(skipCell);

class Equation {
    private equation: string;
    public getEquation = () => this.equation;
    constructor(equation: string) {
        this.equation = equation;
    }
}

const writeEquation = (equation: string) => new Equation(equation);

// Helper functions for creating border styles
const createBorder = (border: IBorder): IBorder => border;

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

const createTopBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  top: style,
  color
});

const createBottomBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  bottom: style,
  color
});

const createLeftBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  left: style,
  color
});

const createRightBorder = (
  style: BorderStyle = BorderStyle.thin, 
  color: string = "#000000"
): IBorder => ({
  right: style,
  color
});

// Generate a unique string key for a border configuration to use in style mapping
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

// Helper functions to create styled cells
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

const createBorderedCell = (
  value: string | number, 
  border: IBorder
): ICell => {
  return createStyledCell(value, { border });
};

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
  createStyledCell,
  createBorderedCell
}
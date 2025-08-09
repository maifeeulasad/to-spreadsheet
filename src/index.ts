import { generateExcel, EnvironmentType } from "./generate-excel";
import { SkipCell, skipCell, Equation, writeEquation, createBorder, createAllBorders, createTopBorder, createBottomBorder, createLeftBorder, createRightBorder, createStyledCell, createBorderedCell } from "./util";

enum ICellType {
  string = "s",
  number = "n",
  skip = "skip",
  equation = "equation",
}

enum BorderStyle {
  none = "none",
  thin = "thin",
  medium = "medium",
  thick = "thick",
  double = "double",
  dotted = "dotted",
  dashed = "dashed",
}

interface IBorder {
  top?: BorderStyle;
  right?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  color?: string; // hex color like "#000000"
}

interface ICellStyle {
  border?: IBorder;
}

interface ICellString {
  type: ICellType.string;
  value: number;
  style?: ICellStyle;
}
interface ICellNumber {
  type: ICellType.number;
  value: number;
  style?: ICellStyle;
}
interface ICellSkip {
  type: ICellType.skip;
}
interface ICellEquation {
  type: ICellType.equation;
  value: Equation;
  style?: ICellStyle;
}

type ICell = ICellString | ICellNumber | ICellSkip | ICellEquation;

interface IRows {
  cells: ICell[];
}

interface ISheet {
  title: string;
  rows: IRows[];
}

interface IWorkbook {
  filename: string;
  sheets: ISheet[];
  strings: string[];
}

interface IPage {
  title: string
  content: (string | number | undefined | SkipCell | Equation | ICell)[][]
}

export { ICell, ISheet, IWorkbook, IRows, ICellType, IPage, BorderStyle, IBorder, ICellStyle }


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

export { generateExcel, sampleData, EnvironmentType, skipCell, writeEquation, createBorder, createAllBorders, createTopBorder, createBottomBorder, createLeftBorder, createRightBorder, createStyledCell, createBorderedCell };

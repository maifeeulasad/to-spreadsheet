import { generateExcel, EnvironmentType } from "./generate-excel";
import { SkipCell, skipCell, Equation, writeEquation } from "./util";

enum ICellType {
  string = "s",
  number = "n",
  skip = "skip",
  equation = "equation",
}

interface ICellString {
  type: ICellType.string;
  value: number;
}
interface ICellNumber {
  type: ICellType.number;
  value: number;
}
interface ICellSkip {
  type: ICellType.skip;
  value: undefined;
}
interface ICellEquation {
  type: ICellType.equation;
  value: Equation;
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
  content: (string | number | undefined | SkipCell | Equation)[][]
}

export { ICell, ISheet, IWorkbook, IRows, ICellType, IPage }


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
  { title: 'Maifee3', content: [['meaw', undefined, "meaw"], ["woof", 'woof']] }
]

// generateExcel(sampleData) // <- check before releasing with `yarn test:compile`

export { generateExcel, sampleData, EnvironmentType, skipCell, writeEquation };

import { generateExcel, EnvironmentType } from "./generate-excel";

enum ICellType {
  string = "s",
  number = "n",
  skip = "skip",
}

interface ICell {
  type: ICellType;
  value: number | undefined;
}

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
  content: (string | number | undefined)[][]
}

export { ICell, ISheet, IWorkbook, IRows, ICellType, IPage }


const sampleData = [
  {
    title: 'Maifee1', content: [
      ['meaw', 'grrr'],
      ['woof', 'smack'],
      [1],
      [1, 2],
      [1, 2, 3],
    ]
  },
  { title: 'Maifee2', content: [[1], [1, 2]] },
  { title: 'Maifee3', content: [['meaw', undefined, "meaw"], ["woof", 'woof']] }
]

// generateExcel(sampleData) // <- check before releasing with `yarn test:compile`

export { generateExcel, sampleData, EnvironmentType };

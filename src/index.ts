import { generateExcel, generateExcelWorkbook } from "./generate-excel";

enum ICellType {
  string = "s",
  number = "n"
}

interface ICell {
  type: ICellType;
  value: any;
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
  content: (string | number)[][]
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
  { title: 'Maifee3', content: [['meaw', "meaw"], ["woof", 'woof']] }
]

// generateExcel(sampleData);

export { generateExcel, generateExcelWorkbook, sampleData };

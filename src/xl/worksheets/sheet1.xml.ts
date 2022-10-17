interface ICellEntry {
  readonly?: boolean;
  value: number | string;
}

// 1->a, 2->b, 26->z, 27->aa
const index = (value: number) => {
  const base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
  value++;
  let result = "";
  do {
    const remainder = value % 26;
    result = base[(remainder || 26) - 1] + result;
    value = Math.floor(value / 26);
  } while (value > 0);
  return result;
};

const row = (row: number, cellValues: ICellEntry[]) =>
  `<row r="${row}" spans="1:${cellValues.length}">${cellValues
    .map((item: ICellEntry, i: number) => cell(index(i) + row, item))
    .join("")}</row>\n`;

const cell = (row: string, value: ICellEntry) =>
  `    <c r="${row}">${cellValue(value)}</c>\n`;

const cellValue = (value: ICellEntry) => `<v>${value.value}</v>`;

const generateSheet1Xml = (grid: ICellEntry[][]) =>
  `<sheetData>${grid
    .map((ro: ICellEntry[], index: number) => row(index, ro))
    .join("")}</sheetData>`;

export { ICellEntry, generateSheet1Xml };

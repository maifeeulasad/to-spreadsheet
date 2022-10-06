const ok = () => console.log("ok");

ok();


let data = [
    [
      { readOnly: true, value: "" },
      { value: "A", readOnly: true },
      { value: "B", readOnly: true },
      { value: "C", readOnly: true },
      { value: "D", readOnly: true },
      { value: "D", readOnly: true }
    ],
    [
      { readOnly: true, value: 1 },
      { value: 1 },
      { value: 3 },
      { value: 3 },
      { value: 3 }
    ],
    [
      { readOnly: true, value: 2 },
      { value: 2 },
      { value: 4 },
      { value: 4 },
      { value: 4 }
    ],
    [
      { readOnly: true, value: 3 },
      { value: 1 },
      { value: 3 },
      { value: 3 },
      { value: 3 }
    ],
    [
      { readOnly: true, value: 4 },
      { value: 2 },
      { value: 4 },
      { value: 4 },
      { value: 4 }
    ]
]

// 1->a, 2->b, 26->z, 27->aa
const index = (value: number) => {
    const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    value++;
    let result = "";
    do {
        const remainder = value % 26;
        result = base[(remainder || 26) - 1] + result;
        value = Math.floor(value / 26);
    } while (value > 0);
    return result;
};

const row = (row: any, cellValues: any) => `<row r="${row}" spans="1:${cellValues.length}">${cellValues.map((item:any, i:number) => cell(index(i)+row,item)).join('')}</row>\n`

const cell = (row:any, value:any) => `\t<c r="${row}">${cellValue(value)}</c>\n`

const cellValue = (value: any) => `<v>${value.value}</v>`


const sheetData = (grid:any) => `<sheetData>${grid.map((ro:any, i:number) => row(i, ro)).join('')}</sheetData>`


const sheet = (grid:any) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{6A0CD924-A203-41BB-8331-42E46B7D5F20}">
	<dimension ref="B3:I7" />
	<sheetViews>
		<sheetView tabSelected="1" workbookViewId="0">
			<selection activeCell="B3" sqref="B3:D4" />
		</sheetView>
	</sheetViews>
	<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25" />
    ${sheetData(grid)}
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
</worksheet>
`

console.log(sheet(data))
//console.log(index(30))
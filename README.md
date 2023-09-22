# to-spreadsheet
npm to package to create spreadsheet in node environment


# NPM
```
npm i to-spreadsheet
```
https://www.npmjs.com/package/to-spreadsheet


# Usage
```ts
import { generateExcel , generateExcelWorkbook } from 'to-spreadsheet/lib/index';
// import { generateExcel , generateExcelWorkbook } from 'to-spreadsheet/lib/index.js'; // <-- if your compiler gives your some import error message, import this instead

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

generateExcel(sampleData);

// or you can directly call it with workbook data-structure
generateExcelWorkbook(sampleWorkbookData)
```

# Package Details
 - package size: `9.2 kB`
 - unpacked size: `38.0 kB`
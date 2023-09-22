# to-spreadsheet
npm to package to create spreadsheet in node environment


# NPM
```
npm i to-spreadsheet
```
https://www.npmjs.com/package/to-spreadsheet


# Usage
```ts

import { generateExcel , generateExcelWorkbook } from 'to-spreadsheet';

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
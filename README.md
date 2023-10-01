# to-spreadsheet
npm package to create spreadsheet in node environment and in browser


[![npm version](https://img.shields.io/npm/v/to-spreadsheet.svg)](https://www.npmjs.com/package/to-spreadsheet)
[![minified](https://badgen.net/bundlephobia/min/to-spreadsheet)](https://bundlephobia.com/result?p=to-spreadsheet)
[![minified + gzipped](https://badgen.net/bundlephobia/minzip/to-spreadsheet)](https://bundlephobia.com/result?p=to-spreadsheet)

[![GitHub stars](https://img.shields.io/github/stars/maifeeulasad/to-spreadsheet)](https://github.com/maifeeulasad/to-spreadsheet/stargazers)
[![GitHub watchers](https://img.shields.io/github/watchers/maifeeulasad/to-spreadsheet)](https://github.com/maifeeulasad/to-spreadsheet/watchers)

# NPM
```
npm i to-spreadsheet
```


# Usage
[![Open in CodeSandbox](https://img.shields.io/badge/Open%20in-CodeSandbox-blue?logo=codesandbox)](https://codesandbox.io/s/to-spreadsheet-example-hdmrvc?file=/src/App.tsx)

```ts
import { generateExcel , EnvironmentType, skipCell, writeEquation } from 'to-spreadsheet/lib/index';

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

generateExcel(sampleData); // <-- by default generate XLSX for node
generateExcel(sampleData, EnvironmentType.BROWSER); // <-- for browser
```

# Features
 - [x] Multiple sheet support
 - [x] Equations
 - [ ] Cell styling
 - [ ] Sheet styling
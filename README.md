# to-spreadsheet
npm package to create spreadsheet in node environment and in browser


[![npm version](https://img.shields.io/npm/v/to-spreadsheet.svg)](https://www.npmjs.com/package/to-spreadsheet)
[![minified](https://badgen.net/bundlephobia/min/to-spreadsheet)](https://badgen.net/bundlephobia/min/to-spreadsheet)
[![minified + gzipped](https://badgen.net/bundlephobia/minzip/to-spreadsheet)](https://badgen.net/bundlephobia/minzip/to-spreadsheet)

[![GitHub stars](https://img.shields.io/github/stars/maifeeulasad/to-spreadsheet)](https://github.com/maifeeulasad/to-spreadsheet/stargazers)
[![GitHub watchers](https://img.shields.io/github/watchers/maifeeulasad/to-spreadsheet)](https://github.com/maifeeulasad/to-spreadsheet/watchers)
[![Commits after release](https://img.shields.io/github/commits-since/maifeeulasad/to-spreadsheet/latest/main?include_prereleases)](https://img.shields.io/github/commits-since/maifeeulasad/to-spreadsheet/latest/main?include_prereleases)

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
 - [x] Cell borders
 - [ ] Cell styling (fonts, colors, alignment)
 - [ ] Sheet styling

## Cell Borders

You can add borders to cells using the new border functionality:

### Basic Usage

```ts
import { 
  generateExcel, 
  createBorderedCell, 
  createAllBorders, 
  createTopBorder, 
  createBottomBorder, 
  createLeftBorder, 
  createRightBorder,
  BorderStyle 
} from 'to-spreadsheet/lib/index';

const data = [
  {
    title: 'BorderDemo',
    content: [
      [
        // Create cells with all borders
        createBorderedCell('Header 1', createAllBorders(BorderStyle.thick, '#000000')),
        createBorderedCell('Header 2', createAllBorders(BorderStyle.thick, '#000000'))
      ],
      [
        // Create cells with specific borders
        createBorderedCell('Data 1', createTopBorder()),
        createBorderedCell(100, createRightBorder()),
        createBorderedCell('Final', createBottomBorder())
      ]
    ]
  }
];

generateExcel(data);
```

### Border Styles

Available border styles:
- `BorderStyle.none` - No border
- `BorderStyle.thin` - Thin border (default)
- `BorderStyle.medium` - Medium border
- `BorderStyle.thick` - Thick border  
- `BorderStyle.double` - Double border
- `BorderStyle.dotted` - Dotted border
- `BorderStyle.dashed` - Dashed border

### Helper Functions

**Border Creation:**
- `createAllBorders(style?, color?)` - Creates borders on all sides
- `createTopBorder(style?, color?)` - Creates only top border
- `createBottomBorder(style?, color?)` - Creates only bottom border
- `createLeftBorder(style?, color?)` - Creates only left border
- `createRightBorder(style?, color?)` - Creates only right border
- `createBorder(borderConfig)` - Creates custom border configuration

**Cell Creation:**
- `createBorderedCell(value, border)` - Creates a cell with border
- `createStyledCell(value, style)` - Creates a cell with custom styling

### Advanced Usage

```ts
import { createStyledCell, BorderStyle } from 'to-spreadsheet/lib/index';

// Custom border configuration
const customCell = createStyledCell('Custom', {
  border: {
    left: BorderStyle.double,
    top: BorderStyle.thin,
    right: BorderStyle.dashed,
    bottom: BorderStyle.thick,
    color: '#FF0000' // Red borders
  }
});

// Mix styled and regular cells
const data = [
  {
    title: 'Mixed',
    content: [
      [customCell, 'Regular Cell', 42]
    ]
  }
];
```

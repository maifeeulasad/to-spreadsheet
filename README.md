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

# Cell Features

## Dates

You can create date cells that are properly formatted in Excel:

### Basic Date Usage

```ts
import { 
  generateExcel, 
  createDateCell, 
  createBorderedDateCell,
  createBackgroundDateCell 
} from 'to-spreadsheet/lib/index';

const data = [
  {
    title: 'DateDemo',
    content: [
      [
        'Event',
        'Date',
        'Styled Date'
      ],
      [
        'Project Start',
        createDateCell(new Date('2024-01-01')),
        createBorderedDateCell(new Date('2024-01-15'), createAllBorders())
      ],
      [
        'Milestone',
        createDateCell(new Date()),
        createBackgroundDateCell(new Date('2024-12-25'), '#FFCCCC')
      ]
    ]
  }
];
```

### Date Helper Functions

- `createDateCell(date, style?)` - Creates a date cell with optional styling
- `createBorderedDateCell(date, border)` - Creates a date cell with borders
- `createBackgroundDateCell(date, backgroundColor)` - Creates a date cell with background color

## Colors

You can add background and foreground colors to cells:

### Basic Color Usage

```ts
import { 
  generateExcel, 
  createBackgroundCell, 
  createForegroundCell, 
  createColoredCell,
  createStyledCell
} from 'to-spreadsheet/lib/index';

const data = [
  {
    title: 'ColorDemo',
    content: [
      [
        'Feature',
        'Background Color',
        'Foreground Color',
        'Both Colors'
      ],
      [
        'Yellow Background',
        createBackgroundCell('Highlighted', '#FFFF00'),
        createForegroundCell('Red Text', '#FF0000'),
        createColoredCell('Green BG, Red Text', '#00FF00', '#FF0000')
      ],
      [
        'Complex Styling',
        createStyledCell('Full Style', {
          backgroundColor: '#FFFFCC',
          foregroundColor: '#0000FF',
          border: createAllBorders(BorderStyle.thick, '#000000')
        }),
        'Regular cell',
        42
      ]
    ]
  }
];
```

### Color Helper Functions

- `createBackgroundCell(value, backgroundColor)` - Creates cell with background color
- `createForegroundCell(value, foregroundColor)` - Creates cell with text color
- `createColoredCell(value, backgroundColor, foregroundColor)` - Creates cell with both colors
- `createStyledCell(value, style)` - Creates cell with full styling options

# Features
 - [x] Multiple sheet support
 - [x] Equations
 - [x] Cell borders
 - [x] Cell styling (background colors, foreground colors, dates)
 - [x] Date cells with proper Excel formatting
 - [x] Cell alignment (horizontal and vertical)
 - [ ] Sheet styling

## Cell Borders

You can add borders to cells using the border functionality:

### Basic Border Usage

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

### Border Helper Functions

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

### Advanced Border Usage

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

## Combined Styling

All styling features can be combined together:

```ts
import { 
  createStyledCell, 
  createAllBorders, 
  BorderStyle 
} from 'to-spreadsheet/lib/index';

const data = [
  {
    title: 'CombinedDemo',
    content: [
      [
        // Cell with background color, text color, and borders
        createStyledCell('Fully Styled', {
          backgroundColor: '#FFFFCC',    // Light yellow background
          foregroundColor: '#0000FF',    // Blue text
          border: createAllBorders(BorderStyle.thick, '#FF0000') // Red thick border
        }),
        
        // Date with background and border
        createDateCell(new Date(), {
          backgroundColor: '#CCFFCC',    // Light green background
          border: createAllBorders(BorderStyle.double, '#008000') // Green double border
        }),
        
        // Regular cells for comparison
        'Plain text',
        42
      ]
    ]
  }
];
```

### Color Format

Colors should be specified in hex format:
- `#FF0000` - Red
- `#00FF00` - Green  
- `#0000FF` - Blue
- `#FFFF00` - Yellow
- `#FF00FF` - Magenta
- `#00FFFF` - Cyan
- `#000000` - Black
- `#FFFFFF` - White
- `#CCCCCC` - Light gray

## Cell Alignment

You can align cell content both horizontally and vertically:

### Basic Alignment Usage

```ts
import { 
  generateExcel, 
  createHorizontallyAlignedCell, 
  createVerticallyAlignedCell, 
  createAlignedCell,
  createCenteredCell,
  HorizontalAlignment, 
  VerticalAlignment 
} from 'to-spreadsheet/lib/index';

const data = [
  {
    title: 'AlignmentDemo',
    content: [
      [
        'Feature',
        'Horizontal',
        'Vertical',
        'Both'
      ],
      [
        'Left Align',
        createHorizontallyAlignedCell('Left Text', HorizontalAlignment.left),
        createVerticallyAlignedCell('Top Text', VerticalAlignment.top),
        createAlignedCell('Top-Left', HorizontalAlignment.left, VerticalAlignment.top)
      ],
      [
        'Center Align',
        createHorizontallyAlignedCell('Center Text', HorizontalAlignment.center),
        createVerticallyAlignedCell('Center Text', VerticalAlignment.center),
        createCenteredCell('Full Center')
      ],
      [
        'Right Align',
        createHorizontallyAlignedCell('Right Text', HorizontalAlignment.right),
        createVerticallyAlignedCell('Bottom Text', VerticalAlignment.bottom),
        createAlignedCell('Bottom-Right', HorizontalAlignment.right, VerticalAlignment.bottom)
      ]
    ]
  }
];
```

### Horizontal Alignment Options

- `HorizontalAlignment.general` - General alignment (Excel default)
- `HorizontalAlignment.left` - Left alignment
- `HorizontalAlignment.center` - Center alignment  
- `HorizontalAlignment.right` - Right alignment
- `HorizontalAlignment.fill` - Fill alignment
- `HorizontalAlignment.justify` - Justify alignment
- `HorizontalAlignment.centerContinuous` - Center across selection
- `HorizontalAlignment.distributed` - Distributed alignment

### Vertical Alignment Options

- `VerticalAlignment.top` - Top alignment
- `VerticalAlignment.center` - Center alignment
- `VerticalAlignment.bottom` - Bottom alignment
- `VerticalAlignment.justify` - Justify alignment
- `VerticalAlignment.distributed` - Distributed alignment

### Alignment Helper Functions

- `createHorizontallyAlignedCell(value, alignment)` - Creates cell with horizontal alignment
- `createVerticallyAlignedCell(value, alignment)` - Creates cell with vertical alignment
- `createAlignedCell(value, horizontal, vertical)` - Creates cell with both alignments
- `createCenteredCell(value)` - Creates center-aligned cell (convenience function)

### Combined with Other Features

Alignment works seamlessly with all other styling features:

```ts
import { 
  createStyledCell, 
  HorizontalAlignment, 
  VerticalAlignment,
  createAllBorders, 
  BorderStyle 
} from 'to-spreadsheet/lib/index';

const data = [
  {
    title: 'ComplexStyling',
    content: [
      [
        // Full styling with alignment, colors, and borders
        createStyledCell('Complete Style', {
          horizontalAlignment: HorizontalAlignment.center,
          verticalAlignment: VerticalAlignment.center,
          backgroundColor: '#CCFFCC',
          foregroundColor: '#FF0000',
          border: createAllBorders(BorderStyle.thick, '#000000')
        }),
        
        // Aligned date cell
        createDateCell(new Date(), {
          horizontalAlignment: HorizontalAlignment.right,
          verticalAlignment: VerticalAlignment.center,
          backgroundColor: '#FFFFCC'
        }),
        
        // Simple centered text
        createCenteredCell('Centered')
      ]
    ]
  }
];
```

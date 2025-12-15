<p align="center">
  <img src="assets/logo.svg" alt="Cellify Logo" width="80" height="80">
</p>

<h1 align="center">Cellify</h1>

<p align="center">
  <a href="https://github.com/abdullahmujahidali/Cellify/actions/workflows/ci.yml"><img src="https://github.com/abdullahmujahidali/Cellify/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://codecov.io/gh/abdullahmujahidali/Cellify"><img src="https://codecov.io/gh/abdullahmujahidali/Cellify/graph/badge.svg" alt="codecov"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
</p>

<p align="center">
  A lightweight, zero-dependency* spreadsheet library for JavaScript/TypeScript.<br>
  Create, read, and manipulate Excel files with ease.
</p>

> \*Only dependency is `fflate` for ZIP compression

## Features

- **Excel Import/Export** - Full .xlsx support with styles, formulas, merged cells
- **CSV Import/Export** - Auto-detection of delimiters and data types
- **Complete Styling** - Fonts, fills, borders, alignment, number formats
- **Cell Merging** - With overlap detection and validation
- **Formulas** - Store and preserve formulas (evaluation coming soon)
- **Freeze Panes** - Lock rows/columns for scrolling
- **Auto Filters** - Enable filtering on data ranges
- **Type-Safe** - Built with TypeScript from the ground up
- **Universal** - Works in Node.js and browsers
- **Accessible** - Built-in a11y helpers for screen readers

## Demo

Try the interactive demo to test import/export functionality:

```bash
npm run demo
```

![Cellify Demo](docs/images/demo-screenshot.png)

## Installation

```bash
npm install cellify
```

## Quick Start

```typescript
import { Workbook, workbookToXlsx } from 'cellify';

// Create a workbook
const workbook = new Workbook();
const sheet = workbook.addSheet('Sales Data');

// Set values
sheet.cell('A1').value = 'Product';
sheet.cell('B1').value = 'Revenue';
sheet.cell('A2').value = 'Widget';
sheet.cell('B2').value = 1500;

// Apply styles
sheet.cell('A1').style = {
  font: { bold: true, color: '#FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#4472C4' },
};

// Export to Excel
const xlsxBuffer = workbookToXlsx(workbook);
// Save or send xlsxBuffer as .xlsx file
```

## Excel (.xlsx) Export

Create Excel files with full styling support:

```typescript
import { Workbook, workbookToXlsx, workbookToXlsxBlob } from 'cellify';

const workbook = new Workbook();
workbook.title = 'Sales Report';
workbook.author = 'Cellify';

const sheet = workbook.addSheet('Q4 Sales');

// Headers with styling
const headers = ['Product', 'Units', 'Price', 'Total'];
headers.forEach((header, i) => {
  const cell = sheet.cell(0, i);
  cell.value = header;
  cell.style = {
    font: { bold: true, size: 12, color: '#FFFFFF' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#2F5496' },
    alignment: { horizontal: 'center' },
    borders: {
      bottom: { style: 'medium', color: '#000000' },
    },
  };
});

// Data rows
const data = [
  ['Laptop', 150, 999.99],
  ['Mouse', 500, 29.99],
  ['Keyboard', 300, 79.99],
];

data.forEach((row, rowIndex) => {
  row.forEach((value, colIndex) => {
    sheet.cell(rowIndex + 1, colIndex).value = value;
  });
  // Formula for total
  sheet.cell(rowIndex + 1, 3).setFormula(`B${rowIndex + 2}*C${rowIndex + 2}`);
});

// Set column widths
sheet.setColumnWidth(0, 15); // Product
sheet.setColumnWidth(1, 10); // Units
sheet.setColumnWidth(2, 12); // Price
sheet.setColumnWidth(3, 12); // Total

// Freeze header row
sheet.freeze(1, 0);

// Enable auto filter
sheet.setAutoFilter('A1:D4');

// Export
const buffer = workbookToXlsx(workbook); // Uint8Array for Node.js
const blob = workbookToXlsxBlob(workbook); // Blob for browsers
```

## Excel (.xlsx) Import

Read existing Excel files:

```typescript
import { xlsxToWorkbook, xlsxBlobToWorkbook } from 'cellify';

// From Uint8Array (Node.js)
const buffer = fs.readFileSync('report.xlsx');
const { workbook, stats, warnings } = xlsxToWorkbook(new Uint8Array(buffer));

console.log(`Imported ${stats.sheetCount} sheets, ${stats.totalCells} cells`);

// Access data
const sheet = workbook.getSheetByIndex(0);
console.log(sheet.cell('A1').value); // Read cell value
console.log(sheet.cell('B2').formula); // Read formula

// From Blob (browsers)
const file = document.getElementById('fileInput').files[0];
const result = await xlsxBlobToWorkbook(file);

// Import options
const { workbook } = xlsxToWorkbook(buffer, {
  sheets: ['Sheet1', 'Summary'], // Import specific sheets
  importFormulas: true,          // Preserve formulas
  importStyles: true,            // Import cell styles
  importMergedCells: true,       // Import merged cells
  importFreezePanes: true,       // Import freeze panes
  maxRows: 1000,                 // Limit rows (0 = unlimited)
  maxCols: 50,                   // Limit columns (0 = unlimited)
  onProgress: (phase, current, total) => {
    console.log(`${phase}: ${current}/${total}`);
  },
});
```

## CSV Import/Export

```typescript
import { sheetToCsv, csvToWorkbook } from 'cellify';

// Export to CSV
const csv = sheetToCsv(sheet);
// "Product,Units,Price,Total\r\nLaptop,150,999.99,149998.5"

// Import from CSV (auto-detects delimiter)
const { workbook } = csvToWorkbook(csvString);

// With options
const csv = sheetToCsv(sheet, {
  delimiter: ';',
  includeHeaders: true,
  dateFormat: 'YYYY-MM-DD',
});
```

## Cell Styling

```typescript
// Font styling
cell.style = {
  font: {
    name: 'Arial',
    size: 14,
    bold: true,
    italic: true,
    underline: 'single', // 'single' | 'double'
    strikethrough: true,
    color: '#FF0000',
  },
};

// Fill/background
cell.style = {
  fill: {
    type: 'pattern',
    pattern: 'solid',
    foregroundColor: '#FFFF00',
  },
};

// Borders
cell.style = {
  borders: {
    top: { style: 'thin', color: '#000000' },
    right: { style: 'medium', color: '#000000' },
    bottom: { style: 'thick', color: '#000000' },
    left: { style: 'dashed', color: '#000000' },
  },
};

// Alignment
cell.style = {
  alignment: {
    horizontal: 'center', // 'left' | 'center' | 'right'
    vertical: 'middle',   // 'top' | 'middle' | 'bottom'
    wrapText: true,
    textRotation: 45,     // -90 to 90 degrees
  },
};

// Number format
cell.style = {
  numberFormat: {
    formatCode: '#,##0.00', // Excel format codes
  },
};

// Apply style to range
sheet.applyStyle('A1:D1', {
  font: { bold: true },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#E0E0E0' },
});
```

## Cell Merging

```typescript
// Merge cells using A1 notation
sheet.mergeCells('A1:D1'); // Merge A1 through D1

// Merge cells using coordinates
sheet.mergeCells({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 });

// Check if cell is part of a merge
const cell = sheet.cell('B1');
if (cell.isMergeMaster) {
  console.log('This is the top-left cell of a merge');
}
if (cell.isMerged) {
  console.log('This cell is merged into another cell');
}

// Unmerge cells
sheet.unmergeCells('A1:D1');
```

## Row and Column Configuration

```typescript
// Set column width (in characters)
sheet.setColumnWidth(0, 20);  // Column A = 20 characters wide
sheet.setColumnWidth('B', 15); // Column B = 15 characters

// Set row height (in points)
sheet.setRowHeight(0, 30); // Row 1 = 30 points tall

// Hide rows/columns
sheet.hideRow(5);
sheet.hideColumn('C');

// Get all column configs
sheet.columns.forEach((config, index) => {
  console.log(`Column ${index}: width=${config.width}`);
});
```

## Freeze Panes

```typescript
// Freeze first row (header)
sheet.freeze(1, 0);

// Freeze first column
sheet.freeze(0, 1);

// Freeze both (first row and first two columns)
sheet.freeze(1, 2);
```

## Formulas

```typescript
// Set a formula
sheet.cell('D2').setFormula('B2*C2');

// With or without leading '='
sheet.cell('D3').setFormula('=SUM(B2:B10)');

// Access formula text
const formula = sheet.cell('D2').formula;
console.log(formula.formula); // 'B2*C2'

// Note: Cellify stores formulas but doesn't evaluate them yet
// Excel will calculate values when the file is opened
```

## Named Ranges

```typescript
// Define a named range
workbook.addDefinedName('SalesData', 'Sheet1!$A$1:$D$100');

// Use in formulas
sheet.cell('E1').setFormula('SUM(SalesData)');
```

## Workbook Properties

```typescript
const workbook = new Workbook();

// Set metadata
workbook.title = 'Q4 Sales Report';
workbook.author = 'John Doe';
workbook.setProperties({
  subject: 'Quarterly Sales',
  company: 'Acme Corp',
  keywords: ['sales', 'q4', '2024'],
});

// These appear in Excel's File > Info panel
```

## Accessibility

Cellify provides helpers for building accessible spreadsheet UIs:

```typescript
import {
  getCellAccessibility,
  getAriaAttributes,
  announceNavigation,
} from 'cellify';

// Generate ARIA attributes for a cell
const a11y = getCellAccessibility(cell, sheet, {
  headerRows: 1,
  headerCols: 1,
});

const ariaProps = getAriaAttributes(a11y);
// { role: 'gridcell', 'aria-colindex': 2, 'aria-rowindex': 3, ... }

// Screen reader announcement
const announcement = announceNavigation(cell);
// { message: 'Cell B3, row 3, column 2, 1500', priority: 'polite' }
```

## API Reference

### Workbook

| Method | Description |
|--------|-------------|
| `addSheet(name?)` | Add a new sheet |
| `getSheet(name)` | Get sheet by name |
| `getSheetByIndex(index)` | Get sheet by index |
| `removeSheet(sheet)` | Remove a sheet |
| `renameSheet(old, new)` | Rename a sheet |

### Sheet

| Method | Description |
|--------|-------------|
| `cell(address)` | Get cell by A1 notation |
| `cell(row, col)` | Get cell by coordinates |
| `mergeCells(range)` | Merge cells in range |
| `unmergeCells(range)` | Unmerge cells |
| `setColumnWidth(col, width)` | Set column width |
| `setRowHeight(row, height)` | Set row height |
| `freeze(rows, cols)` | Freeze panes |
| `setAutoFilter(range)` | Enable auto filter |
| `applyStyle(range, style)` | Apply style to range |

### Cell

| Property | Description |
|----------|-------------|
| `value` | Cell value (string, number, boolean, Date) |
| `style` | Cell style object |
| `formula` | Formula object |
| `address` | A1 notation address |
| `type` | Value type |
| `isMerged` | Is part of merge |
| `isMergeMaster` | Is top-left of merge |

### Format Functions

| Function | Description |
|----------|-------------|
| `workbookToXlsx(workbook)` | Export to Uint8Array |
| `workbookToXlsxBlob(workbook)` | Export to Blob |
| `xlsxToWorkbook(buffer)` | Import from Uint8Array |
| `xlsxBlobToWorkbook(blob)` | Import from Blob |
| `sheetToCsv(sheet)` | Export sheet to CSV |
| `csvToWorkbook(csv)` | Import CSV to workbook |

## Browser Usage

```html
<script type="module">
import { Workbook, workbookToXlsxBlob } from 'cellify';

// Create workbook
const workbook = new Workbook();
const sheet = workbook.addSheet('Data');
sheet.cell('A1').value = 'Hello, Excel!';

// Download as file
const blob = workbookToXlsxBlob(workbook);
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'spreadsheet.xlsx';
a.click();
</script>
```

## Supported Cell Types

| Type | Example | Notes |
|------|---------|-------|
| String | `'Hello'` | Plain text |
| Number | `42`, `3.14` | Integers and decimals |
| Boolean | `true`, `false` | |
| Date | `new Date()` | Converted to Excel serial |
| Formula | `'=SUM(A1:A10)'` | Stored, evaluated by Excel |
| Error | `'#VALUE!'` | Excel error values |

## Border Styles

Available border styles: `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`, `dashDot`, `dashDotDot`, `mediumDashed`, `mediumDashDot`, `mediumDashDotDot`, `slantDashDot`

## Planned Features

- [ ] Formula evaluation engine
- [ ] Hyperlinks
- [ ] Data validation (dropdowns, constraints)
- [ ] Comments/notes
- [ ] Conditional formatting export
- [ ] Charts (basic)
- [ ] Streaming for large files

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for development setup and guidelines.

## License

MIT - see [LICENSE](./LICENSE)

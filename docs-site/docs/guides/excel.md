---
sidebar_position: 1
---

# Excel Import/Export

Cellify provides comprehensive support for reading and writing Excel (.xlsx) files.

## Exporting to Excel

### Basic Export

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Report');

// Add data
sheet.cell(0, 0).value('Sales Report');
sheet.cell(1, 0).value('Product');
sheet.cell(1, 1).value('Revenue');
sheet.cell(2, 0).value('Widget');
sheet.cell(2, 1).value(15000);

// Export
const blob = workbookToXlsxBlob(workbook);
```

### Export with Styling

```typescript
const workbook = new Workbook();
const sheet = workbook.addSheet('Styled Report');

// Header with styling
sheet.cell(0, 0).value('Quarterly Report').style({
  font: { bold: true, size: 16, color: '#FFFFFF' },
  fill: { color: '#059669' },
  alignment: { horizontal: 'center' }
});

// Merge the header across columns
sheet.mergeCells('A1:D1');

// Column headers
const headers = ['Product', 'Q1', 'Q2', 'Q3'];
headers.forEach((header, col) => {
  sheet.cell(1, col).value(header).style({
    font: { bold: true },
    fill: { color: '#F3F4F6' },
    border: { bottom: { style: 'thin', color: '#000000' } }
  });
});

// Data rows
const data = [
  ['Widget A', 5000, 6000, 7500],
  ['Widget B', 3000, 3500, 4000],
];

data.forEach((row, rowIndex) => {
  row.forEach((value, colIndex) => {
    const cell = sheet.cell(rowIndex + 2, colIndex);
    cell.value(value);
    if (typeof value === 'number') {
      cell.style({ numberFormat: '$#,##0' });
    }
  });
});

const blob = workbookToXlsxBlob(workbook);
```

### Multi-Sheet Workbook

```typescript
const workbook = new Workbook();

// Sales sheet
const salesSheet = workbook.addSheet('Sales');
salesSheet.cell(0, 0).value('Sales Data');

// Inventory sheet
const inventorySheet = workbook.addSheet('Inventory');
inventorySheet.cell(0, 0).value('Inventory Data');

// Summary sheet
const summarySheet = workbook.addSheet('Summary');
summarySheet.cell(0, 0).value('Summary');

const blob = workbookToXlsxBlob(workbook);
```

### Export to Uint8Array (Node.js)

```typescript
import { workbookToXlsx } from 'cellify';
import { writeFileSync } from 'fs';

const workbook = new Workbook();
// ... add data ...

const buffer = workbookToXlsx(workbook);
writeFileSync('output.xlsx', buffer);
```

### Export with Comments

```typescript
const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

// Add data with comments
sheet.cell(0, 0).value('Status').comment('Review status indicator');
sheet.cell(0, 1).value('Sales').comment('Q4 2024 figures');
sheet.cell(1, 0).value('Approved');
sheet.cell(1, 1).value(45000).comment('Includes online and retail');

const blob = workbookToXlsxBlob(workbook);
```

## Importing from Excel

### Basic Import

```typescript
import { xlsxBlobToWorkbook } from 'cellify';

// From file input in browser
const fileInput = document.querySelector('input[type="file"]');
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const result = await xlsxBlobToWorkbook(file);

  const { workbook, stats, warnings } = result;

  console.log('Sheets:', workbook.sheetCount);
  console.log('Total cells:', stats.totalCells);
  console.log('Formulas:', stats.formulaCells);
  console.log('Merged ranges:', stats.mergedRanges);

  // Iterate through sheets
  workbook.sheets.forEach(sheet => {
    console.log(`Sheet: ${sheet.name}`);

    // Get sheet dimensions
    const dims = sheet.dimensions;
    if (dims) {
      for (let r = dims.startRow; r <= dims.endRow; r++) {
        for (let c = dims.startCol; c <= dims.endCol; c++) {
          const cell = sheet.getCell(r, c);
          if (cell?.value !== undefined) {
            console.log(`[${r},${c}]: ${cell.value}`);
          }
        }
      }
    }
  });
});
```

### Import with Options

```typescript
const result = await xlsxBlobToWorkbook(file, {
  // Import specific sheets only
  sheets: ['Sheet1', 'Summary'],

  // Or by index
  sheets: [0, 2],

  // Or all sheets (default)
  sheets: 'all',

  // Control what to import
  importFormulas: true,      // Import formulas (default: true)
  importStyles: true,        // Import cell styles (default: true)
  importMergedCells: true,   // Import merged cells (default: true)
  importDimensions: true,    // Import column widths/row heights (default: true)
  importFreezePanes: true,   // Import freeze panes (default: true)
  importProperties: true,    // Import document properties (default: true)
  importComments: true,      // Import cell comments/notes (default: true)

  // Limit data size
  maxRows: 1000,            // Max rows to import (0 = unlimited)
  maxCols: 50,              // Max columns to import (0 = unlimited)
});
```

### Import from Uint8Array (Node.js)

```typescript
import { xlsxToWorkbook } from 'cellify';
import { readFileSync } from 'fs';

const buffer = readFileSync('data.xlsx');
const result = xlsxToWorkbook(new Uint8Array(buffer));

console.log('Imported:', result.stats.totalCells, 'cells');
```

### Accessing Imported Comments

```typescript
const result = await xlsxBlobToWorkbook(file);
const { workbook } = result;

// Iterate through cells and access comments
workbook.sheets.forEach(sheet => {
  for (const cell of sheet.cells()) {
    if (cell.comment) {
      console.log(`Cell ${cell.address}: "${cell.value}" has comment: "${cell.comment}"`);
    }
  }
});
```

## Import Result Structure

```typescript
interface XlsxImportResult {
  workbook: Workbook;
  stats: {
    sheetCount: number;
    totalCells: number;
    formulaCells: number;
    mergedRanges: number;
    durationMs: number;
  };
  warnings: Array<{
    code: string;
    message: string;
    location?: string;
  }>;
}
```

## Supported Features

| Feature | Export | Import |
|---------|--------|--------|
| Cell values (string, number, boolean, date) | ✅ | ✅ |
| Formulas | ✅ | ✅ |
| Cell styles (font, fill, border, alignment) | ✅ | ✅ |
| Number formats | ✅ | ✅ |
| Merged cells | ✅ | ✅ |
| Column widths | ✅ | ✅ |
| Row heights | ✅ | ✅ |
| Freeze panes | ✅ | ✅ |
| Auto filters | ✅ | ✅ |
| Multiple sheets | ✅ | ✅ |
| Document properties | ✅ | ✅ |
| Cell comments/notes | ✅ | ✅ |
| Hyperlinks | ❌ | ❌ |
| Images/Charts | ❌ | ❌ |
| Conditional formatting | ❌ | ❌ |
| Data validation | ❌ | ❌ |

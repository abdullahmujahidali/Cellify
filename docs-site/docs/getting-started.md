---
sidebar_position: 2
---

# Getting Started

This guide will help you get up and running with Cellify in minutes.

## Installation

Install Cellify using your preferred package manager:

```bash
# npm
npm install cellify

# yarn
yarn add cellify

# pnpm
pnpm add cellify
```

## Requirements

- Node.js 20.19.0+ or 22.12.0+
- Modern browser (Chrome, Firefox, Safari, Edge)

## Basic Usage

### Creating a Workbook

```typescript
import { Workbook } from 'cellify';

// Create a new workbook
const workbook = new Workbook();

// Add a sheet
const sheet = workbook.addSheet('My Sheet');

// Set cell values
sheet.cell(0, 0).value('Hello');
sheet.cell(0, 1).value('World');
sheet.cell(1, 0).value(42);
sheet.cell(1, 1).value(new Date());
```

### Exporting to Excel

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

// Add some data
sheet.cell(0, 0).value('Name');
sheet.cell(0, 1).value('Age');
sheet.cell(1, 0).value('Alice');
sheet.cell(1, 1).value(30);

// Export to blob (for browser download)
const blob = workbookToXlsxBlob(workbook);

// Download in browser
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'data.xlsx';
a.click();
URL.revokeObjectURL(url);
```

### Importing from Excel

```typescript
import { xlsxBlobToWorkbook } from 'cellify';

// From file input
const fileInput = document.querySelector('input[type="file"]');
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const result = await xlsxBlobToWorkbook(file);

  console.log('Sheets:', result.workbook.sheetCount);
  console.log('Total cells:', result.stats.totalCells);

  // Access data
  const sheet = result.workbook.sheets[0];
  const value = sheet.getCell(0, 0)?.value;
});
```

### WASM Performance Boost

For large Excel files (10K+ cells), initialize the WASM parser at startup for 10-50x faster imports:

```typescript
import { initXlsxWasm, xlsxBlobToWorkbook } from 'cellify';

// Initialize once at startup
await initXlsxWasm();

// All imports now use high-performance WASM parser
const result = await xlsxBlobToWorkbook(file);
```

See the [Excel Import/Export guide](/docs/guides/excel#wasm-acceleration) for more details.

### Working with CSV

```typescript
import { csvToWorkbook, sheetToCsv } from 'cellify';

// Import CSV
const csvText = `Name,Age,City
Alice,30,NYC
Bob,25,LA`;

const result = csvToWorkbook(csvText);
const sheet = result.workbook.sheets[0];

// Export to CSV
const csv = sheetToCsv(sheet);
console.log(csv);
```

## What's Next?

Now that you have the basics, explore these guides:

- [Excel Import/Export](/docs/guides/excel) - Advanced Excel operations
- [Styling Cells](/docs/guides/styling) - Fonts, colors, borders
- [Merging Cells](/docs/guides/merging) - Create merged cell ranges
- [Formulas](/docs/guides/formulas) - Work with Excel formulas

Or dive into the API reference:

- [Workbook API](/docs/api/workbook)
- [Sheet API](/docs/api/sheet)
- [Cell API](/docs/api/cell)

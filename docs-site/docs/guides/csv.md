---
sidebar_position: 2
---

# CSV Import/Export

Cellify provides comprehensive support for reading and writing CSV files with automatic type detection.

## Importing CSV

### Basic Import

```typescript
import { csvToWorkbook } from 'cellify';

const csv = `Name,Age,City
Alice,30,New York
Bob,25,Los Angeles`;

const workbook = csvToWorkbook(csv);
const sheet = workbook.sheets[0];

console.log(sheet.getCell(0, 0)?.value); // 'Name'
console.log(sheet.getCell(1, 1)?.value); // 30 (number)
```

### Import with Options

```typescript
import { csvToWorkbook } from 'cellify';

const workbook = csvToWorkbook(csvText, {
  // Delimiter detection (auto-detects if not specified)
  delimiter: ',',

  // Quote character
  quoteChar: '"',

  // Sheet name for imported data
  sheetName: 'Imported Data',

  // Starting cell position
  startCell: 'A1',

  // First row contains headers
  hasHeaders: true,

  // Skip empty lines
  skipEmptyLines: true,

  // Trim whitespace from values
  trimValues: true,

  // Auto-detect numbers
  detectNumbers: true,

  // Auto-detect dates
  detectDates: true,

  // Date formats to try (in order)
  dateFormats: ['yyyy-mm-dd', 'mm/dd/yyyy', 'dd/mm/yyyy'],

  // Limit rows imported (0 = unlimited)
  maxRows: 1000,

  // Progress callback
  onProgress: (current, total) => {
    console.log(`Processing row ${current} of ${total}`);
  },
});
```

### Import into Existing Sheet

```typescript
import { csvToSheet } from 'cellify';

const sheet = workbook.addSheet('CSV Data');
const result = csvToSheet(csvText, sheet, {
  startCell: 'B2',
  hasHeaders: true,
});

console.log('Rows imported:', result.rowCount);
console.log('Columns:', result.columnCount);
console.log('Headers:', result.headers);
```

### Import from Buffer (Node.js)

```typescript
import { csvBufferToWorkbook } from 'cellify';
import { readFileSync } from 'fs';

const buffer = readFileSync('data.csv');
const workbook = csvBufferToWorkbook(new Uint8Array(buffer), {
  delimiter: ';',
  detectNumbers: true,
});
```

## Exporting to CSV

### Basic Export

```typescript
import { Workbook, sheetToCsv } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

sheet.cell(0, 0).value = 'Name';
sheet.cell(0, 1).value = 'Age';
sheet.cell(1, 0).value = 'Alice';
sheet.cell(1, 1).value = 30;

const csv = sheetToCsv(sheet);
console.log(csv);
// "Name","Age"
// "Alice",30
```

### Export with Options

```typescript
import { sheetToCsv } from 'cellify';

const csv = sheetToCsv(sheet, {
  // Field delimiter
  delimiter: ',',

  // Row delimiter
  rowDelimiter: '\n',

  // Quote character
  quoteChar: '"',

  // Quote all fields (not just those that need escaping)
  quoteAllFields: false,

  // Include UTF-8 BOM for Excel compatibility
  includeBom: true,

  // How to represent null values
  nullValue: '',

  // Date format
  dateFormat: 'ISO', // 'ISO', 'locale', or custom format

  // Export specific range only
  range: 'A1:D10',
});
```

### Export to Buffer (Node.js)

```typescript
import { sheetToCsvBuffer } from 'cellify';
import { writeFileSync } from 'fs';

const buffer = sheetToCsvBuffer(sheet, {
  includeBom: true,
  delimiter: ',',
});

writeFileSync('output.csv', buffer);
```

### Export Multiple Sheets

```typescript
import { sheetsToCsv } from 'cellify';

const csvMap = sheetsToCsv(workbook.sheets, {
  delimiter: ',',
  includeBom: true,
});

// Map of sheet name to CSV string
csvMap.forEach((csv, sheetName) => {
  console.log(`Sheet: ${sheetName}`);
  console.log(csv);
});
```

## Delimiter Detection

Cellify automatically detects the delimiter when importing CSV files:

```typescript
// Auto-detects comma, semicolon, tab, or pipe
const workbook = csvToWorkbook(csvText);
```

Supported delimiters:
- `,` (comma) - default
- `;` (semicolon) - common in European locales
- `\t` (tab) - TSV files
- `|` (pipe)

## Type Detection

### Numbers

Automatically converts numeric strings to numbers:

```typescript
const csv = `Value
123
45.67
-89.01
1,234.56
$99.99
50%`;

const workbook = csvToWorkbook(csv, { detectNumbers: true });
// Values: 123, 45.67, -89.01, 1234.56, 99.99, 0.5
```

### Dates

Automatically converts date strings to Date objects:

```typescript
const csv = `Date
2024-01-15
01/15/2024
15/01/2024`;

const workbook = csvToWorkbook(csv, {
  detectDates: true,
  dateFormats: ['yyyy-mm-dd', 'mm/dd/yyyy', 'dd/mm/yyyy'],
});
```

### Booleans

Automatically converts boolean strings:

```typescript
const csv = `Active
true
false
TRUE
FALSE`;

const workbook = csvToWorkbook(csv);
// Values: true, false, true, false (boolean type)
```

## RFC 4180 Compliance

Cellify's CSV parser follows RFC 4180:

- Fields containing delimiters, quotes, or newlines are quoted
- Quote characters within quoted fields are escaped by doubling
- Multi-line fields are supported

```typescript
const csv = `Name,Description
Widget,"A ""great"" product"
Gadget,"Features:
- Fast
- Reliable"`;

const workbook = csvToWorkbook(csv);
// Correctly parses escaped quotes and multi-line fields
```

## Import Result

```typescript
interface CsvImportResult {
  rowCount: number;
  columnCount: number;
  headers?: string[];
  warnings: string[];
}
```

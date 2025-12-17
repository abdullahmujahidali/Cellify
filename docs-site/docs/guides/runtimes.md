---
sidebar_position: 3
---

# Runtime Support

Cellify works with Node.js, Bun, and Deno out of the box.

## Node.js

```bash
npm install cellify
```

```typescript
import { Workbook, workbookToXlsx, xlsxToWorkbook } from 'cellify';
import { writeFileSync, readFileSync } from 'fs';

// Create and export
const workbook = new Workbook();
workbook.addSheet('Data').cell(0, 0).value = 'Hello Node.js';

const buffer = workbookToXlsx(workbook);
writeFileSync('output.xlsx', buffer);

// Import
const data = readFileSync('output.xlsx');
const result = await xlsxToWorkbook(new Uint8Array(data));
console.log('Cells:', result.stats.totalCells);
```

## Bun

```bash
bun add cellify
```

```typescript
import { Workbook, workbookToXlsx, xlsxToWorkbook } from 'cellify';

// Create and export
const workbook = new Workbook();
workbook.addSheet('Data').cell(0, 0).value = 'Hello Bun';

const buffer = workbookToXlsx(workbook);
Bun.write('output.xlsx', buffer);

// Import
const file = Bun.file('output.xlsx');
const data = await file.arrayBuffer();
const result = await xlsxToWorkbook(new Uint8Array(data));
console.log('Cells:', result.stats.totalCells);
```

## Deno

```typescript
// Import from npm
import {
  Workbook,
  workbookToXlsx,
  xlsxToWorkbook
} from 'npm:cellify';

// Create and export
const workbook = new Workbook();
workbook.addSheet('Data').cell(0, 0).value = 'Hello Deno';

const buffer = workbookToXlsx(workbook);
await Deno.writeFile('output.xlsx', buffer);

// Import
const data = await Deno.readFile('output.xlsx');
const result = await xlsxToWorkbook(data);
console.log('Cells:', result.stats.totalCells);
```

Run with:
```bash
deno run --allow-read --allow-write script.ts
```

## WASM Acceleration

All runtimes support the optional WASM parser for faster imports:

```typescript
import { initXlsxWasm, xlsxToWorkbook } from 'cellify';

// Pre-initialize for instant first import
await initXlsxWasm();

// Now imports are 10x faster
const result = await xlsxToWorkbook(data);
```

WASM automatically falls back to the JavaScript parser if unavailable.

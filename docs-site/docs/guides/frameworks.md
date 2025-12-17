---
sidebar_position: 2
---

# Framework Integration

Cellify works seamlessly with all major JavaScript frameworks. This guide shows you how to integrate Excel import/export functionality into your application.

## React

### Basic Setup

```tsx
import { useState } from 'react';
import {
  Workbook,
  workbookToXlsxBlob,
  xlsxBlobToWorkbook,
  initXlsxWasm,
} from 'cellify';

function ExcelExporter() {
  const handleExport = () => {
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Report');

    // Add headers
    sheet.cell(0, 0).value = 'Name';
    sheet.cell(0, 1).value = 'Email';
    sheet.cell(0, 2).value = 'Sales';

    // Style headers
    for (let col = 0; col < 3; col++) {
      sheet.cell(0, col).applyStyle({
        font: { bold: true, color: '#FFFFFF' },
        fill: { type: 'pattern', pattern: 'solid', fgColor: '#2563eb' },
      });
    }

    // Add data
    const data = [
      ['Alice', 'alice@example.com', 15000],
      ['Bob', 'bob@example.com', 22000],
    ];

    data.forEach((row, i) => {
      row.forEach((value, j) => {
        sheet.cell(i + 1, j).value = value;
      });
    });

    // Download
    const blob = workbookToXlsxBlob(workbook);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'report.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  };

  return <button onClick={handleExport}>Export to Excel</button>;
}
```

### Import Component

```tsx
import { useState } from 'react';
import { xlsxBlobToWorkbook, type XlsxImportResult } from 'cellify';

function ExcelImporter() {
  const [data, setData] = useState<string[][]>([]);

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const result = await xlsxBlobToWorkbook(file);
    const sheet = result.workbook.sheets[0];
    const dims = sheet.dimensions;

    if (!dims) return;

    const rows: string[][] = [];
    for (let r = dims.startRow; r <= dims.endRow; r++) {
      const row: string[] = [];
      for (let c = dims.startCol; c <= dims.endCol; c++) {
        const cell = sheet.getCell(r, c);
        row.push(cell?.value?.toString() ?? '');
      }
      rows.push(row);
    }

    setData(rows);
  };

  return (
    <div>
      <input type="file" accept=".xlsx" onChange={handleFileChange} />
      <table>
        <tbody>
          {data.map((row, i) => (
            <tr key={i}>
              {row.map((cell, j) => (
                <td key={j}>{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
```

### Custom Hook

```tsx
import { useState, useCallback } from 'react';
import {
  Workbook,
  workbookToXlsxBlob,
  xlsxBlobToWorkbook,
  type XlsxImportResult,
} from 'cellify';

export function useExcel() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<Error | null>(null);

  const exportToExcel = useCallback(
    async (data: Record<string, unknown>[], filename = 'export.xlsx') => {
      setLoading(true);
      setError(null);

      try {
        const workbook = new Workbook();
        const sheet = workbook.addSheet('Data');

        // Headers from object keys
        const headers = Object.keys(data[0] || {});
        headers.forEach((header, col) => {
          sheet.cell(0, col).value = header;
          sheet.cell(0, col).applyStyle({ font: { bold: true } });
        });

        // Data rows
        data.forEach((row, rowIndex) => {
          headers.forEach((header, col) => {
            sheet.cell(rowIndex + 1, col).value = row[header] as string | number;
          });
        });

        const blob = workbookToXlsxBlob(workbook);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
      } catch (e) {
        setError(e as Error);
      } finally {
        setLoading(false);
      }
    },
    []
  );

  const importFromExcel = useCallback(
    async (file: File): Promise<Record<string, unknown>[]> => {
      setLoading(true);
      setError(null);

      try {
        const result = await xlsxBlobToWorkbook(file);
        const sheet = result.workbook.sheets[0];
        const dims = sheet.dimensions;

        if (!dims) return [];

        // First row is headers
        const headers: string[] = [];
        for (let c = dims.startCol; c <= dims.endCol; c++) {
          headers.push(sheet.getCell(0, c)?.value?.toString() ?? `col${c}`);
        }

        // Data rows
        const rows: Record<string, unknown>[] = [];
        for (let r = dims.startRow + 1; r <= dims.endRow; r++) {
          const row: Record<string, unknown> = {};
          for (let c = dims.startCol; c <= dims.endCol; c++) {
            row[headers[c - dims.startCol]] = sheet.getCell(r, c)?.value;
          }
          rows.push(row);
        }

        return rows;
      } catch (e) {
        setError(e as Error);
        return [];
      } finally {
        setLoading(false);
      }
    },
    []
  );

  return { exportToExcel, importFromExcel, loading, error };
}
```

## Vue 3

### Basic Setup

```vue
<script setup lang="ts">
import { ref } from 'vue';
import {
  Workbook,
  workbookToXlsxBlob,
  xlsxBlobToWorkbook,
} from 'cellify';

const data = ref<string[][]>([]);

function handleExport() {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Report');

  sheet.cell(0, 0).value = 'Product';
  sheet.cell(0, 1).value = 'Price';

  sheet.cell(1, 0).value = 'Widget';
  sheet.cell(1, 1).value = 29.99;

  const blob = workbookToXlsxBlob(workbook);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'products.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}

async function handleImport(event: Event) {
  const target = event.target as HTMLInputElement;
  const file = target.files?.[0];
  if (!file) return;

  const result = await xlsxBlobToWorkbook(file);
  const sheet = result.workbook.sheets[0];
  const dims = sheet.dimensions;

  if (!dims) return;

  const rows: string[][] = [];
  for (let r = dims.startRow; r <= dims.endRow; r++) {
    const row: string[] = [];
    for (let c = dims.startCol; c <= dims.endCol; c++) {
      row.push(sheet.getCell(r, c)?.value?.toString() ?? '');
    }
    rows.push(row);
  }

  data.value = rows;
}
</script>

<template>
  <div>
    <button @click="handleExport">Export Excel</button>
    <input type="file" accept=".xlsx" @change="handleImport" />

    <table v-if="data.length">
      <tr v-for="(row, i) in data" :key="i">
        <td v-for="(cell, j) in row" :key="j">{{ cell }}</td>
      </tr>
    </table>
  </div>
</template>
```

### Composable

```ts
// useExcel.ts
import { ref } from 'vue';
import {
  Workbook,
  workbookToXlsxBlob,
  xlsxBlobToWorkbook,
} from 'cellify';

export function useExcel() {
  const loading = ref(false);
  const error = ref<Error | null>(null);

  async function exportToExcel(
    data: Record<string, unknown>[],
    filename = 'export.xlsx'
  ) {
    loading.value = true;
    error.value = null;

    try {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Data');

      const headers = Object.keys(data[0] || {});
      headers.forEach((header, col) => {
        sheet.cell(0, col).value = header;
        sheet.cell(0, col).applyStyle({ font: { bold: true } });
      });

      data.forEach((row, rowIndex) => {
        headers.forEach((header, col) => {
          sheet.cell(rowIndex + 1, col).value = row[header] as string | number;
        });
      });

      const blob = workbookToXlsxBlob(workbook);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      error.value = e as Error;
    } finally {
      loading.value = false;
    }
  }

  async function importFromExcel(file: File) {
    loading.value = true;
    error.value = null;

    try {
      const result = await xlsxBlobToWorkbook(file);
      return result.workbook;
    } catch (e) {
      error.value = e as Error;
      return null;
    } finally {
      loading.value = false;
    }
  }

  return { exportToExcel, importFromExcel, loading, error };
}
```

## Angular

### Service

```typescript
// excel.service.ts
import { Injectable } from '@angular/core';
import {
  Workbook,
  workbookToXlsxBlob,
  xlsxBlobToWorkbook,
  type XlsxImportResult,
} from 'cellify';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  exportToExcel(data: Record<string, unknown>[], filename = 'export.xlsx'): void {
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Data');

    const headers = Object.keys(data[0] || {});
    headers.forEach((header, col) => {
      sheet.cell(0, col).value = header;
      sheet.cell(0, col).applyStyle({ font: { bold: true } });
    });

    data.forEach((row, rowIndex) => {
      headers.forEach((header, col) => {
        sheet.cell(rowIndex + 1, col).value = row[header] as string | number;
      });
    });

    const blob = workbookToXlsxBlob(workbook);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  async importFromExcel(file: File): Promise<XlsxImportResult> {
    return xlsxBlobToWorkbook(file);
  }
}
```

### Component

```typescript
// excel.component.ts
import { Component } from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-excel',
  template: `
    <button (click)="export()">Export</button>
    <input type="file" accept=".xlsx" (change)="import($event)" />
  `,
})
export class ExcelComponent {
  constructor(private excelService: ExcelService) {}

  export() {
    const data = [
      { name: 'Alice', sales: 15000 },
      { name: 'Bob', sales: 22000 },
    ];
    this.excelService.exportToExcel(data, 'sales.xlsx');
  }

  async import(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;

    const result = await this.excelService.importFromExcel(file);
    console.log('Imported:', result.stats.totalCells, 'cells');
  }
}
```

## Vanilla JavaScript

### Browser (ES Modules)

```html
<!DOCTYPE html>
<html>
<head>
  <title>Excel Export</title>
</head>
<body>
  <button id="export">Export to Excel</button>
  <input type="file" id="import" accept=".xlsx" />

  <script type="module">
    import {
      Workbook,
      workbookToXlsxBlob,
      xlsxBlobToWorkbook,
    } from 'https://unpkg.com/cellify/dist/esm/index.js';

    document.getElementById('export').onclick = () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Data');

      sheet.cell(0, 0).value = 'Hello';
      sheet.cell(0, 1).value = 'World';

      const blob = workbookToXlsxBlob(workbook);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'hello.xlsx';
      a.click();
    };

    document.getElementById('import').onchange = async (e) => {
      const file = e.target.files[0];
      const result = await xlsxBlobToWorkbook(file);
      console.log('Cells:', result.stats.totalCells);
    };
  </script>
</body>
</html>
```

## Next Steps

- [Excel Import/Export](./excel.md) - Detailed Excel features
- [Styling Guide](./styling.md) - Cell formatting options
- [Runtime Support](./runtimes.md) - Node.js, Bun, and Deno

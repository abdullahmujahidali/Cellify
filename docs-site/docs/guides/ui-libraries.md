---
sidebar_position: 4
---

# UI Library Integration

Cellify handles Excel import/export while UI libraries handle visualization and editing. This guide shows how to integrate Cellify with popular React table/grid libraries.

## ReactGrid (@silevis/reactgrid)

[ReactGrid](https://www.npmjs.com/package/@silevis/reactgrid) provides a spreadsheet-like editing experience.

### Install

```bash
npm install cellify @silevis/reactgrid
```

### Convert Cellify to ReactGrid

```tsx
import { ReactGrid, Column, Row, CellChange, TextCell } from '@silevis/reactgrid';
import { Workbook, Sheet, workbookToXlsxBlob, xlsxBlobToWorkbook } from 'cellify';

// Convert Cellify Sheet to ReactGrid format
function sheetToReactGrid(sheet: Sheet): { columns: Column[]; rows: Row[] } {
  const dims = sheet.dimensions;
  if (!dims) return { columns: [], rows: [] };

  // Create columns
  const columns: Column[] = [];
  for (let c = dims.startCol; c <= dims.endCol; c++) {
    columns.push({ columnId: c, width: 150 });
  }

  // Create rows
  const rows: Row[] = [];
  for (let r = dims.startRow; r <= dims.endRow; r++) {
    const cells: TextCell[] = [];
    for (let c = dims.startCol; c <= dims.endCol; c++) {
      const cell = sheet.getCell(r, c);
      cells.push({
        type: 'text',
        text: cell?.value?.toString() ?? '',
      });
    }
    rows.push({ rowId: r, cells });
  }

  return { columns, rows };
}

// Convert ReactGrid changes back to Cellify
function applyChangesToSheet(sheet: Sheet, changes: CellChange<TextCell>[]) {
  changes.forEach((change) => {
    const row = change.rowId as number;
    const col = change.columnId as number;
    sheet.cell(row, col).value = change.newCell.text;
  });
}
```

### Full Example

```tsx
import { useState, useCallback } from 'react';
import { ReactGrid, CellChange, TextCell } from '@silevis/reactgrid';
import { Workbook, workbookToXlsxBlob, xlsxBlobToWorkbook } from 'cellify';

function SpreadsheetEditor() {
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<Column[]>([]);

  // Import Excel file
  const handleImport = async (file: File) => {
    const result = await xlsxBlobToWorkbook(file);
    setWorkbook(result.workbook);

    const sheet = result.workbook.sheets[0];
    const { columns, rows } = sheetToReactGrid(sheet);
    setColumns(columns);
    setRows(rows);
  };

  // Handle cell edits
  const handleChanges = (changes: CellChange<TextCell>[]) => {
    if (!workbook) return;

    // Update ReactGrid state
    setRows((prev) => {
      const newRows = [...prev];
      changes.forEach((change) => {
        const rowIdx = newRows.findIndex((r) => r.rowId === change.rowId);
        const colIdx = columns.findIndex((c) => c.columnId === change.columnId);
        if (rowIdx >= 0 && colIdx >= 0) {
          newRows[rowIdx].cells[colIdx] = change.newCell;
        }
      });
      return newRows;
    });

    // Update Cellify workbook
    const sheet = workbook.sheets[0];
    applyChangesToSheet(sheet, changes);
  };

  // Export to Excel
  const handleExport = () => {
    if (!workbook) return;
    const blob = workbookToXlsxBlob(workbook);
    // Download or save to API
  };

  // Save to API
  const handleSaveToApi = async () => {
    if (!workbook) return;
    const blob = workbookToXlsxBlob(workbook);
    await fetch('/api/save', {
      method: 'POST',
      body: blob,
    });
  };

  return (
    <div>
      <input type="file" accept=".xlsx" onChange={(e) => handleImport(e.target.files![0])} />
      <button onClick={handleExport}>Download Excel</button>
      <button onClick={handleSaveToApi}>Save to API</button>

      {rows.length > 0 && (
        <ReactGrid rows={rows} columns={columns} onCellsChanged={handleChanges} />
      )}
    </div>
  );
}
```

## TanStack Table

[TanStack Table](https://tanstack.com/table) is a headless UI library for building tables.

### Install

```bash
npm install cellify @tanstack/react-table
```

### Convert Cellify to TanStack Table

```tsx
import { useReactTable, getCoreRowModel, ColumnDef } from '@tanstack/react-table';
import { Sheet, Workbook, workbookToXlsxBlob, xlsxBlobToWorkbook } from 'cellify';

// Convert Cellify Sheet to TanStack Table format
function sheetToTableData(sheet: Sheet): {
  columns: ColumnDef<Record<string, unknown>>[];
  data: Record<string, unknown>[];
} {
  const dims = sheet.dimensions;
  if (!dims) return { columns: [], data: [] };

  // First row as headers
  const headers: string[] = [];
  for (let c = dims.startCol; c <= dims.endCol; c++) {
    headers.push(sheet.getCell(0, c)?.value?.toString() ?? `col${c}`);
  }

  // Create columns
  const columns: ColumnDef<Record<string, unknown>>[] = headers.map((header, idx) => ({
    accessorKey: header,
    header: header,
    cell: ({ getValue }) => getValue(),
  }));

  // Create data rows (skip header)
  const data: Record<string, unknown>[] = [];
  for (let r = dims.startRow + 1; r <= dims.endRow; r++) {
    const row: Record<string, unknown> = {};
    for (let c = dims.startCol; c <= dims.endCol; c++) {
      row[headers[c - dims.startCol]] = sheet.getCell(r, c)?.value;
    }
    data.push(row);
  }

  return { columns, data };
}

// Convert TanStack Table data back to Cellify
function tableDataToSheet(
  workbook: Workbook,
  headers: string[],
  data: Record<string, unknown>[]
): Sheet {
  const sheet = workbook.addSheet('Data');

  // Write headers
  headers.forEach((header, col) => {
    sheet.cell(0, col).value = header;
    sheet.cell(0, col).applyStyle({ font: { bold: true } });
  });

  // Write data
  data.forEach((row, rowIdx) => {
    headers.forEach((header, col) => {
      sheet.cell(rowIdx + 1, col).value = row[header] as string | number;
    });
  });

  return sheet;
}
```

### Full Example

```tsx
import { useState, useMemo } from 'react';
import { useReactTable, getCoreRowModel, flexRender } from '@tanstack/react-table';
import { Workbook, workbookToXlsxBlob, xlsxBlobToWorkbook } from 'cellify';

function DataTable() {
  const [data, setData] = useState<Record<string, unknown>[]>([]);
  const [columns, setColumns] = useState<ColumnDef<Record<string, unknown>>[]>([]);

  // Import Excel
  const handleImport = async (file: File) => {
    const result = await xlsxBlobToWorkbook(file);
    const sheet = result.workbook.sheets[0];
    const { columns, data } = sheetToTableData(sheet);
    setColumns(columns);
    setData(data);
  };

  // Export to Excel
  const handleExport = () => {
    const workbook = new Workbook();
    const headers = columns.map((c) => c.accessorKey as string);
    tableDataToSheet(workbook, headers, data);

    const blob = workbookToXlsxBlob(workbook);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'export.xlsx';
    a.click();
  };

  // Save to API
  const handleSaveToApi = async () => {
    const workbook = new Workbook();
    const headers = columns.map((c) => c.accessorKey as string);
    tableDataToSheet(workbook, headers, data);

    const blob = workbookToXlsxBlob(workbook);
    await fetch('/api/spreadsheets', {
      method: 'POST',
      headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body: blob,
    });
  };

  const table = useReactTable({
    data,
    columns,
    getCoreRowModel: getCoreRowModel(),
  });

  return (
    <div>
      <input type="file" accept=".xlsx" onChange={(e) => handleImport(e.target.files![0])} />
      <button onClick={handleExport}>Download Excel</button>
      <button onClick={handleSaveToApi}>Save to API</button>

      <table>
        <thead>
          {table.getHeaderGroups().map((headerGroup) => (
            <tr key={headerGroup.id}>
              {headerGroup.headers.map((header) => (
                <th key={header.id}>
                  {flexRender(header.column.columnDef.header, header.getContext())}
                </th>
              ))}
            </tr>
          ))}
        </thead>
        <tbody>
          {table.getRowModel().rows.map((row) => (
            <tr key={row.id}>
              {row.getVisibleCells().map((cell) => (
                <td key={cell.id}>
                  {flexRender(cell.column.columnDef.cell, cell.getContext())}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
```

## Save to API

Both examples above show a "Save to API" pattern. Here's a complete example:

```typescript
// Save workbook as blob to your API
async function saveToApi(workbook: Workbook, endpoint: string) {
  const blob = workbookToXlsxBlob(workbook);

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    body: blob,
  });

  if (!response.ok) {
    throw new Error('Failed to save');
  }

  return response.json();
}

// Load workbook from API
async function loadFromApi(endpoint: string): Promise<Workbook> {
  const response = await fetch(endpoint);
  const blob = await response.blob();
  const result = await xlsxBlobToWorkbook(blob);
  return result.workbook;
}
```

## Extract Data for Submission

If you need cell-level data for form submission:

```typescript
// Get data in cell/row/value format
function extractCellData(sheet: Sheet) {
  const cells = [];
  for (const cell of sheet.cells()) {
    cells.push({
      cellId: cell.address,  // "A1", "B2"
      row: cell.row,
      col: cell.col,
      value: cell.value,
    });
  }
  return cells;
}

// Submit to API
async function submitData(sheet: Sheet) {
  const cellData = extractCellData(sheet);
  await fetch('/api/submit', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ cells: cellData }),
  });
}
```

## Other Compatible Libraries

Cellify works with any table/grid library. The pattern is always:

1. **Import**: `xlsxBlobToWorkbook()` → convert to library format
2. **Edit**: User edits in the UI library
3. **Export**: Convert back to Cellify → `workbookToXlsxBlob()`

Compatible libraries include:
- [AG Grid](https://www.ag-grid.com/)
- [Handsontable](https://handsontable.com/)
- [React Data Grid](https://adazzle.github.io/react-data-grid/)
- [Material UI DataGrid](https://mui.com/x/react-data-grid/)

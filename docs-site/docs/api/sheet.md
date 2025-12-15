---
sidebar_position: 2
---

# Sheet

The `Sheet` class represents a worksheet within a workbook. Sheets contain cells organized in rows and columns.

## Properties

### name

Get or set the sheet name.

```typescript
get name(): string
set name(value: string)
```

```typescript
console.log(sheet.name);
sheet.name = 'Sales Data';
```

### dimensions

Get the used range of the sheet.

```typescript
get dimensions(): RangeDefinition | null
```

Returns `null` if the sheet is empty.

```typescript
const dims = sheet.dimensions;
if (dims) {
  console.log(`Data from row ${dims.startRow} to ${dims.endRow}`);
  console.log(`Data from col ${dims.startCol} to ${dims.endCol}`);
}
```

### rowCount / columnCount / cellCount

Get counts of data in the sheet.

```typescript
get rowCount(): number
get columnCount(): number
get cellCount(): number
```

```typescript
console.log('Rows:', sheet.rowCount);
console.log('Columns:', sheet.columnCount);
console.log('Total cells:', sheet.cellCount);
```

## Cell Access

### cell()

Get or create a cell. Creates the cell if it doesn't exist.

```typescript
// By A1 notation
cell(address: string): Cell

// By row and column (0-based)
cell(row: number, col: number): Cell
```

```typescript
// A1 notation
sheet.cell('A1').value = 'Hello';
sheet.cell('B2').value = 42;

// Row/column indices
sheet.cell(0, 0).value = 'Hello'; // A1
sheet.cell(1, 1).value = 42;      // B2
```

### getCell()

Get a cell if it exists, without creating it.

```typescript
getCell(address: string): Cell | undefined
getCell(row: number, col: number): Cell | undefined
```

```typescript
const cell = sheet.getCell('A1');
if (cell) {
  console.log(cell.value);
}
```

### hasCell()

Check if a cell exists.

```typescript
hasCell(address: string): boolean
hasCell(row: number, col: number): boolean
```

```typescript
if (sheet.hasCell('A1')) {
  console.log('Cell A1 exists');
}
```

### deleteCell()

Delete a cell.

```typescript
deleteCell(address: string): boolean
deleteCell(row: number, col: number): boolean
```

```typescript
sheet.deleteCell('A1');
```

### cells()

Iterate over all cells.

```typescript
*cells(): Generator<Cell>
```

```typescript
for (const cell of sheet.cells()) {
  console.log(`${cell.address}: ${cell.value}`);
}
```

### cellsInRange()

Iterate over cells in a range.

```typescript
*cellsInRange(range: string | RangeDefinition): Generator<Cell>
```

```typescript
for (const cell of sheet.cellsInRange('A1:D10')) {
  console.log(cell.value);
}
```

## Bulk Operations

### setValues()

Set values from a 2D array.

```typescript
setValues(startAddress: string, values: CellValue[][]): this
setValues(startRow: number, startCol: number, values: CellValue[][]): this
```

```typescript
sheet.setValues('A1', [
  ['Name', 'Age', 'City'],
  ['Alice', 30, 'NYC'],
  ['Bob', 25, 'LA'],
]);

// Or with indices
sheet.setValues(0, 0, [
  ['Name', 'Age', 'City'],
  ['Alice', 30, 'NYC'],
]);
```

### getValues()

Get values as a 2D array.

```typescript
getValues(range: string | RangeDefinition): CellValue[][]
```

```typescript
const values = sheet.getValues('A1:C3');
console.log(values);
// [['Name', 'Age', 'City'], ['Alice', 30, 'NYC'], ['Bob', 25, 'LA']]
```

### applyStyle()

Apply a style to a range of cells.

```typescript
applyStyle(range: string | RangeDefinition, style: CellStyle): this
```

```typescript
sheet.applyStyle('A1:D1', {
  font: { bold: true },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#F3F4F6' },
});
```

### clearRange()

Clear all cells in a range.

```typescript
clearRange(range: string | RangeDefinition): this
```

```typescript
sheet.clearRange('A1:D10');
```

## Merge Operations

### mergeCells()

Merge cells in a range.

```typescript
mergeCells(range: string | RangeDefinition): this
```

**Throws:** Error if the range overlaps with an existing merge.

```typescript
sheet.cell('A1').value = 'Title';
sheet.mergeCells('A1:D1');

// Or with range definition
sheet.mergeCells({
  startRow: 0,
  startCol: 0,
  endRow: 0,
  endCol: 3,
});
```

### unmergeCells()

Unmerge a merged range.

```typescript
unmergeCells(range: string | RangeDefinition): this
```

**Throws:** Error if no merge exists at the specified range.

```typescript
sheet.unmergeCells('A1:D1');
```

### merges

Get all merge ranges.

```typescript
get merges(): readonly MergeRange[]
```

```typescript
sheet.merges.forEach(merge => {
  console.log(`Merge: ${merge.startRow},${merge.startCol} to ${merge.endRow},${merge.endCol}`);
});
```

## Row Configuration

### getRow()

Get row configuration.

```typescript
getRow(index: number): RowConfig
```

```typescript
interface RowConfig {
  height?: number;
  hidden?: boolean;
  outlineLevel?: number;
  style?: CellStyle;
}
```

### setRow()

Set row configuration.

```typescript
setRow(index: number, config: RowConfig): this
```

```typescript
sheet.setRow(0, {
  height: 30,
  style: { font: { bold: true } },
});
```

### setRowHeight()

Set row height.

```typescript
setRowHeight(index: number, height: number): this
```

```typescript
sheet.setRowHeight(0, 25); // 25 points
```

### hideRow() / showRow()

Hide or show a row.

```typescript
hideRow(index: number): this
showRow(index: number): this
```

```typescript
sheet.hideRow(5);
sheet.showRow(5);
```

### rows

Get all row configurations.

```typescript
get rows(): ReadonlyMap<number, RowConfig>
```

## Column Configuration

### getColumn()

Get column configuration.

```typescript
getColumn(index: number): ColumnConfig
```

```typescript
interface ColumnConfig {
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
  style?: CellStyle;
}
```

### setColumn()

Set column configuration.

```typescript
setColumn(index: number, config: ColumnConfig): this
```

```typescript
sheet.setColumn(0, {
  width: 20,
  style: { alignment: { horizontal: 'left' } },
});
```

### setColumnWidth()

Set column width.

```typescript
setColumnWidth(index: number, width: number): this
```

```typescript
sheet.setColumnWidth(0, 15); // 15 characters
```

### hideColumn() / showColumn()

Hide or show a column.

```typescript
hideColumn(index: number): this
showColumn(index: number): this
```

```typescript
sheet.hideColumn(2);
sheet.showColumn(2);
```

### columns

Get all column configurations.

```typescript
get columns(): ReadonlyMap<number, ColumnConfig>
```

## View Configuration

### view

Get sheet view settings.

```typescript
get view(): SheetView
```

```typescript
interface SheetView {
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showZeros?: boolean;
  tabSelected?: boolean;
  zoomScale?: number;
  frozenRows?: number;
  frozenCols?: number;
  splitRow?: number;
  splitCol?: number;
}
```

### setView()

Set sheet view settings.

```typescript
setView(view: Partial<SheetView>): this
```

```typescript
sheet.setView({
  showGridLines: true,
  zoomScale: 100,
});
```

### freeze()

Freeze rows and columns.

```typescript
freeze(rows: number, cols?: number): this
```

```typescript
sheet.freeze(1);     // Freeze top row
sheet.freeze(1, 1);  // Freeze top row and first column
sheet.freeze(0, 2);  // Freeze first two columns only
```

### unfreeze()

Remove freeze panes.

```typescript
unfreeze(): this
```

```typescript
sheet.unfreeze();
```

## Auto Filter

### setAutoFilter()

Enable auto filter on a range.

```typescript
setAutoFilter(range: string | RangeDefinition): this
```

```typescript
sheet.setAutoFilter('A1:D100');
```

### removeAutoFilter()

Remove auto filter.

```typescript
removeAutoFilter(): this
```

### autoFilter

Get auto filter configuration.

```typescript
get autoFilter(): AutoFilter | undefined
```

## Conditional Formatting

### addConditionalFormat()

Add a conditional formatting rule.

```typescript
addConditionalFormat(rule: ConditionalFormatRule): this
```

### conditionalFormats

Get all conditional formatting rules.

```typescript
get conditionalFormats(): readonly ConditionalFormatRule[]
```

### clearConditionalFormats()

Remove all conditional formatting.

```typescript
clearConditionalFormats(): this
```

## Protection

### protect()

Protect the sheet.

```typescript
protect(options?: SheetProtection): this
```

```typescript
interface SheetProtection {
  password?: string;
  sheet?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  // ... more options
}
```

```typescript
sheet.protect({
  formatCells: false,
  insertRows: true,
});
```

### unprotect()

Remove sheet protection.

```typescript
unprotect(): this
```

### protection / isProtected

Get protection settings.

```typescript
get protection(): SheetProtection | undefined
get isProtected(): boolean
```

## Page Setup

### pageSetup

Get page setup configuration.

```typescript
get pageSetup(): PageSetup
```

```typescript
interface PageSetup {
  paperSize?: number;
  orientation?: 'portrait' | 'landscape';
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  margins?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
    header?: number;
    footer?: number;
  };
}
```

### setPageSetup()

Set page setup configuration.

```typescript
setPageSetup(setup: Partial<PageSetup>): this
```

```typescript
sheet.setPageSetup({
  orientation: 'landscape',
  margins: { top: 1, bottom: 1, left: 0.75, right: 0.75 },
});
```

## Serialization

### toJSON()

Convert sheet to JSON.

```typescript
toJSON(): Record<string, unknown>
```

```typescript
const json = sheet.toJSON();
```

## Example

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Report');

// Set column widths
[15, 20, 12, 15].forEach((width, col) => {
  sheet.setColumnWidth(col, width);
});

// Add header row
const headers = ['Product', 'Description', 'Qty', 'Price'];
headers.forEach((h, col) => {
  sheet.cell(0, col).value = h;
});

// Style header
sheet.applyStyle('A1:D1', {
  font: { bold: true, color: '#FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#059669' },
});
sheet.setRowHeight(0, 25);

// Add data
sheet.setValues(1, 0, [
  ['Widget A', 'Premium widget', 100, 29.99],
  ['Widget B', 'Standard widget', 250, 19.99],
]);

// Freeze header
sheet.freeze(1);

// Enable filter
sheet.setAutoFilter('A1:D3');

// Merge title
sheet.cell(5, 0).value = 'Total Items: 350';
sheet.mergeCells('A6:D6');
```

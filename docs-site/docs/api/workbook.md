---
sidebar_position: 1
---

# Workbook

The `Workbook` class is the top-level container that holds sheets, properties, and workbook-level settings.

## Constructor

```typescript
const workbook = new Workbook();
```

Creates a new empty workbook with a creation date automatically set.

## Sheet Management

### addSheet()

Add a new sheet to the workbook.

```typescript
addSheet(name?: string): Sheet
```

**Parameters:**
- `name` (optional): Sheet name. Auto-generates "Sheet1", "Sheet2", etc. if not provided.

**Returns:** The newly created `Sheet` instance.

**Throws:** Error if name is invalid (empty, >31 chars, contains invalid characters, or already exists).

```typescript
const sheet = workbook.addSheet('Sales');
const sheet2 = workbook.addSheet(); // Creates "Sheet1"
```

### getSheet()

Get a sheet by name.

```typescript
getSheet(name: string): Sheet | undefined
```

```typescript
const sheet = workbook.getSheet('Sales');
if (sheet) {
  console.log('Found sheet:', sheet.name);
}
```

### getSheetByIndex()

Get a sheet by its 0-based index.

```typescript
getSheetByIndex(index: number): Sheet | undefined
```

```typescript
const firstSheet = workbook.getSheetByIndex(0);
```

### getSheetIndex()

Get the index of a sheet.

```typescript
getSheetIndex(sheet: Sheet | string): number
```

**Returns:** Index (0-based), or -1 if not found.

```typescript
const index = workbook.getSheetIndex('Sales');
```

### removeSheet()

Remove a sheet from the workbook.

```typescript
removeSheet(sheet: Sheet | string): boolean
```

**Returns:** `true` if removed, `false` if not found.

```typescript
workbook.removeSheet('Sales');
workbook.removeSheet(sheet);
```

### renameSheet()

Rename a sheet.

```typescript
renameSheet(oldName: string, newName: string): boolean
```

**Returns:** `true` if renamed, `false` if sheet not found.

**Throws:** Error if new name is invalid.

```typescript
workbook.renameSheet('Sheet1', 'Sales Data');
```

### moveSheet()

Move a sheet to a new position.

```typescript
moveSheet(sheet: Sheet | string, newIndex: number): boolean
```

```typescript
workbook.moveSheet('Summary', 0); // Move to first position
```

### duplicateSheet()

Create a copy of a sheet.

```typescript
duplicateSheet(sheet: Sheet | string, newName?: string): Sheet | undefined
```

```typescript
const copy = workbook.duplicateSheet('Template', 'January');
```

### sheets

Get all sheets as a readonly array.

```typescript
get sheets(): readonly Sheet[]
```

```typescript
workbook.sheets.forEach(sheet => {
  console.log(sheet.name);
});
```

### sheetCount

Get the number of sheets.

```typescript
get sheetCount(): number
```

```typescript
console.log('Sheets:', workbook.sheetCount);
```

## Properties

### properties

Get workbook properties.

```typescript
get properties(): WorkbookProperties
```

```typescript
interface WorkbookProperties {
  title?: string;
  subject?: string;
  author?: string;
  company?: string;
  category?: string;
  keywords?: string[];
  comments?: string;
  manager?: string;
  created?: Date;
  modified?: Date;
  lastModifiedBy?: string;
  revision?: number;
}
```

### setProperties()

Set workbook properties.

```typescript
setProperties(props: Partial<WorkbookProperties>): this
```

```typescript
workbook.setProperties({
  title: 'Annual Report',
  author: 'Finance Team',
  keywords: ['report', 'annual', '2024'],
});
```

### title / author

Convenience getters/setters for common properties.

```typescript
workbook.title = 'My Workbook';
workbook.author = 'John Doe';

console.log(workbook.title);
console.log(workbook.author);
```

## Named Styles

### addNamedStyle()

Create a reusable style.

```typescript
addNamedStyle(name: string, style: CellStyle): this
```

```typescript
workbook.addNamedStyle('header', {
  font: { bold: true, size: 14 },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#4F46E5' },
});
```

### getNamedStyle()

Get a named style.

```typescript
getNamedStyle(name: string): NamedStyle | undefined
```

```typescript
const style = workbook.getNamedStyle('header');
if (style) {
  sheet.cell(0, 0).style = style.style;
}
```

### removeNamedStyle()

Remove a named style.

```typescript
removeNamedStyle(name: string): boolean
```

### namedStyles

Get all named styles.

```typescript
get namedStyles(): ReadonlyMap<string, NamedStyle>
```

## Defined Names

### addDefinedName()

Create a named range or constant.

```typescript
addDefinedName(
  name: string,
  formula: string,
  options?: { scope?: string; comment?: string; hidden?: boolean }
): this
```

```typescript
// Named range
workbook.addDefinedName('SalesData', 'Sheet1!$A$1:$D$100');

// Named constant
workbook.addDefinedName('TaxRate', '0.08');

// Sheet-scoped name
workbook.addDefinedName('LocalRange', '$A$1:$A$10', { scope: 'Sheet1' });
```

### getDefinedName()

Get a defined name.

```typescript
getDefinedName(name: string): DefinedName | undefined
```

```typescript
interface DefinedName {
  name: string;
  formula: string;
  scope?: string;
  comment?: string;
  hidden?: boolean;
}
```

### removeDefinedName()

Remove a defined name.

```typescript
removeDefinedName(name: string): boolean
```

### definedNames

Get all defined names.

```typescript
get definedNames(): ReadonlyMap<string, DefinedName>
```

## Calculation Settings

### calculationMode

Get or set the calculation mode.

```typescript
get calculationMode(): CalculationMode
set calculationMode(mode: CalculationMode)
```

```typescript
type CalculationMode = 'auto' | 'manual' | 'autoNoTable';
```

```typescript
workbook.calculationMode = 'auto';
```

## Workbook View

### view

Get workbook view settings.

```typescript
get view(): WorkbookView
```

```typescript
interface WorkbookView {
  activeSheet?: number;
  firstSheet?: number;
  showSheetTabs?: boolean;
  tabRatio?: number;
}
```

### setView()

Set workbook view settings.

```typescript
setView(view: Partial<WorkbookView>): this
```

```typescript
workbook.setView({
  activeSheet: 0,
  showSheetTabs: true,
});
```

### activeSheet

Get the active sheet.

```typescript
get activeSheet(): Sheet | undefined
```

### setActiveSheet()

Set the active sheet.

```typescript
setActiveSheet(sheet: Sheet | string | number): this
```

```typescript
workbook.setActiveSheet('Sales');
workbook.setActiveSheet(0);
```

## Utility Methods

### touch()

Update the modified timestamp.

```typescript
touch(): this
```

```typescript
workbook.touch();
```

### toJSON()

Convert workbook to a JSON representation.

```typescript
toJSON(): Record<string, unknown>
```

```typescript
const json = workbook.toJSON();
console.log(JSON.stringify(json, null, 2));
```

### fromJSON()

Create a workbook from JSON (static method).

```typescript
static fromJSON(json: Record<string, unknown>): Workbook
```

## Example

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

// Create workbook
const workbook = new Workbook();

// Set properties
workbook.setProperties({
  title: 'Sales Report 2024',
  author: 'Sales Team',
});

// Add named style
workbook.addNamedStyle('currency', {
  numberFormat: { formatCode: '$#,##0.00' },
});

// Add sheets
const summary = workbook.addSheet('Summary');
const details = workbook.addSheet('Details');

// Add defined name
workbook.addDefinedName('AllSales', 'Details!$A$1:$D$100');

// Set active sheet
workbook.setActiveSheet('Summary');

// Export
const blob = workbookToXlsxBlob(workbook);
```

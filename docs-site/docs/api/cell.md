---
sidebar_position: 3
---

# Cell

The `Cell` class represents a single cell in a spreadsheet. Cells can hold values, formulas, styles, and other metadata.

## Properties

### row / col

The cell's position (0-based indices).

```typescript
readonly row: number
readonly col: number
```

```typescript
console.log(`Cell at row ${cell.row}, column ${cell.col}`);
```

### address

The cell's A1 notation address.

```typescript
get address(): string
```

```typescript
console.log(cell.address); // "A1", "B2", etc.
```

## Value

### value

Get or set the cell's value.

```typescript
get value(): CellValue
set value(val: CellValue)
```

Setting a value clears any existing formula.

```typescript
cell.value = 'Hello';        // String
cell.value = 42;             // Number
cell.value = true;           // Boolean
cell.value = new Date();     // Date
cell.value = null;           // Clear value
cell.value = '#DIV/0!';      // Error
```

### type

Get the type of the cell's value.

```typescript
get type(): CellValueType
```

```typescript
type CellValueType = 'string' | 'number' | 'boolean' | 'date' | 'error' | 'formula' | 'null';
```

```typescript
console.log(cell.type); // 'string', 'number', etc.
```

### isEmpty

Check if the cell has no content or styling.

```typescript
get isEmpty(): boolean
```

```typescript
if (cell.isEmpty) {
  console.log('Cell is empty');
}
```

## Formula

### formula

Get the cell's formula.

```typescript
get formula(): CellFormula | undefined
```

```typescript
interface CellFormula {
  formula: string;      // Formula text without '='
  result?: CellValue;   // Cached result
  sharedIndex?: number; // For shared formulas
}
```

```typescript
if (cell.formula) {
  console.log('Formula:', cell.formula.formula);
  console.log('Result:', cell.formula.result);
}
```

### setFormula()

Set a formula on the cell.

```typescript
setFormula(formulaText: string): this
```

The leading `=` is optional.

```typescript
cell.setFormula('=SUM(A1:A10)');
cell.setFormula('A1+B1'); // Also works without '='
```

### clearFormula()

Remove the formula from the cell.

```typescript
clearFormula(): this
```

```typescript
cell.clearFormula();
```

## Style

### style

Get or set the cell's style (replaces existing style).

```typescript
get style(): CellStyle | undefined
set style(style: CellStyle | undefined)
```

```typescript
cell.style = {
  font: { bold: true },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FF0000' },
};
```

### applyStyle()

Apply partial style updates (merges with existing style).

```typescript
applyStyle(style: Partial<CellStyle>): this
```

```typescript
cell.applyStyle({ font: { bold: true } });
cell.applyStyle({ font: { color: '#FF0000' } }); // Keeps bold, adds color
```

## Hyperlink

### hyperlink

Get the cell's hyperlink.

```typescript
get hyperlink(): CellHyperlink | undefined
```

```typescript
interface CellHyperlink {
  target: string;     // URL, file path, or internal reference
  tooltip?: string;
  display?: string;   // Display text
}
```

### setHyperlink()

Set a hyperlink on the cell.

```typescript
setHyperlink(target: string, tooltip?: string): this
```

```typescript
cell.value = 'Visit Google';
cell.setHyperlink('https://google.com', 'Open Google');

// Internal link
cell.setHyperlink('#Sheet2!A1', 'Go to Sheet2');

// Email
cell.setHyperlink('mailto:test@example.com');
```

### clearHyperlink()

Remove the hyperlink.

```typescript
clearHyperlink(): this
```

## Comment

### comment

Get the cell's comment.

```typescript
get comment(): CellComment | undefined
```

```typescript
interface CellComment {
  text: string | RichTextValue;
  author?: string;
  visible?: boolean;
}
```

### setComment()

Add a comment to the cell.

```typescript
setComment(text: string | RichTextValue, author?: string): this
```

```typescript
cell.setComment('This value needs review', 'John Doe');
```

### clearComment()

Remove the comment.

```typescript
clearComment(): this
```

## Data Validation

### validation

Get the cell's validation rules.

```typescript
get validation(): CellValidation | undefined
```

```typescript
interface CellValidation {
  type: ValidationType;
  operator?: ValidationOperator;
  formula1?: string | number | Date;
  formula2?: string | number | Date;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showInputMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  showErrorMessage?: boolean;
  errorStyle?: ValidationErrorStyle;
  errorTitle?: string;
  errorMessage?: string;
}
```

### setValidation()

Set data validation on the cell.

```typescript
setValidation(validation: CellValidation): this
```

```typescript
// Dropdown list
cell.setValidation({
  type: 'list',
  formula1: '"Yes,No,Maybe"',
  showDropDown: true,
  errorMessage: 'Please select from the list',
});

// Number range
cell.setValidation({
  type: 'whole',
  operator: 'between',
  formula1: 1,
  formula2: 100,
  errorMessage: 'Enter a number between 1 and 100',
});
```

### clearValidation()

Remove data validation.

```typescript
clearValidation(): this
```

## Merge Information

### merge

Get merge information (only on master cell).

```typescript
get merge(): MergeRange | undefined
```

```typescript
interface MergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
```

### isMergeMaster

Check if this cell is the top-left cell of a merge.

```typescript
get isMergeMaster(): boolean
```

### mergedInto

Get the master cell address if this cell is part of a merge.

```typescript
get mergedInto(): CellAddress | undefined
```

### isMergedSlave

Check if this cell is part of a merge (but not the master).

```typescript
get isMergedSlave(): boolean
```

### isMerged

Check if this cell is part of any merge.

```typescript
get isMerged(): boolean
```

```typescript
if (cell.isMerged) {
  if (cell.isMergeMaster) {
    console.log('This is the master cell of a merge');
  } else {
    console.log('This cell is merged into:', cell.mergedInto);
  }
}
```

## Utility Methods

### clear()

Clear all content and styling from the cell.

```typescript
clear(): this
```

```typescript
cell.clear(); // Removes value, formula, style, hyperlink, comment, validation
```

Note: Merge information is managed by the Sheet and not cleared by this method.

### clone()

Create a deep copy of the cell.

```typescript
clone(): Cell
```

```typescript
const copy = cell.clone();
```

### toJSON()

Convert cell to a JSON representation.

```typescript
toJSON(): Record<string, unknown>
```

```typescript
const json = cell.toJSON();
console.log(json);
// {
//   row: 0,
//   col: 0,
//   address: 'A1',
//   value: 'Hello',
//   type: 'string',
//   style: { ... }
// }
```

## Cell Value Types

### Supported Types

| Type | Example | Notes |
|------|---------|-------|
| String | `'Hello'` | Text values |
| Number | `42`, `3.14` | Numeric values |
| Boolean | `true`, `false` | Boolean values |
| Date | `new Date()` | JavaScript Date objects |
| Null | `null` | Empty cell |
| Error | `'#DIV/0!'` | Excel error values |
| Rich Text | `{ richText: [...] }` | Formatted text segments |

### Error Types

```typescript
type CellErrorType =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA';
```

### Rich Text

```typescript
interface RichTextValue {
  richText: RichTextRun[];
}

interface RichTextRun {
  text: string;
  font?: {
    name?: string;
    size?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
  };
}
```

```typescript
cell.value = {
  richText: [
    { text: 'Hello ', font: { bold: true } },
    { text: 'World', font: { italic: true, color: '#FF0000' } },
  ],
};
```

## Example

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Demo');

// Basic value
const cell = sheet.cell(0, 0);
cell.value = 'Product Name';

// Styled cell
sheet.cell(0, 1).value = 'Price';
sheet.cell(0, 1).applyStyle({
  font: { bold: true, color: '#FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#059669' },
  alignment: { horizontal: 'center' },
});

// Formula
sheet.cell(1, 0).value = 100;
sheet.cell(1, 1).value = 0.08;
sheet.cell(1, 2).setFormula('A2*B2'); // Tax amount

// Hyperlink
sheet.cell(2, 0).value = 'View Details';
sheet.cell(2, 0).setHyperlink('https://example.com/product/1');

// Comment
sheet.cell(3, 0).value = 1500;
sheet.cell(3, 0).setComment('Q4 projection - needs review');

// Data validation
sheet.cell(4, 0).setValidation({
  type: 'list',
  formula1: '"Pending,Approved,Rejected"',
  showDropDown: true,
});

// Check cell info
const priceCell = sheet.getCell(1, 0);
if (priceCell) {
  console.log('Address:', priceCell.address);  // A2
  console.log('Value:', priceCell.value);      // 100
  console.log('Type:', priceCell.type);        // number
  console.log('Is empty:', priceCell.isEmpty); // false
}
```

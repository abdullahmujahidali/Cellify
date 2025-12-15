---
sidebar_position: 5
---

# Formulas

Cellify supports Excel formulas, allowing you to create dynamic spreadsheets with calculations.

## Setting Formulas

### Basic Formula

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Calculations');

// Set values
sheet.cell(0, 0).value = 10;
sheet.cell(0, 1).value = 20;

// Set formula (with or without leading '=')
sheet.cell(0, 2).setFormula('=A1+B1');
// or
sheet.cell(0, 2).setFormula('A1+B1');
```

### Formula Property

```typescript
// Get formula information
const cell = sheet.getCell(0, 2);
if (cell?.formula) {
  console.log('Formula:', cell.formula.formula);
  console.log('Cached result:', cell.formula.result);
}
```

## Common Formula Examples

### Arithmetic

```typescript
sheet.cell(0, 0).setFormula('A1+B1');      // Addition
sheet.cell(0, 1).setFormula('A1-B1');      // Subtraction
sheet.cell(0, 2).setFormula('A1*B1');      // Multiplication
sheet.cell(0, 3).setFormula('A1/B1');      // Division
sheet.cell(0, 4).setFormula('A1^2');       // Power
```

### SUM and Aggregates

```typescript
sheet.cell(5, 0).setFormula('SUM(A1:A5)');
sheet.cell(5, 1).setFormula('AVERAGE(B1:B5)');
sheet.cell(5, 2).setFormula('MIN(C1:C5)');
sheet.cell(5, 3).setFormula('MAX(D1:D5)');
sheet.cell(5, 4).setFormula('COUNT(E1:E5)');
```

### Conditional

```typescript
// IF statement
sheet.cell(0, 0).setFormula('IF(A1>100,"High","Low")');

// Nested IF
sheet.cell(0, 1).setFormula('IF(A1>100,"High",IF(A1>50,"Medium","Low"))');

// SUMIF
sheet.cell(0, 2).setFormula('SUMIF(A1:A10,">100",B1:B10)');

// COUNTIF
sheet.cell(0, 3).setFormula('COUNTIF(A1:A10,"Yes")');
```

### Lookup

```typescript
// VLOOKUP
sheet.cell(0, 0).setFormula('VLOOKUP(A1,Sheet2!A:B,2,FALSE)');

// INDEX/MATCH
sheet.cell(0, 1).setFormula('INDEX(B1:B10,MATCH(A1,A1:A10,0))');

// XLOOKUP (Excel 365+)
sheet.cell(0, 2).setFormula('XLOOKUP(A1,B:B,C:C)');
```

### Text

```typescript
sheet.cell(0, 0).setFormula('CONCATENATE(A1," ",B1)');
sheet.cell(0, 1).setFormula('UPPER(A1)');
sheet.cell(0, 2).setFormula('LOWER(A1)');
sheet.cell(0, 3).setFormula('LEN(A1)');
sheet.cell(0, 4).setFormula('LEFT(A1,5)');
sheet.cell(0, 5).setFormula('RIGHT(A1,3)');
sheet.cell(0, 6).setFormula('MID(A1,2,4)');
```

### Date and Time

```typescript
sheet.cell(0, 0).setFormula('TODAY()');
sheet.cell(0, 1).setFormula('NOW()');
sheet.cell(0, 2).setFormula('YEAR(A1)');
sheet.cell(0, 3).setFormula('MONTH(A1)');
sheet.cell(0, 4).setFormula('DAY(A1)');
sheet.cell(0, 5).setFormula('DATEDIF(A1,B1,"D")');
```

## Cell References

### Relative References

```typescript
// A1 - relative reference (changes when copied)
sheet.cell(0, 2).setFormula('A1+B1');
```

### Absolute References

```typescript
// $A$1 - absolute reference (stays fixed when copied)
sheet.cell(0, 2).setFormula('$A$1+$B$1');
```

### Mixed References

```typescript
// $A1 - column fixed, row relative
// A$1 - row fixed, column relative
sheet.cell(0, 2).setFormula('$A1+A$1');
```

### Cross-Sheet References

```typescript
// Reference another sheet
sheet.cell(0, 0).setFormula('Sheet2!A1');
sheet.cell(0, 1).setFormula("'Sheet with spaces'!A1");
sheet.cell(0, 2).setFormula('SUM(Sheet2!A1:A10)');
```

## Named Ranges in Formulas

Use defined names in formulas:

```typescript
// Define a named range
workbook.addDefinedName('SalesData', 'Sheet1!$A$1:$A$100');
workbook.addDefinedName('TaxRate', '0.08');

// Use in formulas
sheet.cell(0, 0).setFormula('SUM(SalesData)');
sheet.cell(0, 1).setFormula('A1*TaxRate');
```

## Array Formulas

```typescript
// Array formula (press Ctrl+Shift+Enter in Excel)
sheet.cell(0, 0).setFormula('SUM(A1:A10*B1:B10)');
```

## Clearing Formulas

```typescript
// Clear formula but keep the cell
cell.clearFormula();

// Or clear everything
cell.clear();
```

## Checking Cell Type

```typescript
const cell = sheet.getCell(0, 0);

// Check if cell has a formula
if (cell?.formula) {
  console.log('Has formula:', cell.formula.formula);
}

// Cell type will be 'formula' if it has a formula
if (cell?.type === 'formula') {
  console.log('This is a formula cell');
}
```

## Importing Formulas

When importing Excel files, formulas are preserved by default:

```typescript
import { xlsxBlobToWorkbook } from 'cellify';

const result = await xlsxBlobToWorkbook(file, {
  importFormulas: true, // Default: true
});

console.log('Formula cells:', result.stats.formulaCells);

// Access formula
const cell = result.workbook.sheets[0].getCell(0, 0);
if (cell?.formula) {
  console.log('Formula:', cell.formula.formula);
}
```

To skip formula import:

```typescript
const result = await xlsxBlobToWorkbook(file, {
  importFormulas: false, // Only import calculated values
});
```

## Limitations

Cellify preserves formulas but does not evaluate them. The actual calculation is performed by Excel when the file is opened.

**Supported:**
- Creating and preserving formulas
- All Excel formula syntax
- Named ranges in formulas
- Cross-sheet references
- Cached formula results from imported files

**Not Supported:**
- Runtime formula evaluation in JavaScript
- Formula validation/syntax checking
- Circular reference detection

## Best Practices

1. **Use absolute references for fixed values**:
   ```typescript
   sheet.cell(0, 0).setFormula('A1*$B$1'); // $B$1 is tax rate
   ```

2. **Define names for complex references**:
   ```typescript
   workbook.addDefinedName('PriceList', 'Products!$A$1:$B$100');
   sheet.cell(0, 0).setFormula('VLOOKUP(A1,PriceList,2,FALSE)');
   ```

3. **Handle errors in formulas**:
   ```typescript
   sheet.cell(0, 0).setFormula('IFERROR(A1/B1,0)');
   ```

4. **Use structured table references when applicable**:
   ```typescript
   sheet.cell(0, 0).setFormula('SUM(Table1[Sales])');
   ```

---
sidebar_position: 4
---

# Merging Cells

Cellify supports merging cells to create headers, titles, and complex layouts in your spreadsheets.

## Basic Merge

### Using A1 Notation

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Report');

// Set value in the top-left cell
sheet.cell('A1').value = 'Quarterly Report';

// Merge A1 through D1
sheet.mergeCells('A1:D1');
```

### Using Row/Column Indices

```typescript
// Merge using range definition
sheet.mergeCells({
  startRow: 0,
  startCol: 0,
  endRow: 0,
  endCol: 3, // A1:D1
});
```

## Merge Rules

1. **Value in top-left cell**: Only the top-left cell of a merged range retains its value
2. **No overlapping merges**: Cannot merge cells that are already part of another merge
3. **Style applies to all**: Styling the master cell affects the entire merged area

```typescript
// Set value BEFORE merging
sheet.cell(0, 0).value = 'Title';

// Then merge
sheet.mergeCells('A1:D1');

// Style the master cell
sheet.cell(0, 0).applyStyle({
  font: { bold: true, size: 16 },
  alignment: { horizontal: 'center', vertical: 'middle' },
});
```

## Unmerging Cells

```typescript
// Unmerge using the same range
sheet.unmergeCells('A1:D1');

// Or using range definition
sheet.unmergeCells({
  startRow: 0,
  startCol: 0,
  endRow: 0,
  endCol: 3,
});
```

## Getting Merge Information

### List All Merges

```typescript
// Get all merged ranges in the sheet
const merges = sheet.merges;

merges.forEach(merge => {
  console.log(`Merge: ${merge.startRow},${merge.startCol} to ${merge.endRow},${merge.endCol}`);
});
```

### Check if Cell is Merged

```typescript
const cell = sheet.getCell(0, 0);

// Is this the master cell of a merge?
if (cell?.isMergeMaster) {
  console.log('This is the top-left cell of a merge');
  console.log('Merge range:', cell.merge);
}

// Is this cell part of a merge (but not the master)?
if (cell?.isMergedSlave) {
  console.log('This cell is merged into:', cell.mergedInto);
}

// Is this cell part of any merge?
if (cell?.isMerged) {
  console.log('This cell is part of a merged range');
}
```

## Common Patterns

### Report Title

```typescript
// Full-width title
sheet.cell(0, 0).value = 'Annual Sales Report 2024';
sheet.mergeCells('A1:F1');
sheet.cell(0, 0).applyStyle({
  font: { bold: true, size: 18 },
  alignment: { horizontal: 'center' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#1F2937' },
});
```

### Category Headers

```typescript
// Main header spanning multiple columns
sheet.cell(0, 0).value = 'Q1 Sales';
sheet.mergeCells('A1:C1');

sheet.cell(0, 3).value = 'Q2 Sales';
sheet.mergeCells('D1:F1');

// Sub-headers
['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'].forEach((month, col) => {
  sheet.cell(1, col).value = month;
});
```

### Vertical Merge

```typescript
// Row labels spanning multiple rows
sheet.cell(0, 0).value = 'North Region';
sheet.mergeCells('A1:A3');
sheet.cell(0, 0).applyStyle({
  alignment: { vertical: 'middle', horizontal: 'center' },
  textRotation: 90, // Optional: rotate text
});
```

### Complex Layout

```typescript
const workbook = new Workbook();
const sheet = workbook.addSheet('Complex');

// Title (row 0)
sheet.cell(0, 0).value = 'Department Summary';
sheet.mergeCells('A1:E1');
sheet.cell(0, 0).applyStyle({
  font: { bold: true, size: 16 },
  alignment: { horizontal: 'center' },
});

// Department sections
const departments = ['Engineering', 'Marketing', 'Sales'];
let currentRow = 1;

departments.forEach(dept => {
  // Department header
  sheet.cell(currentRow, 0).value = dept;
  sheet.mergeCells({
    startRow: currentRow,
    startCol: 0,
    endRow: currentRow,
    endCol: 4,
  });
  sheet.cell(currentRow, 0).applyStyle({
    font: { bold: true },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#E5E7EB' },
  });

  currentRow++;

  // Data rows
  for (let i = 0; i < 3; i++) {
    sheet.cell(currentRow, 0).value = `Employee ${i + 1}`;
    sheet.cell(currentRow, 1).value = 'Role';
    sheet.cell(currentRow, 2).value = 50000 + Math.random() * 50000;
    currentRow++;
  }
});
```

## Error Handling

### Overlapping Merges

```typescript
sheet.mergeCells('A1:C1');

try {
  // This will throw an error - overlaps with A1:C1
  sheet.mergeCells('B1:D1');
} catch (error) {
  console.error('Cannot merge: overlaps with existing merge');
}
```

### Unmerge Non-existent Range

```typescript
try {
  // This will throw if no merge exists at this range
  sheet.unmergeCells('E1:F1');
} catch (error) {
  console.error('No merge found at specified range');
}
```

## Best Practices

1. **Set value before merging**: Always set the cell value before calling `mergeCells()`

2. **Style the master cell**: Apply styles to the top-left cell of the merge

3. **Use consistent ranges**: When unmerging, use the exact same range used for merging

4. **Center content**: Merged cells often look better with centered alignment:
   ```typescript
   sheet.cell(0, 0).applyStyle({
     alignment: { horizontal: 'center', vertical: 'middle' },
   });
   ```

5. **Consider column widths**: Adjust column widths to accommodate merged content:
   ```typescript
   sheet.setColumnWidth(0, 20);
   sheet.setColumnWidth(1, 15);
   ```

## Importing Merged Cells

When importing Excel files, merged cells are preserved:

```typescript
import { xlsxBlobToWorkbook } from 'cellify';

const result = await xlsxBlobToWorkbook(file);
const sheet = result.workbook.sheets[0];

// Check import stats
console.log('Merged ranges:', result.stats.mergedRanges);

// Iterate merges
sheet.merges.forEach(merge => {
  const masterCell = sheet.getCell(merge.startRow, merge.startCol);
  console.log(`Merged value: ${masterCell?.value}`);
});
```

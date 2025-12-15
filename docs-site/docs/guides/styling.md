---
sidebar_position: 3
---

# Styling Cells

Cellify provides comprehensive styling options for cells including fonts, fills, borders, alignment, and number formats.

## Basic Styling

### Inline Style

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Styled');

// Apply style when setting value
sheet.cell(0, 0).value = 'Header';
sheet.cell(0, 0).style = {
  font: { bold: true, size: 14 },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#4F46E5' },
};
```

### Using applyStyle()

The `applyStyle()` method merges styles instead of replacing:

```typescript
// Start with bold
sheet.cell(0, 0).applyStyle({ font: { bold: true } });

// Add color (keeps bold)
sheet.cell(0, 0).applyStyle({ font: { color: '#FF0000' } });

// Result: bold red text
```

## Font Styling

```typescript
sheet.cell(0, 0).style = {
  font: {
    name: 'Arial',           // Font family
    size: 12,                // Size in points
    color: '#000000',        // Hex color
    bold: true,
    italic: true,
    underline: 'single',     // 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting'
    strikethrough: true,
    superscript: true,       // or subscript: true
  },
};
```

### Common Font Styles

```typescript
// Bold header
sheet.cell(0, 0).applyStyle({
  font: { bold: true, size: 14 },
});

// Red warning text
sheet.cell(0, 0).applyStyle({
  font: { color: '#DC2626', bold: true },
});

// Italic note
sheet.cell(0, 0).applyStyle({
  font: { italic: true, color: '#6B7280' },
});
```

## Fill Styling

### Solid Fill

```typescript
sheet.cell(0, 0).style = {
  fill: {
    type: 'pattern',
    pattern: 'solid',
    foregroundColor: '#059669', // Green background
  },
};
```

### Pattern Fills

```typescript
sheet.cell(0, 0).style = {
  fill: {
    type: 'pattern',
    pattern: 'lightGray', // Pattern type
    foregroundColor: '#000000',
    backgroundColor: '#FFFFFF',
  },
};
```

Available patterns:
- `none`, `solid`
- `darkGray`, `mediumGray`, `lightGray`, `gray125`, `gray0625`
- `darkHorizontal`, `darkVertical`, `darkDown`, `darkUp`, `darkGrid`, `darkTrellis`
- `lightHorizontal`, `lightVertical`, `lightDown`, `lightUp`, `lightGrid`, `lightTrellis`

### Gradient Fills

```typescript
// Linear gradient
sheet.cell(0, 0).style = {
  fill: {
    type: 'gradient',
    gradientType: 'linear',
    degree: 90, // Rotation angle (0-360)
    stops: [
      { position: 0, color: '#4F46E5' },
      { position: 1, color: '#06B6D4' },
    ],
  },
};

// Path gradient
sheet.cell(0, 0).style = {
  fill: {
    type: 'gradient',
    gradientType: 'path',
    left: 0.5,
    right: 0.5,
    top: 0.5,
    bottom: 0.5,
    stops: [
      { position: 0, color: '#FFFFFF' },
      { position: 1, color: '#4F46E5' },
    ],
  },
};
```

## Border Styling

### Single Border

```typescript
sheet.cell(0, 0).style = {
  borders: {
    bottom: { style: 'thin', color: '#000000' },
  },
};
```

### All Borders

```typescript
sheet.cell(0, 0).style = {
  borders: {
    top: { style: 'thin', color: '#000000' },
    right: { style: 'thin', color: '#000000' },
    bottom: { style: 'thin', color: '#000000' },
    left: { style: 'thin', color: '#000000' },
  },
};
```

### Border Styles

Available border styles:
- `none`, `thin`, `medium`, `thick`
- `dashed`, `dotted`, `double`, `hair`
- `mediumDashed`, `dashDot`, `mediumDashDot`
- `dashDotDot`, `mediumDashDotDot`, `slantDashDot`

### Diagonal Borders

```typescript
sheet.cell(0, 0).style = {
  borders: {
    diagonal: { style: 'thin', color: '#FF0000' },
    diagonalUp: true,   // Bottom-left to top-right
    diagonalDown: true, // Top-left to bottom-right
  },
};
```

## Alignment

```typescript
sheet.cell(0, 0).style = {
  alignment: {
    horizontal: 'center',    // 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'
    vertical: 'middle',      // 'top', 'middle', 'bottom', 'justify', 'distributed'
    wrapText: true,          // Wrap long text
    shrinkToFit: true,       // Shrink text to fit cell
    indent: 2,               // Indentation level (0-255)
    textRotation: 45,        // Rotation angle (-90 to 90, or 255 for vertical)
    readingOrder: 'leftToRight', // 'contextDependent', 'leftToRight', 'rightToLeft'
  },
};
```

### Common Alignments

```typescript
// Center both horizontally and vertically
sheet.cell(0, 0).applyStyle({
  alignment: { horizontal: 'center', vertical: 'middle' },
});

// Right-aligned numbers
sheet.cell(0, 0).applyStyle({
  alignment: { horizontal: 'right' },
});

// Wrapped text
sheet.cell(0, 0).applyStyle({
  alignment: { wrapText: true },
});
```

## Number Formats

```typescript
sheet.cell(0, 0).style = {
  numberFormat: {
    formatCode: '$#,##0.00',
    category: 'currency',
  },
};
```

### Common Format Codes

| Category | Format Code | Example Output |
|----------|-------------|----------------|
| General | `General` | 1234 |
| Number | `0.00` | 1234.00 |
| Number | `#,##0` | 1,234 |
| Currency | `$#,##0.00` | $1,234.00 |
| Percentage | `0%` | 50% |
| Percentage | `0.00%` | 50.00% |
| Date | `yyyy-mm-dd` | 2024-01-15 |
| Date | `mm/dd/yyyy` | 01/15/2024 |
| Date | `d-mmm-yy` | 15-Jan-24 |
| Time | `h:mm AM/PM` | 2:30 PM |
| Time | `h:mm:ss` | 14:30:00 |
| Scientific | `0.00E+00` | 1.23E+03 |
| Fraction | `# ?/?` | 1 1/2 |
| Text | `@` | (as text) |

### Custom Number Formats

```typescript
// Accounting format
sheet.cell(0, 0).applyStyle({
  numberFormat: { formatCode: '_($* #,##0.00_)' },
});

// Negative numbers in red
sheet.cell(0, 0).applyStyle({
  numberFormat: { formatCode: '#,##0.00;[Red]-#,##0.00' },
});

// Phone number
sheet.cell(0, 0).applyStyle({
  numberFormat: { formatCode: '(###) ###-####' },
});
```

## Cell Protection

```typescript
sheet.cell(0, 0).style = {
  protection: {
    locked: true,   // Prevent editing when sheet is protected
    hidden: true,   // Hide formula in formula bar
  },
};
```

## Applying Styles to Ranges

```typescript
// Apply style to entire range
sheet.applyStyle('A1:D1', {
  font: { bold: true },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#F3F4F6' },
});

// Apply style with range definition
sheet.applyStyle(
  { startRow: 0, startCol: 0, endRow: 0, endCol: 3 },
  { font: { bold: true } }
);
```

## Named Styles

Create reusable styles at the workbook level:

```typescript
// Define named styles
workbook.addNamedStyle('header', {
  font: { bold: true, size: 14, color: '#FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#4F46E5' },
  alignment: { horizontal: 'center' },
});

workbook.addNamedStyle('currency', {
  numberFormat: { formatCode: '$#,##0.00' },
  alignment: { horizontal: 'right' },
});

// Apply named styles
const headerStyle = workbook.getNamedStyle('header');
if (headerStyle) {
  sheet.cell(0, 0).style = headerStyle.style;
}
```

## Complete Example

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Report');

// Title row
sheet.cell(0, 0).value = 'Sales Report Q4 2024';
sheet.cell(0, 0).applyStyle({
  font: { bold: true, size: 18, color: '#1F2937' },
});
sheet.mergeCells('A1:D1');

// Header row
const headers = ['Product', 'Units', 'Price', 'Total'];
headers.forEach((header, col) => {
  sheet.cell(1, col).value = header;
  sheet.cell(1, col).applyStyle({
    font: { bold: true, color: '#FFFFFF' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#059669' },
    alignment: { horizontal: 'center' },
    borders: {
      bottom: { style: 'medium', color: '#047857' },
    },
  });
});

// Data rows
const data = [
  ['Widget A', 150, 29.99, 4498.5],
  ['Widget B', 200, 19.99, 3998.0],
  ['Widget C', 75, 49.99, 3749.25],
];

data.forEach((row, rowIndex) => {
  row.forEach((value, colIndex) => {
    const cell = sheet.cell(rowIndex + 2, colIndex);
    cell.value = value;

    // Apply number format for currency columns
    if (colIndex >= 2) {
      cell.applyStyle({
        numberFormat: { formatCode: '$#,##0.00' },
        alignment: { horizontal: 'right' },
      });
    }
  });
});

// Total row
sheet.cell(5, 0).value = 'Total';
sheet.cell(5, 0).applyStyle({ font: { bold: true } });
sheet.cell(5, 3).value = 12245.75;
sheet.cell(5, 3).applyStyle({
  font: { bold: true },
  numberFormat: { formatCode: '$#,##0.00' },
  borders: { top: { style: 'double', color: '#000000' } },
});

const blob = workbookToXlsxBlob(workbook);
```

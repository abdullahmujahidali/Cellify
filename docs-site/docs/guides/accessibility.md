---
sidebar_position: 6
---

# Accessibility

Cellify includes features to help create accessible spreadsheets that work well with screen readers and assistive technologies.

## Document Properties

Set descriptive document properties to help users understand the spreadsheet:

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();

workbook.setProperties({
  title: 'Q4 2024 Sales Report',
  subject: 'Quarterly sales data by region and product',
  author: 'Finance Team',
  keywords: ['sales', 'quarterly', 'report', '2024'],
  comments: 'Contains sales figures for all regions',
});
```

## Sheet Names

Use descriptive, meaningful sheet names:

```typescript
// Good: Descriptive names
workbook.addSheet('Sales Summary');
workbook.addSheet('Regional Breakdown');
workbook.addSheet('Product Details');

// Avoid: Generic names
// workbook.addSheet('Sheet1');
// workbook.addSheet('Data');
```

## Table Headers

Always include clear headers for data tables:

```typescript
const sheet = workbook.addSheet('Sales Data');

// Define clear headers
const headers = ['Product Name', 'Category', 'Units Sold', 'Revenue', 'Profit Margin'];
headers.forEach((header, col) => {
  sheet.cell(0, col).value = header;
  sheet.cell(0, col).applyStyle({
    font: { bold: true },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#E5E7EB' },
  });
});
```

## Freeze Panes for Navigation

Freeze header rows to keep them visible while scrolling:

```typescript
// Freeze top row (headers)
sheet.freeze(1, 0);

// Freeze top row and first column
sheet.freeze(1, 1);
```

## Cell Comments for Context

Add comments to provide additional context:

```typescript
// Add explanatory comment
sheet.cell(0, 3).value = 'Revenue';
sheet.cell(0, 3).setComment(
  'Total revenue in USD before taxes and discounts',
  'Finance Team'
);

// Highlight cells with comments
sheet.cell(1, 3).value = 150000;
sheet.cell(1, 3).setComment(
  'Includes one-time promotional sales of $25,000'
);
```

## Color Contrast

Ensure sufficient color contrast for readability:

```typescript
// Good contrast: Dark text on light background
sheet.cell(0, 0).applyStyle({
  font: { color: '#1F2937' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#F3F4F6' },
});

// Good contrast: Light text on dark background
sheet.cell(0, 1).applyStyle({
  font: { color: '#FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#1F2937' },
});

// Avoid: Low contrast combinations
// sheet.cell(0, 2).applyStyle({
//   font: { color: '#9CA3AF' },
//   fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#E5E7EB' },
// });
```

## Don't Rely on Color Alone

Use multiple indicators, not just color, to convey information:

```typescript
// Good: Color + text indicator
sheet.cell(0, 0).value = 'Status: Approved';
sheet.cell(0, 0).applyStyle({
  font: { color: '#059669' }, // Green
});

sheet.cell(1, 0).value = 'Status: Rejected';
sheet.cell(1, 0).applyStyle({
  font: { color: '#DC2626' }, // Red
});

// Also good: Color + icon
sheet.cell(2, 0).value = '✓ Complete';
sheet.cell(3, 0).value = '✗ Incomplete';
```

## Clear Number Formats

Use appropriate number formats for clarity:

```typescript
// Currency with symbol
sheet.cell(0, 0).value = 1234.56;
sheet.cell(0, 0).applyStyle({
  numberFormat: { formatCode: '$#,##0.00' },
});

// Percentages clearly marked
sheet.cell(0, 1).value = 0.156;
sheet.cell(0, 1).applyStyle({
  numberFormat: { formatCode: '0.0%' },
});

// Dates in unambiguous format
sheet.cell(0, 2).value = new Date(2024, 0, 15);
sheet.cell(0, 2).applyStyle({
  numberFormat: { formatCode: 'yyyy-mm-dd' },
});
```

## Logical Reading Order

Structure data in a logical reading order (left to right, top to bottom):

```typescript
const sheet = workbook.addSheet('Report');

// Title
sheet.cell(0, 0).value = 'Monthly Sales Report';
sheet.mergeCells('A1:D1');

// Headers (row 1)
['Date', 'Product', 'Quantity', 'Amount'].forEach((h, i) => {
  sheet.cell(1, i).value = h;
});

// Data (rows 2+)
// ...

// Summary at the end (not embedded in the middle)
sheet.cell(10, 0).value = 'Total';
sheet.cell(10, 3).value = totalAmount;
```

## Avoid Empty Cells in Headers

Don't leave gaps in header rows:

```typescript
// Good: All headers filled
['ID', 'Name', 'Department', 'Start Date'].forEach((h, i) => {
  sheet.cell(0, i).value = h;
});

// Avoid: Empty cells in header row
// This can confuse screen readers
```

## Use Merged Cells Sparingly

Merged cells can be difficult for screen readers to navigate:

```typescript
// Use merges only for clear visual grouping
sheet.cell(0, 0).value = 'Q1 2024';
sheet.mergeCells('A1:C1'); // Title spanning 3 columns

// Avoid excessive or complex merge patterns
```

## Alternative Text for Complex Data

For complex data representations, consider adding a summary sheet:

```typescript
const summarySheet = workbook.addSheet('Summary');
summarySheet.cell(0, 0).value = 'About This Workbook';
summarySheet.cell(1, 0).value = 'This workbook contains quarterly sales data organized as follows:';
summarySheet.cell(2, 0).value = '- Sheet "Sales": Raw sales transactions';
summarySheet.cell(3, 0).value = '- Sheet "By Region": Sales grouped by geographic region';
summarySheet.cell(4, 0).value = '- Sheet "By Product": Sales grouped by product category';
```

## Auto Filter for Data Tables

Enable filtering for large data sets:

```typescript
// Add data
const headers = ['Name', 'Department', 'Salary', 'Start Date'];
headers.forEach((h, i) => sheet.cell(0, i).value = h);

// ... add data rows ...

// Enable auto filter
sheet.setAutoFilter('A1:D100');
```

## Complete Accessible Spreadsheet Example

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

const workbook = new Workbook();

// Set document properties
workbook.setProperties({
  title: 'Employee Directory',
  subject: 'Company employee contact information',
  author: 'HR Department',
  keywords: ['employees', 'directory', 'contacts'],
});

// Create main data sheet
const sheet = workbook.addSheet('Employee List');

// Headers with clear labels
const headers = ['Employee ID', 'Full Name', 'Department', 'Email', 'Phone', 'Start Date'];
headers.forEach((header, col) => {
  const cell = sheet.cell(0, col);
  cell.value = header;
  cell.applyStyle({
    font: { bold: true, color: '#FFFFFF' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#1F2937' },
    alignment: { horizontal: 'center' },
  });
});

// Set appropriate column widths
[12, 25, 20, 30, 15, 12].forEach((width, col) => {
  sheet.setColumnWidth(col, width);
});

// Freeze header row
sheet.freeze(1, 0);

// Add data with proper formatting
const employees = [
  { id: 'E001', name: 'Alice Smith', dept: 'Engineering', email: 'alice@company.com', phone: '555-0101', start: new Date(2020, 5, 15) },
  { id: 'E002', name: 'Bob Johnson', dept: 'Marketing', email: 'bob@company.com', phone: '555-0102', start: new Date(2019, 2, 1) },
];

employees.forEach((emp, row) => {
  const r = row + 1;
  sheet.cell(r, 0).value = emp.id;
  sheet.cell(r, 1).value = emp.name;
  sheet.cell(r, 2).value = emp.dept;
  sheet.cell(r, 3).value = emp.email;
  sheet.cell(r, 4).value = emp.phone;
  sheet.cell(r, 5).value = emp.start;
  sheet.cell(r, 5).applyStyle({
    numberFormat: { formatCode: 'yyyy-mm-dd' },
  });
});

// Enable filtering
sheet.setAutoFilter('A1:F' + (employees.length + 1));

// Export
const blob = workbookToXlsxBlob(workbook);
```

## Testing Accessibility

1. **Screen reader testing**: Open the exported file in Excel and test with a screen reader (NVDA, JAWS, VoiceOver)

2. **Keyboard navigation**: Ensure all data can be accessed using only keyboard navigation

3. **High contrast mode**: Test the spreadsheet in Windows High Contrast mode

4. **Zoom testing**: Verify readability at 200% zoom

5. **Color blindness simulation**: Use tools to check color combinations work for users with color vision deficiencies

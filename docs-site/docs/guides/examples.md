---
sidebar_position: 7
---

# Real-World Examples

Practical examples for common spreadsheet tasks.

## Invoice Generator

Create a professional invoice with styling, formulas, and merged cells:

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

function createInvoice(invoiceData: {
  invoiceNumber: string;
  date: Date;
  customer: { name: string; address: string; email: string };
  items: Array<{ description: string; quantity: number; price: number }>;
}) {
  const workbook = new Workbook();
  workbook.title = `Invoice ${invoiceData.invoiceNumber}`;

  const sheet = workbook.addSheet('Invoice');

  // Company header
  sheet.cell('A1').value = 'ACME Corporation';
  sheet.cell('A1').style = {
    font: { bold: true, size: 20, color: '#059669' },
  };
  sheet.mergeCells('A1:D1');

  sheet.cell('A2').value = '123 Business Street, City, State 12345';
  sheet.cell('A2').style = { font: { color: '#6B7280' } };
  sheet.mergeCells('A2:D2');

  // Invoice details
  sheet.cell('A4').value = `Invoice #: ${invoiceData.invoiceNumber}`;
  sheet.cell('A4').style = { font: { bold: true } };

  sheet.cell('A5').value = `Date: ${invoiceData.date.toLocaleDateString()}`;

  // Customer info
  sheet.cell('A7').value = 'Bill To:';
  sheet.cell('A7').style = { font: { bold: true, color: '#059669' } };
  sheet.cell('A8').value = invoiceData.customer.name;
  sheet.cell('A9').value = invoiceData.customer.address;
  sheet.cell('A10').value = invoiceData.customer.email;

  // Items table header
  const headerRow = 12;
  const headers = ['Description', 'Quantity', 'Unit Price', 'Total'];

  headers.forEach((header, col) => {
    sheet.cell(headerRow, col).value = header;
    sheet.cell(headerRow, col).style = {
      font: { bold: true, color: '#FFFFFF' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#059669' },
      alignment: { horizontal: 'center' },
    };
  });

  // Items
  let currentRow = headerRow + 1;
  invoiceData.items.forEach((item, index) => {
    sheet.cell(currentRow, 0).value = item.description;
    sheet.cell(currentRow, 1).value = item.quantity;
    sheet.cell(currentRow, 1).style = { alignment: { horizontal: 'center' } };
    sheet.cell(currentRow, 2).value = item.price;
    sheet.cell(currentRow, 2).style = { numberFormat: { formatCode: '$#,##0.00' } };
    sheet.cell(currentRow, 3).setFormula(`B${currentRow + 1}*C${currentRow + 1}`);
    sheet.cell(currentRow, 3).style = { numberFormat: { formatCode: '$#,##0.00' } };

    // Alternate row colors
    if (index % 2 === 0) {
      for (let c = 0; c < 4; c++) {
        sheet.cell(currentRow, c).style = {
          ...sheet.cell(currentRow, c).style,
          fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#F3F4F6' },
        };
      }
    }
    currentRow++;
  });

  // Total row
  const totalRow = currentRow + 1;
  sheet.cell(totalRow, 2).value = 'Total:';
  sheet.cell(totalRow, 2).style = { font: { bold: true }, alignment: { horizontal: 'right' } };
  sheet.cell(totalRow, 3).setFormula(`SUM(D${headerRow + 2}:D${currentRow})`);
  sheet.cell(totalRow, 3).style = {
    font: { bold: true, size: 14 },
    numberFormat: { formatCode: '$#,##0.00' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#D1FAE5' },
  };

  // Column widths
  sheet.setColumnWidth(0, 40);  // Description
  sheet.setColumnWidth(1, 12);  // Quantity
  sheet.setColumnWidth(2, 15);  // Unit Price
  sheet.setColumnWidth(3, 15);  // Total

  return workbookToXlsxBlob(workbook);
}

// Usage
const invoice = createInvoice({
  invoiceNumber: 'INV-2024-001',
  date: new Date(),
  customer: {
    name: 'John Doe',
    address: '456 Customer Lane, Town, State 67890',
    email: 'john@example.com',
  },
  items: [
    { description: 'Web Development Services', quantity: 40, price: 150 },
    { description: 'UI/UX Design', quantity: 20, price: 125 },
    { description: 'Server Hosting (Monthly)', quantity: 1, price: 99 },
  ],
});
```

## Sales Report with Charts Data

Create a sales report with data organized for Excel charts:

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

function createSalesReport(salesData: {
  months: string[];
  products: Array<{ name: string; sales: number[] }>;
}) {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Sales Report');

  // Title
  sheet.cell('A1').value = 'Monthly Sales Report 2024';
  sheet.cell('A1').style = {
    font: { bold: true, size: 18 },
  };
  sheet.mergeCells('A1:G1');

  // Headers
  const headerRow = 3;
  sheet.cell(headerRow, 0).value = 'Product';

  salesData.months.forEach((month, col) => {
    sheet.cell(headerRow, col + 1).value = month;
  });
  sheet.cell(headerRow, salesData.months.length + 1).value = 'Total';
  sheet.cell(headerRow, salesData.months.length + 2).value = 'Average';

  // Style headers
  for (let c = 0; c <= salesData.months.length + 2; c++) {
    sheet.cell(headerRow, c).style = {
      font: { bold: true, color: '#FFFFFF' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#1F2937' },
      alignment: { horizontal: 'center' },
    };
  }

  // Data rows
  salesData.products.forEach((product, rowIndex) => {
    const row = headerRow + 1 + rowIndex;
    sheet.cell(row, 0).value = product.name;

    product.sales.forEach((sale, col) => {
      sheet.cell(row, col + 1).value = sale;
      sheet.cell(row, col + 1).style = {
        numberFormat: { formatCode: '$#,##0' },
      };
    });

    // Total formula
    const lastDataCol = String.fromCharCode(65 + salesData.months.length);
    sheet.cell(row, salesData.months.length + 1).setFormula(
      `SUM(B${row + 1}:${lastDataCol}${row + 1})`
    );
    sheet.cell(row, salesData.months.length + 1).style = {
      font: { bold: true },
      numberFormat: { formatCode: '$#,##0' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#ECFDF5' },
    };

    // Average formula
    sheet.cell(row, salesData.months.length + 2).setFormula(
      `AVERAGE(B${row + 1}:${lastDataCol}${row + 1})`
    );
    sheet.cell(row, salesData.months.length + 2).style = {
      numberFormat: { formatCode: '$#,##0.00' },
    };
  });

  // Column totals row
  const totalRow = headerRow + 1 + salesData.products.length + 1;
  sheet.cell(totalRow, 0).value = 'Monthly Total';
  sheet.cell(totalRow, 0).style = { font: { bold: true } };

  for (let c = 1; c <= salesData.months.length; c++) {
    const col = String.fromCharCode(65 + c);
    sheet.cell(totalRow, c).setFormula(
      `SUM(${col}${headerRow + 2}:${col}${totalRow - 1})`
    );
    sheet.cell(totalRow, c).style = {
      font: { bold: true },
      numberFormat: { formatCode: '$#,##0' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FEF3C7' },
    };
  }

  // Freeze header rows
  sheet.freeze(headerRow + 1, 1);

  // Auto filter
  const lastCol = String.fromCharCode(65 + salesData.months.length + 2);
  sheet.setAutoFilter(`A${headerRow + 1}:${lastCol}${totalRow - 1}`);

  // Column widths
  sheet.setColumnWidth(0, 20);
  for (let c = 1; c <= salesData.months.length + 2; c++) {
    sheet.setColumnWidth(c, 12);
  }

  return workbookToXlsxBlob(workbook);
}

// Usage
const report = createSalesReport({
  months: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
  products: [
    { name: 'Product A', sales: [12000, 15000, 13500, 18000, 16000, 21000] },
    { name: 'Product B', sales: [8000, 9500, 11000, 10500, 12000, 14000] },
    { name: 'Product C', sales: [5000, 6000, 5500, 7000, 8000, 9500] },
  ],
});
```

## Data Import and Transform

Import an Excel file, transform the data, and export a summary:

```typescript
import { xlsxBlobToWorkbook, Workbook, workbookToXlsxBlob } from 'cellify';

async function analyzeAndSummarize(file: Blob) {
  // Import the file
  const { workbook: sourceWorkbook, stats } = await xlsxBlobToWorkbook(file);

  console.log(`Imported ${stats.totalCells} cells from ${stats.sheetCount} sheets`);

  // Create summary workbook
  const summaryWorkbook = new Workbook();
  summaryWorkbook.title = 'Data Summary';

  // Summary sheet
  const summarySheet = summaryWorkbook.addSheet('Summary');

  summarySheet.cell('A1').value = 'Data Analysis Summary';
  summarySheet.cell('A1').style = { font: { bold: true, size: 16 } };
  summarySheet.mergeCells('A1:D1');

  // Sheet statistics
  summarySheet.cell('A3').value = 'Sheet Name';
  summarySheet.cell('B3').value = 'Rows';
  summarySheet.cell('C3').value = 'Columns';
  summarySheet.cell('D3').value = 'Cells';

  for (let c = 0; c < 4; c++) {
    summarySheet.cell(2, c).style = {
      font: { bold: true },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#E5E7EB' },
    };
  }

  sourceWorkbook.sheets.forEach((sheet, index) => {
    const dims = sheet.dimensions;
    const row = index + 3;

    summarySheet.cell(row, 0).value = sheet.name;

    if (dims) {
      const rows = dims.endRow - dims.startRow + 1;
      const cols = dims.endCol - dims.startCol + 1;
      summarySheet.cell(row, 1).value = rows;
      summarySheet.cell(row, 2).value = cols;
      summarySheet.cell(row, 3).value = rows * cols;
    } else {
      summarySheet.cell(row, 1).value = 0;
      summarySheet.cell(row, 2).value = 0;
      summarySheet.cell(row, 3).value = 0;
    }
  });

  // Copy first sheet data
  const dataSheet = summaryWorkbook.addSheet('Data Copy');
  const sourceSheet = sourceWorkbook.sheets[0];
  const dims = sourceSheet.dimensions;

  if (dims) {
    for (let r = dims.startRow; r <= dims.endRow; r++) {
      for (let c = dims.startCol; c <= dims.endCol; c++) {
        const sourceCell = sourceSheet.getCell(r, c);
        if (sourceCell) {
          const targetCell = dataSheet.cell(r, c);
          targetCell.value = sourceCell.value;

          // Copy formula if exists
          if (sourceCell.formula) {
            targetCell.setFormula(sourceCell.formula.formula);
          }

          // Copy style if exists
          if (sourceCell.style) {
            targetCell.style = sourceCell.style;
          }
        }
      }
    }
  }

  return workbookToXlsxBlob(summaryWorkbook);
}
```

## Employee Directory

Create an employee directory with search-friendly formatting:

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

interface Employee {
  id: number;
  name: string;
  department: string;
  email: string;
  phone: string;
  startDate: Date;
  salary: number;
}

function createEmployeeDirectory(employees: Employee[]) {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Employees');

  // Title
  sheet.cell('A1').value = 'Employee Directory';
  sheet.cell('A1').style = {
    font: { bold: true, size: 18, color: '#1F2937' },
  };
  sheet.mergeCells('A1:G1');

  sheet.cell('A2').value = `Last Updated: ${new Date().toLocaleDateString()}`;
  sheet.cell('A2').style = { font: { italic: true, color: '#6B7280' } };

  // Headers
  const headers = ['ID', 'Name', 'Department', 'Email', 'Phone', 'Start Date', 'Salary'];
  const headerRow = 4;

  headers.forEach((header, col) => {
    sheet.cell(headerRow, col).value = header;
    sheet.cell(headerRow, col).style = {
      font: { bold: true, color: '#FFFFFF' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#059669' },
      alignment: { horizontal: 'center' },
      borders: {
        bottom: { style: 'medium', color: '#047857' },
      },
    };
  });

  // Employee data
  employees.forEach((emp, index) => {
    const row = headerRow + 1 + index;
    const isEven = index % 2 === 0;
    const bgColor = isEven ? '#F9FAFB' : '#FFFFFF';

    sheet.cell(row, 0).value = emp.id;
    sheet.cell(row, 1).value = emp.name;
    sheet.cell(row, 2).value = emp.department;
    sheet.cell(row, 3).value = emp.email;
    sheet.cell(row, 4).value = emp.phone;
    sheet.cell(row, 5).value = emp.startDate;
    sheet.cell(row, 6).value = emp.salary;

    // Apply styles
    for (let c = 0; c < 7; c++) {
      sheet.cell(row, c).style = {
        fill: { type: 'pattern', pattern: 'solid', foregroundColor: bgColor },
      };
    }

    // Format date
    sheet.cell(row, 5).style = {
      ...sheet.cell(row, 5).style,
      numberFormat: { formatCode: 'YYYY-MM-DD' },
    };

    // Format salary
    sheet.cell(row, 6).style = {
      ...sheet.cell(row, 6).style,
      numberFormat: { formatCode: '$#,##0' },
    };
  });

  // Column widths
  sheet.setColumnWidth(0, 8);   // ID
  sheet.setColumnWidth(1, 25);  // Name
  sheet.setColumnWidth(2, 20);  // Department
  sheet.setColumnWidth(3, 30);  // Email
  sheet.setColumnWidth(4, 15);  // Phone
  sheet.setColumnWidth(5, 12);  // Start Date
  sheet.setColumnWidth(6, 12);  // Salary

  // Freeze header
  sheet.freeze(headerRow + 1, 0);

  // Enable filtering
  const lastRow = headerRow + employees.length;
  sheet.setAutoFilter(`A${headerRow + 1}:G${lastRow + 1}`);

  return workbookToXlsxBlob(workbook);
}

// Usage
const directory = createEmployeeDirectory([
  {
    id: 1,
    name: 'Alice Johnson',
    department: 'Engineering',
    email: 'alice@company.com',
    phone: '555-0101',
    startDate: new Date('2020-03-15'),
    salary: 95000,
  },
  {
    id: 2,
    name: 'Bob Smith',
    department: 'Marketing',
    email: 'bob@company.com',
    phone: '555-0102',
    startDate: new Date('2019-07-22'),
    salary: 75000,
  },
  // ... more employees
]);
```

## Budget Tracker

Create a monthly budget tracker with categories and totals:

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

function createBudgetTracker(year: number) {
  const workbook = new Workbook();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  const categories = [
    { name: 'Income', items: ['Salary', 'Freelance', 'Investments'], isIncome: true },
    { name: 'Housing', items: ['Rent/Mortgage', 'Utilities', 'Insurance'], isIncome: false },
    { name: 'Transportation', items: ['Car Payment', 'Gas', 'Maintenance'], isIncome: false },
    { name: 'Food', items: ['Groceries', 'Dining Out'], isIncome: false },
    { name: 'Savings', items: ['Emergency Fund', '401k', 'Investments'], isIncome: false },
  ];

  const sheet = workbook.addSheet('Budget');

  // Title
  sheet.cell('A1').value = `${year} Budget Tracker`;
  sheet.cell('A1').style = { font: { bold: true, size: 20 } };
  sheet.mergeCells('A1:N1');

  // Month headers
  sheet.cell('A3').value = 'Category';
  months.forEach((month, index) => {
    sheet.cell(2, index + 1).value = month;
    sheet.cell(2, index + 1).style = {
      font: { bold: true, color: '#FFFFFF' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#374151' },
      alignment: { horizontal: 'center' },
    };
  });
  sheet.cell(2, 13).value = 'Total';
  sheet.cell(2, 13).style = {
    font: { bold: true, color: '#FFFFFF' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#1F2937' },
    alignment: { horizontal: 'center' },
  };

  let currentRow = 3;

  categories.forEach(category => {
    // Category header
    sheet.cell(currentRow, 0).value = category.name;
    sheet.cell(currentRow, 0).style = {
      font: { bold: true },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        foregroundColor: category.isIncome ? '#D1FAE5' : '#FEE2E2'
      },
    };

    // Total for category row
    const categoryStartRow = currentRow + 1;
    currentRow++;

    // Items
    category.items.forEach(item => {
      sheet.cell(currentRow, 0).value = `  ${item}`;

      // Empty cells for data entry (or 0 as placeholder)
      for (let m = 1; m <= 12; m++) {
        sheet.cell(currentRow, m).value = 0;
        sheet.cell(currentRow, m).style = {
          numberFormat: { formatCode: '$#,##0.00' },
        };
      }

      // Row total formula
      sheet.cell(currentRow, 13).setFormula(`SUM(B${currentRow + 1}:M${currentRow + 1})`);
      sheet.cell(currentRow, 13).style = {
        font: { bold: true },
        numberFormat: { formatCode: '$#,##0.00' },
        fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#F3F4F6' },
      };

      currentRow++;
    });

    // Category subtotal
    const categoryEndRow = currentRow - 1;
    sheet.cell(currentRow, 0).value = `${category.name} Subtotal`;
    sheet.cell(currentRow, 0).style = { font: { bold: true, italic: true } };

    for (let m = 1; m <= 13; m++) {
      const col = String.fromCharCode(65 + m);
      sheet.cell(currentRow, m).setFormula(`SUM(${col}${categoryStartRow + 1}:${col}${categoryEndRow + 1})`);
      sheet.cell(currentRow, m).style = {
        font: { bold: true },
        numberFormat: { formatCode: '$#,##0.00' },
        fill: {
          type: 'pattern',
          pattern: 'solid',
          foregroundColor: category.isIncome ? '#A7F3D0' : '#FECACA'
        },
      };
    }

    currentRow += 2; // Add spacing
  });

  // Net summary row
  sheet.cell(currentRow, 0).value = 'Net (Income - Expenses)';
  sheet.cell(currentRow, 0).style = { font: { bold: true, size: 12 } };

  for (let m = 1; m <= 13; m++) {
    // This is simplified - in real app you'd calculate income row - expense rows
    sheet.cell(currentRow, m).value = 0;
    sheet.cell(currentRow, m).style = {
      font: { bold: true, size: 12 },
      numberFormat: { formatCode: '$#,##0.00' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FEF3C7' },
    };
  }

  // Column widths
  sheet.setColumnWidth(0, 25);
  for (let m = 1; m <= 13; m++) {
    sheet.setColumnWidth(m, 12);
  }

  // Freeze headers
  sheet.freeze(3, 1);

  return workbookToXlsxBlob(workbook);
}
```

## Tips and Best Practices

### Performance

- Use `cell(row, col)` instead of `cell('A1')` in loops for better performance
- Batch style applications using `applyStyle(range, style)`
- For large datasets, consider limiting export size with pagination

### Memory

- Process large files in chunks when importing
- Clear references to workbooks when done
- Use streaming approaches for very large files (coming soon)

### Compatibility

- Test exported files in multiple Excel versions
- Use standard number format codes for best compatibility
- Avoid very large merged cell ranges (can cause issues in older Excel)

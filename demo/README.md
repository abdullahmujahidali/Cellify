# Cellify Demo

Interactive demo for testing Cellify's Excel import/export functionality.

## Running the Demo

1. Install dependencies (if not already done):

```bash
npm install
```

2. Start the demo server:

```bash
npm run demo
```

3. Open http://localhost:5173 in your browser (Vite will show the actual URL)

## Features

### Export Examples

- **Basic Example** - Simple data table
- **Styled Spreadsheet** - Fonts, colors, borders, merged cells
- **With Formulas** - Calculations that Excel will evaluate
- **Multi-Sheet** - Workbook with multiple worksheets
- **With Comments** - Cell comments with authors
- **With Hyperlinks** - URLs, email links, and internal references

### Import Testing

- Drag and drop `.xlsx` or `.csv` files
- View imported data in a table preview
- See import statistics (cells, formulas, merges)
- Re-export imported files

### Cell Editing

- **Click** to select a cell
- **Double-click** or **Enter** to edit
- **Arrow keys** / **Tab** to navigate
- **Escape** to cancel editing
- **Ctrl+Z** / **Cmd+Z** to undo changes

### Context Menu (Right-Click)

Right-click on any cell to access:

- **Copy** (Ctrl+C) - Copy cell value, style, and comment
- **Cut** (Ctrl+X) - Cut cell to clipboard
- **Paste** (Ctrl+V) - Paste from clipboard
- **Fill Color** - Set cell background color
- **Text Color** - Set font color
- **Bold** (Ctrl+B) - Toggle bold formatting
- **Italic** (Ctrl+I) - Toggle italic formatting
- **Add/Edit/Delete Comment** - Manage cell comments
- **Clear Cell** (Delete) - Clear cell content

### Column Filters

Click the filter button (▼) in any column header to:

- Search values in the column
- Select/deselect specific values to show
- Apply multiple filters across columns
- Clear filters to show all data

### Column & Row Selection

- **Click column header** (A, B, C...) to select entire column
- **Click row header** (1, 2, 3...) to select entire row
- Selected columns/rows are highlighted in green

### Column & Row Resizing

- Drag the right edge of column headers to resize columns
- Drag the bottom edge of row headers to resize rows

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Enter / F2 | Start editing selected cell |
| Escape | Cancel editing |
| Arrow keys | Navigate between cells |
| Tab / Shift+Tab | Move to next/previous cell |
| Ctrl+C / Cmd+C | Copy cell |
| Ctrl+X / Cmd+X | Cut cell |
| Ctrl+V / Cmd+V | Paste cell |
| Ctrl+B / Cmd+B | Toggle bold |
| Ctrl+I / Cmd+I | Toggle italic |
| Ctrl+Z / Cmd+Z | Undo last change |
| Delete / Backspace | Clear cell |

## Development

The demo uses Vite for development, which:

- Handles ES module resolution
- Hot reloads on changes
- Compiles TypeScript on-the-fly

## Implementation Reference

The demo serves as a reference implementation showing how to build interactive spreadsheet UIs using Cellify. Key patterns include:

### Cell Selection & Editing

```javascript
// Select a cell
const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
td.classList.add('selected');

// Edit cell value
const cell = sheet.cell(row, col);
cell.value = newValue;

// Edit with formula
cell.setFormula('SUM(A1:A10)');
```

### Cell Styling

```javascript
const cell = sheet.cell(row, col);

// Set fill color
cell.style = {
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FFFF00' }
};

// Set font style
cell.style = {
  font: { bold: true, italic: true, color: '#FF0000' }
};
```

### Comments

```javascript
// Add comment
cell.setComment('This is a comment', 'Author Name');

// Read comment
const comment = cell.comment;
console.log(comment.text, comment.author);

// Delete comment
cell.comment = undefined;
```

### Filtering (UI Pattern)

```javascript
// Get unique values from a column
const values = new Set();
for (let r = 0; r < rowCount; r++) {
  const cell = sheet.getCell(r, colIndex);
  values.add(cell?.value ?? '(Empty)');
}

// Filter rows by hiding DOM elements
rows.forEach(tr => {
  const cellValue = getCellValue(tr.dataset.row, colIndex);
  if (!allowedValues.has(cellValue)) {
    tr.classList.add('filtered-out');
  }
});
```

See the full implementation in `demo/index.html` for complete examples.

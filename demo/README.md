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

### Import Testing

- Drag and drop `.xlsx` or `.csv` files
- View imported data in a table preview
- See import statistics (cells, formulas, merges)
- Re-export imported files

## Development

The demo uses Vite for development, which:
- Handles ES module resolution
- Hot reloads on changes
- Compiles TypeScript on-the-fly

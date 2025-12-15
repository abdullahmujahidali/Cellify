# Cellify

A lightweight, affordable spreadsheet library for building Excel-like experiences.

> **Status:** Alpha - Core data structures are implemented. Excel import/export coming soon.

## Why Cellify?

Most spreadsheet libraries are either expensive, bloated, or lack essential features. Cellify aims to be:

- **Lightweight** - Minimal dependencies (just ZIP compression)
- **Full-featured** - Merging, styling, formulas, validation
- **Type-safe** - Built with TypeScript from the ground up
- **Universal** - Works in Node.js and browsers

## Installation

```bash
npm install cellify
```

## Quick Start

```typescript
import { Workbook } from 'cellify';

// Create a workbook
const workbook = new Workbook();
const sheet = workbook.addSheet('Sales Data');

// Set values
sheet.cell('A1').value = 'Product';
sheet.cell('B1').value = 'Revenue';
sheet.cell('A2').value = 'Widget';
sheet.cell('B2').value = 1500;

// Apply styles
sheet.applyStyle('A1:B1', {
  font: { bold: true, size: 12 },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#4472C4' },
  alignment: { horizontal: 'center' }
});

// Merge cells
sheet.mergeCells('A1:B1');

// Set column widths
sheet.setColumnWidth(0, 20);
sheet.setColumnWidth(1, 15);

// Freeze header row
sheet.freeze(1);
```

## Features

### Implemented

- [x] Workbook, Sheet, Cell data structures
- [x] Cell values (string, number, boolean, date, rich text)
- [x] Cell formulas (storage, not evaluation yet)
- [x] Complete styling system (fonts, fills, borders, alignment)
- [x] Cell merging with overlap detection
- [x] Row/column configuration (height, width, hidden)
- [x] Freeze panes
- [x] Auto filters (data structure)
- [x] Conditional formatting (data structure)
- [x] Data validation
- [x] Hyperlinks and comments
- [x] Named ranges
- [x] Sheet protection

### Planned

- [ ] Excel (.xlsx) import/export
- [ ] CSV import/export
- [ ] Formula evaluation engine
- [ ] Streaming for large files

## Documentation

- [Architecture Decisions](./docs/decisions/README.md) - Why we made certain technical choices

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for development setup and guidelines.

## License

MIT - see [LICENSE](./LICENSE)

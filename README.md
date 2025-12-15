# Cellify

[![CI](https://github.com/abdullahmujahidali/Cellify/actions/workflows/ci.yml/badge.svg)](https://github.com/abdullahmujahidali/Cellify/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/abdullahmujahidali/Cellify/graph/badge.svg)](https://codecov.io/gh/abdullahmujahidali/Cellify)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A lightweight, affordable spreadsheet library for building Excel-like experiences.

> **Status:** Alpha - Core data structures are implemented. Excel import/export coming soon.

## Why Cellify?

Most spreadsheet libraries are either expensive, bloated, or lack essential features. Cellify aims to be:

- **Lightweight** - Minimal dependencies (just ZIP compression)
- **Full-featured** - Merging, styling, formulas, validation
- **Type-safe** - Built with TypeScript from the ground up
- **Universal** - Works in Node.js and browsers
- **Accessible** - Built-in a11y helpers for screen readers

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
- [x] Accessibility helpers (ARIA attributes, screen reader announcements)

### Planned

- [ ] Excel (.xlsx) import/export
- [ ] CSV import/export
- [ ] Formula evaluation engine
- [ ] Streaming for large files

## Accessibility

Cellify provides built-in helpers for creating accessible spreadsheet UIs:

```typescript
import {
  getCellAccessibility,
  getAriaAttributes,
  announceNavigation,
} from 'cellify';

// Generate ARIA attributes for a cell
const a11y = getCellAccessibility(cell, sheet, {
  headerRows: 1,
  headerCols: 1,
});

const ariaProps = getAriaAttributes(a11y);
// Returns: { role: 'gridcell', 'aria-colindex': 2, 'aria-rowindex': 3, ... }

// Announce navigation for screen readers
const announcement = announceNavigation(cell);
// Returns: { message: 'Cell B3, row 3, column 2, 1500', type: 'navigation', priority: 'polite' }
```

See [ADR-004: Accessibility](./docs/decisions/004-accessibility.md) for design details.

## Documentation

- [Architecture Decisions](./docs/decisions/README.md) - Why we made certain technical choices

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for development setup and guidelines.

## License

MIT - see [LICENSE](./LICENSE)

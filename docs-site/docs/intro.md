---
sidebar_position: 1
slug: /
---

# Introduction

**Cellify** is a lightweight, zero-dependency spreadsheet library for JavaScript and TypeScript. It provides a simple, intuitive API for creating, reading, and manipulating Excel (.xlsx) and CSV files.

## Features

- **Zero Dependencies** - Only uses [fflate](https://github.com/101arrowz/fflate) for ZIP compression (~8KB)
- **Full TypeScript Support** - Written in TypeScript with complete type definitions
- **Excel Import/Export** - Read and write `.xlsx` files with full formatting support
- **CSV Import/Export** - Handle CSV files with configurable delimiters
- **Rich Styling** - Fonts, colors, borders, alignment, number formats
- **Cell Operations** - Merge cells, freeze panes, auto-filters
- **Formulas** - Preserve and create Excel formulas
- **Accessibility** - Built-in helpers for screen readers

## Why Cellify?

| Feature | Cellify | ExcelJS | SheetJS |
|---------|---------|---------|---------|
| Bundle Size | ~8KB | ~1MB | ~500KB |
| Dependencies | 1 (fflate) | 15+ | 0 |
| TypeScript | Native | Yes | Partial |
| Styling | Full | Full | Limited |
| Streaming | No | Yes | Yes |

Cellify is perfect for:
- **Client-side Excel generation** - Small bundle, works in browsers
- **Simple spreadsheet tasks** - When you don't need streaming for huge files
- **TypeScript projects** - First-class type support

## Quick Example

```typescript
import { Workbook, workbookToXlsxBlob } from 'cellify';

// Create a workbook
const workbook = new Workbook();
const sheet = workbook.addSheet('Sales Report');

// Add headers with styling
sheet.cell(0, 0).value('Product').style({ font: { bold: true } });
sheet.cell(0, 1).value('Revenue').style({ font: { bold: true } });

// Add data
sheet.cell(1, 0).value('Widget A');
sheet.cell(1, 1).value(1500);

// Export to Excel
const blob = workbookToXlsxBlob(workbook);
```

## Installation

```bash
npm install cellify
```

Ready to get started? Head to the [Getting Started](/docs/getting-started) guide!

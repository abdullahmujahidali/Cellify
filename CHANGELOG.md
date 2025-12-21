# Changelog

All notable changes to Cellify will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **Event System**
  - `sheet.on()` and `sheet.off()` for subscribing to sheet events
  - `cellChange` event for value/formula changes
  - `cellStyleChange` event for style changes
  - `cellAdded` event when new cells are created
  - `cellDeleted` event when cells are deleted
  - Wildcard `'*'` listener for all events
  - `sheet.setEventsEnabled()` to disable events during bulk operations

- **Change Tracking**
  - `sheet.getChanges()` returns all changes since last commit
  - `sheet.commitChanges()` clears the change buffer
  - `sheet.changeCount` property for pending change count
  - Each change has unique ID, type, address, old/new values, and timestamp

## [0.1.0] - 2025-12-21

### Added

- **Core Features**
  - `Workbook` class for managing spreadsheet documents
  - `Sheet` class with cell management, row/column configuration
  - `Cell` class with values, formulas, styles, comments, hyperlinks, and validation

- **Excel Support**
  - XLSX import with `xlsxToWorkbook()` and `xlsxBlobToWorkbook()`
  - XLSX export with `workbookToXlsx()` and `workbookToXlsxBlob()`
  - Shared strings and style registry for optimized file size
  - Optional WASM acceleration for large files

- **CSV Support**
  - CSV import with `csvToWorkbook()`, `csvToSheet()`, `csvBufferToWorkbook()`
  - CSV export with `sheetToCsv()`, `sheetToCsvBuffer()`, `sheetsToCsv()`
  - Automatic delimiter detection (comma, semicolon, tab, pipe)
  - Smart type detection (numbers, dates, booleans, percentages, currency)
  - RFC 4180 compliant parsing and writing

- **Styling**
  - Font styling (bold, italic, underline, color, size, family)
  - Fill patterns and colors
  - Border styles (thin, medium, thick, double, dashed, dotted)
  - Cell alignment (horizontal, vertical, text wrap, rotation)
  - Number formats

- **Cell Features**
  - Formula support (storage and cached results)
  - Cell comments with author
  - Hyperlinks with tooltips
  - Data validation (whole, decimal, list, date, time, textLength, custom)
  - Merged cells

- **Sheet Features**
  - Row height and column width configuration
  - Hidden rows and columns
  - Frozen panes
  - Auto-filter
  - Sheet protection
  - Named ranges

- **Accessibility**
  - ARIA attribute helpers
  - Screen reader announcements
  - Keyboard navigation support

- **Developer Experience**
  - Full TypeScript support with comprehensive types
  - ESM and CommonJS builds
  - Zero dependencies for core (only fflate for compression)
  - Works in Node.js, Bun, Deno, and browsers

[Unreleased]: https://github.com/abdullahmujahidali/Cellify/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/abdullahmujahidali/Cellify/releases/tag/v0.1.0

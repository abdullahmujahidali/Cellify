# Changelog

All notable changes to Cellify will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.3.0] - 2025-12-22

### Added

- **Search & Replace**
  - `sheet.find(options)` to find the first cell matching search criteria
  - `sheet.findAll(options)` to find all cells matching search criteria
  - `sheet.replace(options)` to find and replace the first match
  - `sheet.replaceAll(options)` to find and replace all matches
  - Search by string, number, or regular expression
  - Options for case-sensitive matching, whole cell matching
  - Search in values, formulas, or both
  - Search within specific range or entire sheet

- **Copy/Paste Operations**
  - `sheet.copyRange(range)` to copy cells to internal clipboard
  - `sheet.cutRange(range)` to cut cells (copy and clear originals)
  - `sheet.pasteRange(target, options)` to paste clipboard contents
  - `sheet.duplicateRange(source, target)` to copy and paste in one operation
  - `sheet.clearClipboard()` to clear the internal clipboard
  - `sheet.hasClipboard` property to check if clipboard has content
  - Paste options: `valuesOnly`, `stylesOnly`, `transpose`
  - Preserves values, styles, and formulas when copying

- **Row/Column Operations**
  - `sheet.insertRow(index, count)` to insert rows at position
  - `sheet.insertColumn(index, count)` to insert columns at position
  - `sheet.deleteRow(index, count)` to delete rows
  - `sheet.deleteColumn(index, count)` to delete columns
  - `sheet.moveRow(from, to)` to move a row to new position
  - `sheet.moveColumn(from, to)` to move a column to new position
  - Preserves cell values, styles, formulas, and comments
  - Shifts existing cells correctly when inserting/deleting
  - Updates row/column configurations (height, width) when shifting

- **Data Import/Export Helpers**
  - `sheet.fromArray(data, options)` to populate sheet from 2D array
  - `sheet.fromObjects(data, options)` to populate sheet from array of objects
  - `sheet.toArray(options)` to export sheet data as 2D array
  - `sheet.toObjects<T>()` to export sheet data as typed array of objects
  - `sheet.appendRow(values)` to append a single row at the end
  - `sheet.appendRows(rows)` to append multiple rows
  - Options for custom start position, header styling, column selection
  - Handles empty arrays and sheets gracefully
  - Round-trip preservation (fromArray -> toArray)

## [0.2.0] - 2025-12-21

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

- **Undo/Redo**
  - `sheet.undo()` and `sheet.redo()` for reversing changes
  - `sheet.canUndo` and `sheet.canRedo` to check availability
  - `sheet.undoCount` and `sheet.redoCount` for history size
  - `sheet.batch(() => {...})` to group changes as single undo step
  - `sheet.clearHistory()` to clear undo/redo stacks
  - `sheet.setMaxUndoHistory(n)` to limit history size (default: 100)

- **Sorting**
  - `sheet.sort(column, options)` for single column sorting
  - `sheet.sortBy(columns, options)` for multi-column sorting
  - Ascending/descending order support
  - Header row preservation with `hasHeader` option
  - Numeric sorting for string numbers with `numeric` option
  - Case-insensitive sorting by default
  - Date values sorted correctly
  - Null values sorted to end
  - Preserves cell styles, formulas, and comments when sorting
  - Range-specific sorting with `range` option

- **Filtering**
  - `sheet.filter(column, criteria)` for single column filtering
  - `sheet.filterBy(filters)` for multi-column filtering
  - `sheet.clearFilter()` to remove all filters
  - `sheet.clearColumnFilter(column)` to remove filter on specific column
  - Criteria options: `equals`, `notEquals`, `contains`, `startsWith`, `endsWith`
  - Numeric criteria: `greaterThan`, `lessThan`, `between`, `notBetween`
  - Value list criteria: `in`, `notIn`
  - Empty checks: `isEmpty`, `isNotEmpty`
  - Custom filter function support
  - Case-insensitive string matching
  - `sheet.isRowFiltered(row)` to check if row is hidden by filter
  - `sheet.activeFilters` to get current filter configuration
  - `sheet.filteredRows` to get set of filtered row indices

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

[Unreleased]: https://github.com/abdullahmujahidali/Cellify/compare/v0.3.0...HEAD
[0.3.0]: https://github.com/abdullahmujahidali/Cellify/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/abdullahmujahidali/Cellify/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/abdullahmujahidali/Cellify/releases/tag/v0.1.0

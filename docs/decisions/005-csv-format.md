# ADR-005: CSV Import/Export Implementation

**Date:** 2025-12-15
**Status:** Accepted

## Context

CSV (Comma-Separated Values) is the simplest and most widely supported spreadsheet format. Implementing CSV support first provides:

1. A quick win for users who need basic import/export
2. A foundation for the import/export architecture
3. Test coverage for value serialization

## Decision

### RFC 4180 Compliance

We follow RFC 4180 for CSV formatting:

- Fields containing delimiter, quote, or newline are quoted
- Quote characters within quoted fields are doubled (`""`)
- CRLF (`\r\n`) as default line ending for Excel compatibility
- Optional BOM for UTF-8 encoding recognition

### Auto-Detection

For imports, we auto-detect the delimiter by analyzing the first line:

- Count occurrences of common delimiters (`,`, `;`, `\t`, `|`) outside quotes
- Use the delimiter with the highest count
- Users can override with explicit `delimiter` option

### Type Detection

When importing, we attempt to convert strings to appropriate types:

| Pattern | Type | Example |
|---------|------|---------|
| `true`, `false` (case-insensitive) | Boolean | `TRUE` → `true` |
| Numeric pattern | Number | `42` → `42` |
| Percentage | Number (decimal) | `50%` → `0.5` |
| Currency (with symbol) | Number | `$100` → `100` |
| ISO date | Date | `2024-01-15` → Date object |

Type detection is configurable via `detectNumbers` and `detectDates` options.

### No Dependencies

CSV parsing is implemented from scratch:

- Custom RFC 4180 parser (~100 lines)
- No regex-based parsing (handles edge cases properly)
- Streaming-ready architecture (future enhancement)

### API Design

**Export:**

```typescript
sheetToCsv(sheet, options?)      // Returns string
sheetToCsvBuffer(sheet, options?) // Returns Uint8Array
sheetsToCsv(sheets, options?)     // Returns Map<name, csv>
```

**Import:**

```typescript
csvToWorkbook(csv, options?)      // Returns Workbook
csvToSheet(csv, sheet, options?)  // Imports into existing sheet
csvBufferToWorkbook(buffer, options?) // From Uint8Array
```

## Consequences

### Positive

- Simple, well-tested implementation
- Handles edge cases (quotes, newlines, encoding)
- Flexible options for different use cases
- Foundation for Excel format implementation

### Negative

- No streaming support yet (loads entire file into memory)
- Limited date format detection
- Style information is lost (inherent to CSV format)

### Future Enhancements

- Streaming parser for large files
- More date format patterns
- Column type hints from headers

## Notes

CSV is a lossy format - styling, formulas, and cell metadata are not preserved. This is documented behavior. For full fidelity, users should use Excel format (planned).

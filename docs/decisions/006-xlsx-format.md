# ADR-006: XLSX Import/Export Implementation

**Date:** 2024-12-15
**Status:** Accepted

## Context

Cellify needs to read and write Excel (.xlsx) files, which is the primary format for spreadsheet interchange. The XLSX format (OOXML) is a ZIP archive containing XML files that describe workbook structure, shared strings, styles, and worksheet data.

Key considerations:

- Keep the minimal-dependencies philosophy (ADR-002)
- Support all core features: values, formulas, styles, merges, freeze panes
- Work in both Node.js and browsers
- Handle large files efficiently

## Decision

### 1. ZIP Compression

Use `fflate` for ZIP compression/decompression:

- Only 8KB gzipped (vs 120KB+ for JSZip)
- Synchronous API for simplicity
- Works in Node.js and browsers
- Tree-shakeable

### 2. XML Parsing (Import)

Use regex-based parsing instead of DOM/SAX parsers:

- Zero additional dependencies
- Predictable performance
- Sufficient for well-formed OOXML
- Custom `xlsx.parser.ts` module with:
  - `parseElement(xml, tagName)` - Parse single element
  - `parseElements(xml, tagName)` - Parse all matching elements
  - `parseCellRef(ref)` - Convert A1 notation to coordinates
  - `unescapeXml(str)` - Handle XML entities

### 3. XML Generation (Export)

Use template string concatenation:

- No XML library needed
- Full control over output format
- Helper functions in `xlsx.xml.ts`:
  - `escapeXml(str)` - Escape special characters
  - `el(tag, attrs, content)` - Create XML element
  - `emptyEl(tag, attrs)` - Create self-closing element

### 4. Module Structure

```
src/formats/xlsx/
├── index.ts           # Public exports
├── xlsx.writer.ts     # Export: workbookToXlsx()
├── xlsx.reader.ts     # Import: xlsxToWorkbook()
├── xlsx.reader.types.ts  # Import options and types
├── xlsx.parser.ts     # XML parsing utilities
├── xlsx.types.ts      # Shared types, namespaces
├── xlsx.utils.ts      # Date conversion, cell refs
├── xlsx.xml.ts        # XML generation utilities
├── xlsx.parts.ts      # Static XML parts
├── xlsx.strings.ts    # Shared strings table
└── xlsx.styles.ts     # Style XML generation
```

### 5. Import Strategy

1. **Unzip** the XLSX file using fflate
2. **Parse workbook.xml** for sheet list and relationships
3. **Parse sharedStrings.xml** for string lookup table
4. **Parse styles.xml** for:
   - Number formats (detect dates)
   - Fonts, fills, borders
   - Cell format index (xf) mapping
5. **Parse each worksheet** for:
   - Cell values with type detection
   - Formulas
   - Merged cells
   - Column widths, row heights
   - Freeze panes
   - Auto filters
6. **Parse docProps/core.xml** for document properties

### 6. Cell Type Detection

| Excel Type | Detection | Cellify Type |
|------------|-----------|--------------|
| (none) | `<v>` contains number | Number or Date* |
| t="s" | Shared string index | String |
| t="b" | Boolean 0/1 | Boolean |
| t="e" | Error value | Error string |
| t="inlineStr" | `<is><t>` content | String |
| t="str" | Formula string result | String |

*Dates detected by number format ID or format code pattern

### 7. Import Options

```typescript
interface XlsxImportOptions {
  sheets?: string[] | number[] | 'all';
  importFormulas?: boolean;
  importStyles?: boolean;
  importMergedCells?: boolean;
  importDimensions?: boolean;
  importFreezePanes?: boolean;
  importProperties?: boolean;
  maxRows?: number;
  maxCols?: number;
  onProgress?: (phase, current, total) => void;
}
```

## Consequences

### Benefits

- **Minimal footprint**: Only adds ~8KB (fflate) to bundle
- **No XML dependencies**: Custom parsing is sufficient for OOXML
- **Full control**: Can optimize for specific XLSX structures
- **Testable**: Pure functions, easy to unit test

### Trade-offs

- **No streaming**: Files loaded entirely into memory
- **Basic XML parsing**: Won't handle malformed XML gracefully
- **No rich text preservation**: Flattened to plain text on import
- **No image/chart support**: Data-focused implementation

### Future Considerations

- Add streaming support for very large files
- Support hyperlinks and comments
- Add data validation import/export
- Consider conditional formatting export

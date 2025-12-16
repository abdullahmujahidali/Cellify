# ADR 007: WebAssembly Parser for Performance

## Status

Accepted

## Context

Large XLSX files (100K+ cells) were taking 60-90+ seconds to import using the JavaScript-based regex XML parser. This created a poor user experience for files with many sheets or cells.

### Performance Analysis

The bottleneck was identified in the XML parsing layer (`xlsx.parser.ts`):

- Regex-based XML parsing accounted for ~70-80% of import time
- `parseElements()` created new RegExp objects repeatedly
- `findInnerContent()` did double passes through XML strings
- String operations in JavaScript are inherently slower than native code

## Decision

Implement an optional WebAssembly (WASM) parser using Rust with the following approach:

1. **Rust + quick-xml**: Use Rust's `quick-xml` crate for high-performance streaming XML parsing
2. **wasm-bindgen**: Expose Rust functions to JavaScript with automatic type conversion
3. **Automatic Fallback**: JavaScript parser remains as fallback when WASM is unavailable
4. **Pre-built Distribution**: Include compiled WASM in the npm package for zero-config usage

### Architecture

```
src/formats/xlsx/
├── xlsx.reader.ts       # Main reader (orchestrates parsing)
├── xlsx.parser.ts       # JavaScript parser (fallback)
├── xlsx.wasm.ts         # WASM module wrapper
├── xlsx.parser.wasm.ts  # Accelerated parser interface
└── wasm/                # Built WASM module
    ├── cellify_wasm.js
    ├── cellify_wasm.d.ts
    └── cellify_wasm_bg.wasm

wasm/
├── Cargo.toml           # Rust dependencies
├── src/lib.rs           # Rust parser implementation
└── build.sh             # Build script
```

### WASM Module Functions

- `parse_worksheet(xml)` - Parse worksheet cells, rows, merges
- `parse_shared_strings(xml)` - Parse shared strings table
- `parse_styles(xml)` - Parse fonts, fills, borders, number formats
- `parse_workbook(xml)` - Parse sheet metadata
- `parse_relationships(xml)` - Parse part relationships

## Consequences

### Positive

- **10-50x Performance Improvement**: Large files import in seconds instead of minutes
- **Zero Configuration**: WASM is pre-built and auto-initialized
- **Graceful Degradation**: Falls back to JavaScript if WASM unavailable
- **No Breaking Changes**: Existing API unchanged, WASM is opt-out not opt-in
- **Bundle Size**: WASM module is ~112KB (gzipped ~40KB)

### Negative

- **Build Complexity**: Requires Rust toolchain for development
- **CI Complexity**: Additional build step in GitHub Actions
- **Debugging**: WASM errors are harder to debug than JavaScript

### Performance Results

| File | Cells | JS Parser | WASM Parser | Speedup |
|------|-------|-----------|-------------|---------|
| Small | 1K | 50ms | 10ms | 5x |
| Medium | 10K | 500ms | 50ms | 10x |
| Large | 117K | 88s | ~5s | 17x |

## Usage

```typescript
import { initXlsxWasm, xlsxToWorkbook } from 'cellify';

// Optional: Initialize at startup for best performance
await initXlsxWasm();

// WASM is used automatically when available
const { workbook } = xlsxToWorkbook(buffer);

// Disable WASM for specific imports
const { workbook } = xlsxToWorkbook(buffer, { useWasm: false });
```

## Alternatives Considered

1. **Web Workers**: Move parsing to background thread
   - Rejected: Still slow, just non-blocking

2. **Streaming Parser**: Process XML in chunks
   - Rejected: Complex implementation, moderate gains

3. **Native Node Addon**: Use N-API for Node.js
   - Rejected: Not portable to browsers

4. **Different JS Parser**: Use SAX-style parser
   - Rejected: Still limited by JavaScript performance

## References

- [quick-xml](https://github.com/tafia/quick-xml) - Fast Rust XML parser
- [wasm-bindgen](https://rustwasm.github.io/wasm-bindgen/) - Rust/JS interop
- [wasm-pack](https://rustwasm.github.io/wasm-pack/) - WASM build tool

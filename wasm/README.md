# Cellify WASM Parser

High-performance XLSX XML parser using WebAssembly for 10-50x faster file imports.

## Prerequisites

1. **Rust** - Install via [rustup](https://rustup.rs/):

   ```bash
   curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
   ```

2. **wasm-pack** - Install using cargo:

   ```bash
   cargo install wasm-pack
   ```

3. **WASM target** - Add the WebAssembly target:

   ```bash
   rustup target add wasm32-unknown-unknown
   ```

## Building

From the project root:

```bash
npm run build:wasm
```

Or manually from the wasm directory:

```bash
cd wasm
./build.sh
```

The build will output to `src/formats/xlsx/wasm/`.

## Usage

The WASM parser is automatically used when available. Initialize it at application startup for best performance:

```typescript
import { initXlsxWasm, xlsxToWorkbook } from 'cellify';

// Initialize WASM at startup (optional but recommended)
await initXlsxWasm();

// Import will automatically use WASM if available
const { workbook } = xlsxToWorkbook(xlsxBuffer);
```

### Checking WASM Status

```typescript
import { isXlsxWasmReady } from 'cellify';

if (isXlsxWasmReady()) {
  console.log('Using WASM-accelerated parser');
} else {
  console.log('Using JavaScript parser (fallback)');
}
```

### Disabling WASM

To force JavaScript parsing:

```typescript
const { workbook } = xlsxToWorkbook(buffer, { useWasm: false });
```

## Performance

Expected performance improvements for large files:

| File Size | JS Parser | WASM Parser | Speedup |
|-----------|-----------|-------------|---------|
| 10K cells | ~500ms    | ~50ms       | 10x     |
| 100K cells| ~5s       | ~200ms      | 25x     |
| 1M cells  | ~50s      | ~2s         | 25x     |

## Development

### Running Rust Tests

```bash
cd wasm
cargo test
```

### Debug Build

```bash
wasm-pack build --target web --out-dir ../src/formats/xlsx/wasm --dev
```

## Architecture

The WASM module provides parsers for:

- **Worksheet XML** - Cell data, rows, columns, merges
- **Shared Strings** - String deduplication table
- **Styles** - Fonts, fills, borders, number formats
- **Workbook** - Sheet metadata and relationships
- **Relationships** - Part linking and hyperlinks

All parsing uses `quick-xml`, a high-performance Rust XML parser that uses zero-copy parsing and streaming for minimal memory overhead.

## Fallback Behavior

If WASM is unavailable (not built, failed to load, or disabled), the library automatically falls back to the JavaScript parser. This ensures:

- No breaking changes for existing users
- Graceful degradation in environments without WASM support
- Optional opt-in for performance benefits

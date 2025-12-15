# ADR-003: Core Data Structures

**Date:** 2024-12-15
**Status:** Accepted

## Context

We needed to design the core data structures for representing spreadsheets in memory. The key considerations were:

1. Memory efficiency for large spreadsheets
2. Fast cell access by address
3. Support for sparse data (most cells are empty)
4. Clean API for developers

## Decision

### Hierarchy: Workbook → Sheet → Cell

We chose a traditional hierarchy that mirrors Excel's structure:

- **Workbook**: Top-level container holding sheets, styles, and metadata
- **Sheet**: A single worksheet with cells, merges, and configuration
- **Cell**: Individual cell with value, formula, style, and metadata

### Sparse Cell Storage

Cells are stored in a `Map<string, Cell>` where the key is `"row,col"` (e.g., `"0,0"` for A1).

**Why Map over 2D Array:**

- Real spreadsheets are sparse - most cells are empty
- A 2D array for 1M rows × 16K columns would be massive even if empty
- Map only stores cells that have data
- O(1) access by key

**Why string key over nested Maps:**

- `Map<number, Map<number, Cell>>` requires two lookups
- String key allows single lookup with `"row,col"`
- Slightly more memory per key, but simpler code

### Cell Class vs Plain Object

We chose a Cell class rather than plain objects because:

- Methods for common operations (`applyStyle`, `setFormula`, `clear`)
- Encapsulation of internal state (merge info, validation)
- Fluent API for chaining: `cell.setFormula('=A1').applyStyle({...})`
- TypeScript gets better inference with classes

### Style Storage

Styles are stored directly on cells, not in a shared style table.

**Trade-off:**

- Pros: Simpler API, no indirection
- Cons: More memory if many cells share identical styles

**Mitigation:** For Excel export, we'll deduplicate styles into a style table. Internally, memory is less critical than API simplicity for most use cases.

### Merge Tracking

Merges are tracked in two places:

1. Sheet level: Array of merge ranges for iteration
2. Cell level: Master cell knows its merge range; slave cells know their master

This dual tracking allows:

- Fast iteration over all merges (sheet array)
- Fast lookup of merge status for any cell (cell property)

## Consequences

### Positive

- Clean, intuitive API
- Efficient for sparse data
- Easy to serialize/deserialize
- Type-safe with full IntelliSense

### Negative

- Style deduplication needed for export
- String key parsing has minor overhead
- Cell objects have some memory overhead vs plain values

### Trade-offs Accepted

- Memory overhead of Cell objects for API ergonomics
- Per-cell style storage for simplicity over shared styles

## Notes

If memory becomes a concern for very large spreadsheets, we could:

1. Add a "compact mode" that uses plain objects
2. Implement style sharing with reference counting
3. Use typed arrays for numeric-heavy sheets

These optimizations can be added without breaking the public API.

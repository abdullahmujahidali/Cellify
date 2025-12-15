# ADR-002: Minimal Dependencies Philosophy

**Date:** 2025-11-25
**Status:** Accepted

## Context

We needed to decide how much to rely on external libraries vs building custom implementations.

## Decision

### Build Custom (Core Value)
- Data structures (Workbook, Sheet, Cell, Range)
- Cell merging and row/col span logic
- Styling system (fonts, fills, borders, alignment)
- Formula parser and evaluator
- CSV import/export
- Excel XML generation/parsing

### External Dependencies (Pragmatic)
Only two runtime dependencies:

1. **fflate** - ZIP compression/decompression
   - Excel .xlsx files are ZIP archives containing XML
   - ZIP compression is a well-defined, stable algorithm
   - Writing our own would be reinventing the wheel with no benefit
   - fflate is small (~8KB), fast, and tree-shakeable

2. **vitest** (dev only) - Testing framework
   - Modern, fast, TypeScript-native
   - Compatible with Jest APIs (familiar to most developers)

### What We Don't Use
- **XML parsing libraries**: We'll write a minimal XML parser tailored to Excel's specific XML schemas. This gives us control and reduces bloat.
- **Styling libraries**: Styling is just data structures - no library needed
- **Date libraries**: We'll handle Excel's date serial numbers ourselves
- **Formula parsing libraries**: Custom parser gives us full control over supported functions

## Consequences

### Positive
- Minimal bundle size
- No dependency hell or version conflicts
- Full control over behavior
- No abandoned dependency risk
- Deep understanding of the codebase

### Negative
- More code to write and maintain
- Must implement standard algorithms ourselves
- Potential for bugs that libraries have already solved

### Trade-offs Accepted
- We accept more initial development effort for long-term maintainability
- We accept reimplementing some wheels where the wheel is simple enough
- We refuse to reimplement complex, well-defined algorithms (like ZIP compression)

## Notes
This decision can be revisited if specific implementations prove too complex or error-prone. The goal is pragmatic minimalism, not dogmatic zero-dependency.

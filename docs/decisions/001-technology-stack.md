# ADR-001: Technology Stack

**Date:** 2024-12-15
**Status:** Accepted

## Context

We needed to choose the core technology stack for Cellify, a spreadsheet library that aims to be a lightweight, affordable alternative to existing solutions like ExcelJS.

## Decision

### Language: TypeScript

**Chosen over:** JavaScript, Rust + WASM

**Reasoning:**
- **Type safety**: A spreadsheet library has complex data structures (cells, styles, ranges, formulas). TypeScript catches errors at compile time and provides excellent IDE support.
- **Developer experience**: Most web developers know TypeScript. This lowers the barrier for contributions when we open source.
- **Performance is sufficient**: ExcelJS handles millions of cells in pure JavaScript. For most use cases, the overhead of JavaScript/TypeScript is negligible compared to I/O operations (file reading/writing).
- **Rust + WASM rejected**: While faster, it would significantly increase complexity, reduce the contributor pool, and add build complexity. If specific hot paths become bottlenecks later, we can optimize those to WASM without rewriting the entire library.

### Target: Universal (Node.js + Browser)

**Reasoning:**
- Server-side: Generate Excel/CSV files, parse uploads
- Browser: Manipulate data, integrate with UI frameworks
- Same API everywhere reduces learning curve
- Headless-first approach makes this natural - the core is just data structures

### Rendering: Headless First

**Chosen over:** Canvas-first, DOM-first

**Reasoning:**
- Separates concerns: data layer vs presentation
- Maximum flexibility: users can plug in their own rendering
- Smaller bundle for users who only need data manipulation
- Rendering can be added as optional packages later
- Allows the same core to work in Node.js (no DOM) and browser

### Build Output: Dual ESM/CJS

**Reasoning:**
- ESM: Modern bundlers, native ES modules
- CJS: Legacy Node.js, older tooling
- Separate type declarations for best TypeScript support

## Consequences

### Positive
- Wide adoption potential due to familiar technology
- Easy to contribute to
- Works everywhere JavaScript runs
- Clean separation of concerns

### Negative
- Not the absolute fastest possible (vs native code)
- Bundle size larger than a minimal Rust/WASM solution

### Risks
- If performance becomes critical, may need to optimize hot paths
- Mitigation: Design APIs that allow internal optimization without breaking changes

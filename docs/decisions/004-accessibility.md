# ADR-004: Accessibility Architecture

**Date:** 2025-12-15
**Status:** Accepted

## Context

Spreadsheets are complex interactive data structures. Making them accessible to users with disabilities (visual impairments, motor disabilities, cognitive disabilities) requires careful consideration at the library level.

As a headless library, Cellify doesn't render UI directly. However, we can provide accessibility metadata and helpers that rendering layers can use to create accessible experiences.

## Decision

### Headless Accessibility Approach

We provide accessibility support through:

1. **Metadata Types**: TypeScript types for accessibility attributes
2. **Helper Functions**: Generate ARIA attributes, descriptions, and announcements
3. **Documentation**: Guide for rendering accessible spreadsheets

We do NOT:
- Render any UI (that's the consumer's job)
- Mandate specific rendering technologies
- Force accessibility patterns that might not fit all use cases

### Key Accessibility Features

**Cell Accessibility Metadata:**
- Role identification (gridcell, rowheader, columnheader)
- Scope for headers (row, col, rowgroup, colgroup)
- Header references for data cells
- Position information (aria-rowindex, aria-colindex)
- Span information for merged cells
- State information (selected, readonly, invalid)

**Sheet Accessibility Metadata:**
- Grid dimensions for virtualized rendering
- Header row/column configuration
- Multi-select capability
- Caption and summary support

**Screen Reader Support:**
- Human-readable cell descriptions
- Value descriptions for formatted numbers (e.g., "25 percent")
- Error value descriptions (e.g., "division by zero error")
- Navigation announcements
- Selection announcements

**Keyboard Navigation Helpers:**
- Configuration types for navigation modes
- Custom shortcut definitions
- Standard spreadsheet shortcuts documented

### ARIA Compliance

We follow WAI-ARIA 1.2 grid pattern recommendations:

- `role="grid"` for the spreadsheet container
- `role="row"` for rows
- `role="gridcell"` for data cells
- `role="rowheader"` and `role="columnheader"` for headers
- Proper `aria-colindex` and `aria-rowindex` for large/virtualized grids
- `aria-colspan` and `aria-rowspan` for merged cells

### Value Text Generation

For screen readers, we generate human-readable descriptions:

| Cell Value | Display | Screen Reader |
|------------|---------|---------------|
| `0.25` with `%` format | `25%` | "25 percent" |
| `100` with `$` format | `$100.00` | "100 dollars" |
| `#DIV/0!` | `#DIV/0!` | "division by zero error" |
| `null` | (empty) | "empty" |

## Consequences

### Positive

- Rendering layers have everything needed for WCAG compliance
- Consistent accessibility patterns across different renderers
- Screen reader users get meaningful descriptions
- Keyboard navigation is well-defined

### Negative

- Rendering layers must implement the patterns correctly
- Some overhead in generating metadata for every cell
- May not cover all edge cases in complex spreadsheets

### Trade-offs Accepted

- We provide helpers, not enforcement - accessibility is opt-in
- Focus on most common patterns; exotic layouts may need custom handling
- English-centric descriptions (i18n would be a future enhancement)

## Usage Example

```typescript
import {
  Sheet,
  getCellAccessibility,
  getAriaAttributes,
  announceNavigation
} from 'cellify';

// In a React renderer
function CellComponent({ cell, sheet }) {
  const a11y = getCellAccessibility(cell, sheet, {
    headerRows: 1,
    headerCols: 1,
  });

  const ariaProps = getAriaAttributes(a11y);

  return (
    <td {...ariaProps}>
      {cell.value}
    </td>
  );
}

// On navigation
function handleArrowKey(cell) {
  const announcement = announceNavigation(cell);
  // Send to live region for screen reader
  liveRegion.textContent = announcement.message;
}
```

## References

- [WAI-ARIA Grid Pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/)
- [WCAG 2.1 Guidelines](https://www.w3.org/WAI/WCAG21/quickref/)
- [Accessible Data Tables](https://www.w3.org/WAI/tutorials/tables/)

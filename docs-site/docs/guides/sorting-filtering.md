---
sidebar_position: 11
---

# Sorting & Filtering

Cellify provides powerful sorting and filtering capabilities for organizing and viewing your spreadsheet data.

## Sorting

### Single Column Sort

Sort rows by the values in a single column:

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

// Sample data
sheet.setValues('A1', [
  ['Name', 'Age', 'City'],
  ['Charlie', 30, 'NYC'],
  ['Alice', 25, 'LA'],
  ['Bob', 28, 'Chicago'],
]);

// Sort by column A (Name) ascending
sheet.sort('A', { hasHeader: true });

// Sort by column B (Age) descending
sheet.sort('B', { hasHeader: true, descending: true });
```

### Sort Options

```typescript
sheet.sort('A', {
  descending: false,    // Sort direction (default: false = ascending)
  hasHeader: true,      // Preserve header row (default: false)
  range: 'A1:C10',      // Sort specific range only
  numeric: true,        // Sort string numbers numerically
  caseSensitive: false, // Case-sensitive comparison (default: false)
});
```

### Multi-Column Sort

Sort by multiple columns with different options for each:

```typescript
// Sort by Name, then by Age (descending)
sheet.sortBy([
  { column: 'A' },                      // Primary sort
  { column: 'B', descending: true },    // Secondary sort
], { hasHeader: true });
```

### Sorting Preserves Data

When sorting, Cellify preserves:
- Cell styles (bold, colors, etc.)
- Formulas
- Hyperlinks
- Comments

```typescript
sheet.cell('A1').value = 'Important';
sheet.cell('A1').style = { font: { bold: true } };

sheet.sort('A');

// Style is preserved after sorting
console.log(sheet.cell('A3').style?.font?.bold); // true (if sorted to row 3)
```

### Null Values

Null/empty values are always sorted to the end, regardless of sort direction:

```typescript
sheet.cell('A1').value = 'B';
sheet.cell('A2').value = null;
sheet.cell('A3').value = 'A';

sheet.sort('A');
// Result: A, B, null
```

## Filtering

### Basic Filtering

Filter rows to show only those matching specific criteria:

```typescript
// Show only rows where Status equals 'Active'
sheet.filter('A', { equals: 'Active' });

// Show rows where Price is greater than 100
sheet.filter('B', { greaterThan: 100 });
```

### Filter Criteria

#### Equality

```typescript
sheet.filter('A', { equals: 'Active' });
sheet.filter('A', { notEquals: 'Inactive' });
```

#### String Operations

All string operations are case-insensitive by default:

```typescript
sheet.filter('A', { contains: 'test' });
sheet.filter('A', { notContains: 'draft' });
sheet.filter('A', { startsWith: 'Report' });
sheet.filter('A', { endsWith: '.pdf' });
```

#### Numeric Operations

```typescript
sheet.filter('A', { greaterThan: 100 });
sheet.filter('A', { greaterThanOrEqual: 100 });
sheet.filter('A', { lessThan: 50 });
sheet.filter('A', { lessThanOrEqual: 50 });
sheet.filter('A', { between: [10, 100] });
sheet.filter('A', { notBetween: [0, 10] });
```

#### Value Lists

```typescript
// Show rows where Status is 'Active' or 'Pending'
sheet.filter('A', { in: ['Active', 'Pending'] });

// Hide rows where Status is 'Deleted' or 'Archived'
sheet.filter('A', { notIn: ['Deleted', 'Archived'] });
```

#### Empty Checks

```typescript
sheet.filter('A', { isEmpty: true });      // Show only empty cells
sheet.filter('A', { isNotEmpty: true });   // Show only non-empty cells
```

#### Custom Filter Function

For complex filtering logic, use a custom function:

```typescript
// Show only even numbers
sheet.filter('A', {
  custom: (value) => typeof value === 'number' && value % 2 === 0
});

// Show dates in the current year
sheet.filter('A', {
  custom: (value) => value instanceof Date && value.getFullYear() === 2024
});
```

### Multi-Column Filtering

Filter by multiple columns (AND logic):

```typescript
sheet.filterBy([
  { column: 'A', criteria: { equals: 'Active' } },
  { column: 'B', criteria: { greaterThan: 100 } },
]);
// Shows rows where Status = 'Active' AND Price > 100
```

### Filter with Header

Preserve the header row when filtering:

```typescript
sheet.filter('A', { equals: 'Active' }, { hasHeader: true });
```

### Clearing Filters

```typescript
// Clear all filters
sheet.clearFilter();

// Clear filter on specific column only
sheet.clearColumnFilter('A');
```

### Checking Filter State

```typescript
// Check if a specific row is filtered (hidden)
if (sheet.isRowFiltered(5)) {
  console.log('Row 5 is hidden by filter');
}

// Get all filtered row indices
console.log(`${sheet.filteredRows.size} rows are hidden`);

// Get active filter configuration
for (const [colIndex, criteria] of sheet.activeFilters) {
  console.log(`Column ${colIndex} has filter:`, criteria);
}
```

## Combining Sort and Filter

You can combine sorting and filtering for powerful data views:

```typescript
const workbook = new Workbook();
const sheet = workbook.addSheet('Sales');

// Load data
sheet.setValues('A1', [
  ['Product', 'Category', 'Sales'],
  ['Widget A', 'Electronics', 500],
  ['Widget B', 'Electronics', 300],
  ['Gadget X', 'Accessories', 150],
  ['Gadget Y', 'Electronics', 800],
]);

// Filter to Electronics only
sheet.filter('B', { equals: 'Electronics' }, { hasHeader: true });

// Sort by Sales descending
sheet.sort('C', { hasHeader: true, descending: true });

// Now showing only Electronics, sorted by highest sales first
```

## Performance Tips

1. **Apply filters before sorting** when possible to reduce the dataset size.

2. **Use batch operations** for multiple changes:
   ```typescript
   sheet.filterBy([
     { column: 'A', criteria: { equals: 'Active' } },
     { column: 'B', criteria: { greaterThan: 100 } },
   ]);
   ```

3. **Clear filters before applying new ones** if you want to start fresh:
   ```typescript
   sheet.clearFilter();
   sheet.filter('A', { equals: 'New Status' });
   ```

import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Sheet Filtering', () => {
  describe('basic filter', () => {
    it('should filter by equals', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Active';
      sheet.cell('A2').value = 'Inactive';
      sheet.cell('A3').value = 'Active';

      sheet.filter('A', { equals: 'Active' });

      expect(sheet.isRowFiltered(0)).toBe(false); // Active - shown
      expect(sheet.isRowFiltered(1)).toBe(true);  // Inactive - hidden
      expect(sheet.isRowFiltered(2)).toBe(false); // Active - shown
    });

    it('should filter by notEquals', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';
      sheet.cell('A3').value = 'C';

      sheet.filter('A', { notEquals: 'B' });

      expect(sheet.isRowFiltered(0)).toBe(false); // A - shown
      expect(sheet.isRowFiltered(1)).toBe(true);  // B - hidden
      expect(sheet.isRowFiltered(2)).toBe(false); // C - shown
    });

    it('should filter by column index', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Yes';
      sheet.cell('A2').value = 'No';
      sheet.cell('A3').value = 'Yes';

      sheet.filter(0, { equals: 'Yes' }); // Column index 0 = A

      expect(sheet.isRowFiltered(0)).toBe(false);
      expect(sheet.isRowFiltered(1)).toBe(true);
      expect(sheet.isRowFiltered(2)).toBe(false);
    });
  });

  describe('string filters', () => {
    it('should filter by contains', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello World';
      sheet.cell('A2').value = 'Goodbye';
      sheet.cell('A3').value = 'Hello There';

      sheet.filter('A', { contains: 'Hello' });

      expect(sheet.isRowFiltered(0)).toBe(false); // Contains Hello
      expect(sheet.isRowFiltered(1)).toBe(true);  // Doesn't contain
      expect(sheet.isRowFiltered(2)).toBe(false); // Contains Hello
    });

    it('should filter by startsWith', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Apple';
      sheet.cell('A2').value = 'Banana';
      sheet.cell('A3').value = 'Apricot';

      sheet.filter('A', { startsWith: 'A' });

      expect(sheet.isRowFiltered(0)).toBe(false); // Starts with A
      expect(sheet.isRowFiltered(1)).toBe(true);  // Starts with B
      expect(sheet.isRowFiltered(2)).toBe(false); // Starts with A
    });

    it('should filter by endsWith', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'test.txt';
      sheet.cell('A2').value = 'data.csv';
      sheet.cell('A3').value = 'file.txt';

      sheet.filter('A', { endsWith: '.txt' });

      expect(sheet.isRowFiltered(0)).toBe(false);
      expect(sheet.isRowFiltered(1)).toBe(true);
      expect(sheet.isRowFiltered(2)).toBe(false);
    });

    it('should be case insensitive', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'ACTIVE';
      sheet.cell('A2').value = 'active';
      sheet.cell('A3').value = 'Active';

      sheet.filter('A', { equals: 'active' });

      expect(sheet.isRowFiltered(0)).toBe(false);
      expect(sheet.isRowFiltered(1)).toBe(false);
      expect(sheet.isRowFiltered(2)).toBe(false);
    });
  });

  describe('numeric filters', () => {
    it('should filter by greaterThan', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 10;
      sheet.cell('A2').value = 50;
      sheet.cell('A3').value = 100;

      sheet.filter('A', { greaterThan: 25 });

      expect(sheet.isRowFiltered(0)).toBe(true);  // 10 <= 25
      expect(sheet.isRowFiltered(1)).toBe(false); // 50 > 25
      expect(sheet.isRowFiltered(2)).toBe(false); // 100 > 25
    });

    it('should filter by lessThan', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 10;
      sheet.cell('A2').value = 50;
      sheet.cell('A3').value = 100;

      sheet.filter('A', { lessThan: 50 });

      expect(sheet.isRowFiltered(0)).toBe(false); // 10 < 50
      expect(sheet.isRowFiltered(1)).toBe(true);  // 50 not < 50
      expect(sheet.isRowFiltered(2)).toBe(true);  // 100 not < 50
    });

    it('should filter by between', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 5;
      sheet.cell('A2').value = 15;
      sheet.cell('A3').value = 25;

      sheet.filter('A', { between: [10, 20] });

      expect(sheet.isRowFiltered(0)).toBe(true);  // 5 not in range
      expect(sheet.isRowFiltered(1)).toBe(false); // 15 in range
      expect(sheet.isRowFiltered(2)).toBe(true);  // 25 not in range
    });
  });

  describe('value list filters', () => {
    it('should filter by in (value list)', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Red';
      sheet.cell('A2').value = 'Green';
      sheet.cell('A3').value = 'Blue';
      sheet.cell('A4').value = 'Yellow';

      sheet.filter('A', { in: ['Red', 'Blue'] });

      expect(sheet.isRowFiltered(0)).toBe(false); // Red in list
      expect(sheet.isRowFiltered(1)).toBe(true);  // Green not in list
      expect(sheet.isRowFiltered(2)).toBe(false); // Blue in list
      expect(sheet.isRowFiltered(3)).toBe(true);  // Yellow not in list
    });

    it('should filter by notIn', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';
      sheet.cell('A3').value = 'C';

      sheet.filter('A', { notIn: ['B'] });

      expect(sheet.isRowFiltered(0)).toBe(false);
      expect(sheet.isRowFiltered(1)).toBe(true);
      expect(sheet.isRowFiltered(2)).toBe(false);
    });
  });

  describe('empty filters', () => {
    it('should filter by isEmpty', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Value';
      sheet.cell('A2').value = null;
      sheet.cell('A3').value = '';

      sheet.filter('A', { isEmpty: true });

      expect(sheet.isRowFiltered(0)).toBe(true);  // Not empty
      expect(sheet.isRowFiltered(1)).toBe(false); // null is empty
      expect(sheet.isRowFiltered(2)).toBe(false); // '' is empty
    });

    it('should filter by isNotEmpty', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Value';
      sheet.cell('A2').value = null;
      sheet.cell('A3').value = 'Another';

      sheet.filter('A', { isNotEmpty: true });

      expect(sheet.isRowFiltered(0)).toBe(false); // Has value
      expect(sheet.isRowFiltered(1)).toBe(true);  // null is empty
      expect(sheet.isRowFiltered(2)).toBe(false); // Has value
    });
  });

  describe('custom filter', () => {
    it('should filter using custom function', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 1;
      sheet.cell('A2').value = 2;
      sheet.cell('A3').value = 3;
      sheet.cell('A4').value = 4;

      // Only show even numbers
      sheet.filter('A', {
        custom: (value) => typeof value === 'number' && value % 2 === 0
      });

      expect(sheet.isRowFiltered(0)).toBe(true);  // 1 is odd
      expect(sheet.isRowFiltered(1)).toBe(false); // 2 is even
      expect(sheet.isRowFiltered(2)).toBe(true);  // 3 is odd
      expect(sheet.isRowFiltered(3)).toBe(false); // 4 is even
    });
  });

  describe('filter with header', () => {
    it('should preserve header row when hasHeader is true', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Status'; // Header
      sheet.cell('A2').value = 'Active';
      sheet.cell('A3').value = 'Inactive';
      sheet.cell('A4').value = 'Active';

      sheet.filter('A', { equals: 'Active' }, { hasHeader: true });

      expect(sheet.isRowFiltered(0)).toBe(false); // Header not filtered
      expect(sheet.isRowFiltered(1)).toBe(false); // Active - shown
      expect(sheet.isRowFiltered(2)).toBe(true);  // Inactive - hidden
      expect(sheet.isRowFiltered(3)).toBe(false); // Active - shown
    });
  });

  describe('multi-column filter', () => {
    it('should filter by multiple columns', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Active';
      sheet.cell('B1').value = 100;
      sheet.cell('A2').value = 'Active';
      sheet.cell('B2').value = 50;
      sheet.cell('A3').value = 'Inactive';
      sheet.cell('B3').value = 100;

      sheet.filterBy([
        { column: 'A', criteria: { equals: 'Active' } },
        { column: 'B', criteria: { greaterThan: 75 } }
      ]);

      expect(sheet.isRowFiltered(0)).toBe(false); // Active AND > 75
      expect(sheet.isRowFiltered(1)).toBe(true);  // Active but NOT > 75
      expect(sheet.isRowFiltered(2)).toBe(true);  // Not Active
    });
  });

  describe('clear filter', () => {
    it('should clear all filters', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';
      sheet.cell('A3').value = 'C';

      sheet.filter('A', { equals: 'A' });

      expect(sheet.isRowFiltered(1)).toBe(true);
      expect(sheet.isRowFiltered(2)).toBe(true);

      sheet.clearFilter();

      expect(sheet.isRowFiltered(0)).toBe(false);
      expect(sheet.isRowFiltered(1)).toBe(false);
      expect(sheet.isRowFiltered(2)).toBe(false);
    });

    it('should clear filter on specific column', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Active';
      sheet.cell('B1').value = 100;
      sheet.cell('A2').value = 'Inactive';
      sheet.cell('B2').value = 100;

      sheet.filterBy([
        { column: 'A', criteria: { equals: 'Active' } },
        { column: 'B', criteria: { greaterThan: 50 } }
      ]);

      // Row 2 is hidden (Inactive)
      expect(sheet.isRowFiltered(1)).toBe(true);

      // Clear only column A filter
      sheet.clearColumnFilter('A');

      // Now row 2 should be visible (only B filter active, and B2=100 > 50)
      expect(sheet.isRowFiltered(1)).toBe(false);
    });
  });

  describe('activeFilters property', () => {
    it('should return active filters', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';

      expect(sheet.activeFilters.size).toBe(0);

      sheet.filter('A', { equals: 'Test' });

      expect(sheet.activeFilters.size).toBe(1);
      expect(sheet.activeFilters.get(0)).toEqual({ equals: 'Test' });
    });
  });

  describe('filteredRows property', () => {
    it('should return filtered row indices', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';
      sheet.cell('A3').value = 'A';

      sheet.filter('A', { equals: 'A' });

      expect(sheet.filteredRows.size).toBe(1);
      expect(sheet.filteredRows.has(1)).toBe(true);
    });
  });

  describe('filter on empty sheet', () => {
    it('should handle empty sheet gracefully', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      // Should not throw
      sheet.filter('A', { equals: 'test' });

      expect(sheet.activeFilters.size).toBe(1);
      expect(sheet.filteredRows.size).toBe(0);
    });
  });
});

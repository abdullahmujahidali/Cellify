import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Sheet Sorting', () => {
  describe('basic sort', () => {
    it('should sort by string column ascending', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Charlie';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('A3').value = 'Bob';

      sheet.sort('A');

      expect(sheet.cell('A1').value).toBe('Alice');
      expect(sheet.cell('A2').value).toBe('Bob');
      expect(sheet.cell('A3').value).toBe('Charlie');
    });

    it('should sort by string column descending', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Alice';
      sheet.cell('A2').value = 'Charlie';
      sheet.cell('A3').value = 'Bob';

      sheet.sort('A', { descending: true });

      expect(sheet.cell('A1').value).toBe('Charlie');
      expect(sheet.cell('A2').value).toBe('Bob');
      expect(sheet.cell('A3').value).toBe('Alice');
    });

    it('should sort by numeric column', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 30;
      sheet.cell('A2').value = 10;
      sheet.cell('A3').value = 20;

      sheet.sort('A');

      expect(sheet.cell('A1').value).toBe(10);
      expect(sheet.cell('A2').value).toBe(20);
      expect(sheet.cell('A3').value).toBe(30);
    });

    it('should sort by column index', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'C';
      sheet.cell('A2').value = 'A';
      sheet.cell('A3').value = 'B';

      sheet.sort(0); // Column index 0 = A

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('B');
      expect(sheet.cell('A3').value).toBe('C');
    });
  });

  describe('sort with header', () => {
    it('should preserve header row when hasHeader is true', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Name'; // Header
      sheet.cell('A2').value = 'Charlie';
      sheet.cell('A3').value = 'Alice';
      sheet.cell('A4').value = 'Bob';

      sheet.sort('A', { hasHeader: true });

      expect(sheet.cell('A1').value).toBe('Name'); // Header unchanged
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('A3').value).toBe('Bob');
      expect(sheet.cell('A4').value).toBe('Charlie');
    });
  });

  describe('sort preserves row data', () => {
    it('should move entire rows when sorting', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Charlie';
      sheet.cell('B1').value = 30;
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 25;
      sheet.cell('A3').value = 'Bob';
      sheet.cell('B3').value = 28;

      sheet.sort('A');

      expect(sheet.cell('A1').value).toBe('Alice');
      expect(sheet.cell('B1').value).toBe(25);
      expect(sheet.cell('A2').value).toBe('Bob');
      expect(sheet.cell('B2').value).toBe(28);
      expect(sheet.cell('A3').value).toBe('Charlie');
      expect(sheet.cell('B3').value).toBe(30);
    });

    it('should preserve cell styles when sorting', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'B';
      sheet.cell('A1').style = { font: { bold: true } };
      sheet.cell('A2').value = 'A';
      sheet.cell('A2').style = { font: { italic: true } };

      sheet.sort('A');

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A1').style?.font?.italic).toBe(true);
      expect(sheet.cell('A2').value).toBe('B');
      expect(sheet.cell('A2').style?.font?.bold).toBe(true);
    });
  });

  describe('sort with nulls', () => {
    it('should sort null values to the end', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'B';
      sheet.cell('A2').value = null;
      sheet.cell('A3').value = 'A';

      sheet.sort('A');

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('B');
      expect(sheet.cell('A3').value).toBeNull();
    });
  });

  describe('numeric sort option', () => {
    it('should sort string numbers numerically when numeric option is true', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = '10';
      sheet.cell('A2').value = '2';
      sheet.cell('A3').value = '1';

      sheet.sort('A', { numeric: true });

      expect(sheet.cell('A1').value).toBe('1');
      expect(sheet.cell('A2').value).toBe('2');
      expect(sheet.cell('A3').value).toBe('10');
    });
  });

  describe('case sensitive sort', () => {
    it('should be case insensitive by default', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'banana';
      sheet.cell('A2').value = 'Apple';
      sheet.cell('A3').value = 'cherry';

      sheet.sort('A');

      expect(sheet.cell('A1').value).toBe('Apple');
      expect(sheet.cell('A2').value).toBe('banana');
      expect(sheet.cell('A3').value).toBe('cherry');
    });
  });

  describe('sortBy multiple columns', () => {
    it('should sort by primary then secondary column', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'B';
      sheet.cell('B1').value = 2;
      sheet.cell('A2').value = 'A';
      sheet.cell('B2').value = 2;
      sheet.cell('A3').value = 'A';
      sheet.cell('B3').value = 1;
      sheet.cell('A4').value = 'B';
      sheet.cell('B4').value = 1;

      sheet.sortBy([{ column: 'A' }, { column: 'B' }]);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe(1);
      expect(sheet.cell('A2').value).toBe('A');
      expect(sheet.cell('B2').value).toBe(2);
      expect(sheet.cell('A3').value).toBe('B');
      expect(sheet.cell('B3').value).toBe(1);
      expect(sheet.cell('A4').value).toBe('B');
      expect(sheet.cell('B4').value).toBe(2);
    });

    it('should support mixed ascending/descending in multi-column sort', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 1;
      sheet.cell('A2').value = 'A';
      sheet.cell('B2').value = 2;
      sheet.cell('A3').value = 'B';
      sheet.cell('B3').value = 1;

      sheet.sortBy([{ column: 'A' }, { column: 'B', descending: true }]);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe(2); // Descending
      expect(sheet.cell('A2').value).toBe('A');
      expect(sheet.cell('B2').value).toBe(1);
    });
  });

  describe('sort range', () => {
    it('should only sort specified range', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'C';
      sheet.cell('A2').value = 'B';
      sheet.cell('A3').value = 'A';
      sheet.cell('A4').value = 'Z'; // Outside range

      sheet.sort('A', { range: 'A1:A3' });

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('B');
      expect(sheet.cell('A3').value).toBe('C');
      expect(sheet.cell('A4').value).toBe('Z'); // Unchanged
    });
  });

  describe('sort with dates', () => {
    it('should sort date values correctly', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = new Date('2024-03-01');
      sheet.cell('A2').value = new Date('2024-01-01');
      sheet.cell('A3').value = new Date('2024-02-01');

      sheet.sort('A');

      expect((sheet.cell('A1').value as Date).getMonth()).toBe(0); // January
      expect((sheet.cell('A2').value as Date).getMonth()).toBe(1); // February
      expect((sheet.cell('A3').value as Date).getMonth()).toBe(2); // March
    });
  });

  describe('sort on empty sheet', () => {
    it('should handle empty sheet gracefully', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      // Should not throw
      sheet.sort('A');

      expect(sheet.dimensions).toBeNull();
    });
  });
});

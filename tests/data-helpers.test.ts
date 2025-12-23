import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Sheet Data Import/Export Helpers', () => {
  describe('fromArray', () => {
    it('should populate sheet from 2D array', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.fromArray([
        ['Name', 'Age'],
        ['Alice', 25],
        ['Bob', 30],
      ]);

      expect(sheet.cell('A1').value).toBe('Name');
      expect(sheet.cell('B1').value).toBe('Age');
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('B2').value).toBe(25);
    });

    it('should support custom start position', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.fromArray([['Hello']], { startRow: 2, startCol: 3 });

      expect(sheet.cell('D3').value).toBe('Hello');
    });

    it('should apply header style', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.fromArray(
        [['Header1', 'Header2'], ['Data1', 'Data2']],
        { headers: true, headerStyle: { font: { bold: true } } }
      );

      expect(sheet.cell('A1').style?.font?.bold).toBe(true);
      expect(sheet.cell('A2').style).toBeUndefined();
    });
  });

  describe('fromObjects', () => {
    it('should populate sheet from array of objects', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.fromObjects([
        { name: 'Alice', age: 25 },
        { name: 'Bob', age: 30 },
      ]);

      expect(sheet.cell('A1').value).toBe('name');
      expect(sheet.cell('B1').value).toBe('age');
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('B2').value).toBe(25);
    });

    it('should skip headers when includeHeaders is false', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.fromObjects(
        [{ name: 'Alice', age: 25 }],
        { includeHeaders: false }
      );

      expect(sheet.cell('A1').value).toBe('Alice');
    });

    it('should use specified columns', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.fromObjects(
        [{ name: 'Alice', age: 25, city: 'NYC' }],
        { columns: ['name', 'city'] }
      );

      expect(sheet.cell('A1').value).toBe('name');
      expect(sheet.cell('B1').value).toBe('city');
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('B2').value).toBe('NYC');
    });

    it('should handle empty array', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.fromObjects([]);

      expect(sheet.dimensions).toBeNull();
    });
  });

  describe('toArray', () => {
    it('should export sheet data as 2D array', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Name';
      sheet.cell('B1').value = 'Age';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 25;

      const data = sheet.toArray();

      expect(data).toEqual([
        ['Name', 'Age'],
        ['Alice', 25],
      ]);
    });

    it('should export specific range', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('A2').value = 'C';
      sheet.cell('B2').value = 'D';

      const data = sheet.toArray({ range: 'A1:A2' });

      expect(data).toEqual([['A'], ['C']]);
    });

    it('should handle empty sheet', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      const data = sheet.toArray();

      expect(data).toEqual([]);
    });
  });

  describe('toObjects', () => {
    it('should export sheet data as array of objects', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Name';
      sheet.cell('B1').value = 'Age';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 25;
      sheet.cell('A3').value = 'Bob';
      sheet.cell('B3').value = 30;

      const data = sheet.toObjects();

      expect(data).toEqual([
        { Name: 'Alice', Age: 25 },
        { Name: 'Bob', Age: 30 },
      ]);
    });

    it('should handle missing values', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Name';
      sheet.cell('B1').value = 'Age';
      sheet.cell('A2').value = 'Alice';
      // B2 is missing

      const data = sheet.toObjects();

      expect(data[0]).toEqual({ Name: 'Alice', Age: null });
    });

    it('should handle empty sheet', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      const data = sheet.toObjects();

      expect(data).toEqual([]);
    });
  });

  describe('appendRow', () => {
    it('should append row to end of sheet', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';

      const rowIndex = sheet.appendRow(['Second', 'Data']);

      expect(rowIndex).toBe(1);
      expect(sheet.cell('A2').value).toBe('Second');
      expect(sheet.cell('B2').value).toBe('Data');
    });

    it('should append to empty sheet', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      const rowIndex = sheet.appendRow(['First']);

      expect(rowIndex).toBe(0);
      expect(sheet.cell('A1').value).toBe('First');
    });
  });

  describe('appendRows', () => {
    it('should append multiple rows', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Header';

      const startIndex = sheet.appendRows([
        ['Row 1', 'Data 1'],
        ['Row 2', 'Data 2'],
      ]);

      expect(startIndex).toBe(1);
      expect(sheet.cell('A2').value).toBe('Row 1');
      expect(sheet.cell('A3').value).toBe('Row 2');
    });
  });

  describe('round-trip', () => {
    it('should preserve data through fromArray -> toArray', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      const original = [
        ['Name', 'Age', 'Active'],
        ['Alice', 25, true],
        ['Bob', 30, false],
      ];

      sheet.fromArray(original);
      const exported = sheet.toArray();

      expect(exported).toEqual(original);
    });

    it('should preserve data through fromObjects -> toObjects', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      const original = [
        { Name: 'Alice', Age: 25 },
        { Name: 'Bob', Age: 30 },
      ];

      sheet.fromObjects(original, { includeHeaders: true });
      const exported = sheet.toObjects();

      expect(exported).toEqual(original);
    });
  });
});

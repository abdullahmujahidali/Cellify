import { describe, it, expect } from 'vitest';
import { Sheet } from '../src/core/Sheet.js';
import {
  sheetToCsv,
  sheetToCsvBuffer,
  sheetsToCsv,
  csvToWorkbook,
  csvToSheet,
  csvBufferToWorkbook,
} from '../src/formats/index.js';

describe('CSV Export', () => {
  describe('sheetToCsv', () => {
    it('should export empty sheet', () => {
      const sheet = new Sheet('Test');
      const csv = sheetToCsv(sheet);
      expect(csv).toBe('');
    });

    it('should export single cell', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Hello';
      const csv = sheetToCsv(sheet);
      expect(csv).toBe('Hello');
    });

    it('should export multiple cells', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Name';
      sheet.cell('B1').value = 'Age';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 30;

      const csv = sheetToCsv(sheet);
      expect(csv).toBe('Name,Age\r\nAlice,30');
    });

    it('should handle null values', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'A';
      sheet.cell('C1').value = 'C';

      const csv = sheetToCsv(sheet);
      expect(csv).toBe('A,,C');
    });

    it('should escape fields with commas', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Hello, World';

      const csv = sheetToCsv(sheet);
      expect(csv).toBe('"Hello, World"');
    });

    it('should escape fields with quotes', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Say "Hello"';

      const csv = sheetToCsv(sheet);
      expect(csv).toBe('"Say ""Hello"""');
    });

    it('should escape fields with newlines', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Line1\nLine2';

      const csv = sheetToCsv(sheet);
      expect(csv).toBe('"Line1\nLine2"');
    });

    it('should use custom delimiter', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';

      const csv = sheetToCsv(sheet, { delimiter: ';' });
      expect(csv).toBe('A;B');
    });

    it('should use custom row delimiter', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';

      const csv = sheetToCsv(sheet, { rowDelimiter: '\n' });
      expect(csv).toBe('A\nB');
    });

    it('should quote all fields when requested', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Simple';

      const csv = sheetToCsv(sheet, { quoteAllFields: true });
      expect(csv).toBe('"Simple"');
    });

    it('should include BOM when requested', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Test';

      const csv = sheetToCsv(sheet, { includeBom: true });
      expect(csv.charCodeAt(0)).toBe(0xfeff);
      expect(csv.slice(1)).toBe('Test');
    });

    it('should format boolean values', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = true;
      sheet.cell('B1').value = false;

      const csv = sheetToCsv(sheet);
      expect(csv).toBe('TRUE,FALSE');
    });

    it('should format date values in ISO format', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = new Date('2024-06-15');

      const csv = sheetToCsv(sheet, { dateFormat: 'ISO' });
      expect(csv).toBe('2024-06-15');
    });

    it('should export specific range', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'A1';
      sheet.cell('B1').value = 'B1';
      sheet.cell('A2').value = 'A2';
      sheet.cell('B2').value = 'B2';
      sheet.cell('C3').value = 'C3';

      const csv = sheetToCsv(sheet, { range: 'A1:B2' });
      expect(csv).toBe('A1,B1\r\nA2,B2');
    });
  });

  describe('sheetToCsvBuffer', () => {
    it('should return Uint8Array', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Hello';

      const buffer = sheetToCsvBuffer(sheet);
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(new TextDecoder().decode(buffer)).toBe('Hello');
    });
  });

  describe('sheetsToCsv', () => {
    it('should export multiple sheets', () => {
      const sheet1 = new Sheet('Sheet1');
      sheet1.cell('A1').value = 'Data1';

      const sheet2 = new Sheet('Sheet2');
      sheet2.cell('A1').value = 'Data2';

      const result = sheetsToCsv([sheet1, sheet2]);

      expect(result.get('Sheet1')).toBe('Data1');
      expect(result.get('Sheet2')).toBe('Data2');
    });
  });
});

describe('CSV Import', () => {
  describe('csvToWorkbook', () => {
    it('should import simple CSV', () => {
      const csv = 'A,B,C';
      const workbook = csvToWorkbook(csv);

      expect(workbook.sheetCount).toBe(1);
      const sheet = workbook.getSheetByIndex(0)!;
      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('B');
      expect(sheet.cell('C1').value).toBe('C');
    });

    it('should import multi-row CSV', () => {
      const csv = 'Name,Age\nAlice,30\nBob,25';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('Name');
      expect(sheet.cell('B1').value).toBe('Age');
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('B2').value).toBe(30);
    });

    it('should use custom sheet name', () => {
      const csv = 'Test';
      const workbook = csvToWorkbook(csv, { sheetName: 'MyData' });

      expect(workbook.getSheet('MyData')).toBeDefined();
    });

    it('should handle quoted fields', () => {
      const csv = '"Hello, World","Test"';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('Hello, World');
      expect(sheet.cell('B1').value).toBe('Test');
    });

    it('should handle escaped quotes', () => {
      const csv = '"Say ""Hello"""';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('Say "Hello"');
    });

    it('should handle multi-line fields', () => {
      const csv = '"Line1\nLine2"';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('Line1\nLine2');
    });

    it('should detect numbers', () => {
      const csv = '42,3.14,-100,1000';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe(42);
      expect(sheet.cell('B1').value).toBe(3.14);
      expect(sheet.cell('C1').value).toBe(-100);
      expect(sheet.cell('D1').value).toBe(1000);
    });

    it('should detect percentages', () => {
      const csv = '50%,25%';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe(0.5);
      expect(sheet.cell('B1').value).toBe(0.25);
    });

    it('should detect booleans', () => {
      const csv = 'true,false,TRUE,FALSE';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe(true);
      expect(sheet.cell('B1').value).toBe(false);
      expect(sheet.cell('C1').value).toBe(true);
      expect(sheet.cell('D1').value).toBe(false);
    });

    it('should skip empty lines', () => {
      const csv = 'A\n\nB\n\nC';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('B');
      expect(sheet.cell('A3').value).toBe('C');
    });

    it('should handle Windows line endings', () => {
      const csv = 'A\r\nB\r\nC';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('B');
      expect(sheet.cell('A3').value).toBe('C');
    });

    it('should trim values when requested', () => {
      const csv = '  A  ,  B  ';
      const workbook = csvToWorkbook(csv, { trimValues: true });
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('B');
    });

    it('should respect maxRows option', () => {
      const csv = 'A\nB\nC\nD\nE';
      const workbook = csvToWorkbook(csv, { maxRows: 3 });
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('B');
      expect(sheet.cell('A3').value).toBe('C');
      expect(sheet.getCell('A4')).toBeUndefined();
    });

    it('should auto-detect semicolon delimiter', () => {
      const csv = 'A;B;C\n1;2;3';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('B');
      expect(sheet.cell('C1').value).toBe('C');
    });

    it('should auto-detect tab delimiter', () => {
      const csv = 'A\tB\tC\n1\t2\t3';
      const workbook = csvToWorkbook(csv);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('B');
      expect(sheet.cell('C1').value).toBe('C');
    });
  });

  describe('csvToSheet', () => {
    it('should import into existing sheet', () => {
      const csv = 'A,B';
      const sheet = new Sheet('Existing');
      sheet.cell('A1').value = 'Original';

      csvToSheet(csv, sheet);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('B');
    });

    it('should import at custom start position', () => {
      const csv = 'A,B';
      const sheet = new Sheet('Test');

      csvToSheet(csv, sheet, { startCell: 'C3' });

      expect(sheet.getCell('A1')).toBeUndefined();
      expect(sheet.cell('C3').value).toBe('A');
      expect(sheet.cell('D3').value).toBe('B');
    });

    it('should return import result', () => {
      const csv = 'A,B,C\n1,2,3\n4,5,6';
      const sheet = new Sheet('Test');

      const result = csvToSheet(csv, sheet);

      expect(result.rowCount).toBe(3);
      expect(result.columnCount).toBe(3);
    });

    it('should extract headers when hasHeaders is true', () => {
      const csv = 'Name,Age,City\nAlice,30,NYC';
      const sheet = new Sheet('Test');

      const result = csvToSheet(csv, sheet, { hasHeaders: true });

      expect(result.headers).toEqual(['Name', 'Age', 'City']);
    });
  });

  describe('csvBufferToWorkbook', () => {
    it('should import from Uint8Array', () => {
      const csv = 'Hello,World';
      const buffer = new TextEncoder().encode(csv);

      const workbook = csvBufferToWorkbook(buffer);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('Hello');
      expect(sheet.cell('B1').value).toBe('World');
    });

    it('should handle BOM in buffer', () => {
      const csv = '\uFEFFHello,World';
      const buffer = new TextEncoder().encode(csv);

      const workbook = csvBufferToWorkbook(buffer);
      const sheet = workbook.getSheetByIndex(0)!;

      expect(sheet.cell('A1').value).toBe('Hello');
    });
  });

  describe('roundtrip', () => {
    it('should preserve data through export/import cycle', () => {
      const original = new Sheet('Test');
      original.cell('A1').value = 'Name';
      original.cell('B1').value = 'Score';
      original.cell('A2').value = 'Alice';
      original.cell('B2').value = 95;
      original.cell('A3').value = 'Bob';
      original.cell('B3').value = 87;

      // Export
      const csv = sheetToCsv(original);

      // Import
      const workbook = csvToWorkbook(csv);
      const imported = workbook.getSheetByIndex(0)!;

      // Verify
      expect(imported.cell('A1').value).toBe('Name');
      expect(imported.cell('B1').value).toBe('Score');
      expect(imported.cell('A2').value).toBe('Alice');
      expect(imported.cell('B2').value).toBe(95);
      expect(imported.cell('A3').value).toBe('Bob');
      expect(imported.cell('B3').value).toBe(87);
    });

    it('should preserve special characters through roundtrip', () => {
      const original = new Sheet('Test');
      original.cell('A1').value = 'Hello, "World"';
      original.cell('A2').value = 'Line1\nLine2';

      const csv = sheetToCsv(original);
      const workbook = csvToWorkbook(csv);
      const imported = workbook.getSheetByIndex(0)!;

      expect(imported.cell('A1').value).toBe('Hello, "World"');
      expect(imported.cell('A2').value).toBe('Line1\nLine2');
    });
  });
});

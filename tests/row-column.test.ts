import { describe, it, expect, beforeEach } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';
import type { Sheet } from '../src/core/Sheet.js';

describe('Sheet Row/Column Operations', () => {
  let sheet: Sheet;

  beforeEach(() => {
    const workbook = new Workbook();
    sheet = workbook.addSheet('Test');
  });

  describe('insertRow', () => {
    it('should insert a single row', () => {
      sheet.cell('A1').value = 'Row 1';
      sheet.cell('A2').value = 'Row 2';
      sheet.cell('A3').value = 'Row 3';

      sheet.insertRow(1);

      expect(sheet.cell('A1').value).toBe('Row 1');
      expect(sheet.cell('A2').value).toBeNull();
      expect(sheet.cell('A3').value).toBe('Row 2');
      expect(sheet.cell('A4').value).toBe('Row 3');
    });

    it('should insert multiple rows', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';

      sheet.insertRow(1, 3);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A5').value).toBe('B');
    });

    it('should insert at the beginning', () => {
      sheet.cell('A1').value = 'First';

      sheet.insertRow(0, 2);

      expect(sheet.cell('A3').value).toBe('First');
    });

    it('should preserve cell styles', () => {
      sheet.cell('A1').value = 'Styled';
      sheet.cell('A1').style = { font: { bold: true } };

      sheet.insertRow(0);

      expect(sheet.cell('A2').value).toBe('Styled');
      expect(sheet.cell('A2').style?.font?.bold).toBe(true);
    });

    it('should shift row configurations', () => {
      sheet.setRowHeight(1, 30);

      sheet.insertRow(0);

      expect(sheet.getRow(2).height).toBe(30);
    });
  });

  describe('insertColumn', () => {
    it('should insert a single column', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.insertColumn(1);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBeNull();
      expect(sheet.cell('C1').value).toBe('B');
      expect(sheet.cell('D1').value).toBe('C');
    });

    it('should insert multiple columns', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';

      sheet.insertColumn(1, 2);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('D1').value).toBe('B');
    });
  });

  describe('deleteRow', () => {
    it('should delete a single row', () => {
      sheet.cell('A1').value = 'Row 1';
      sheet.cell('A2').value = 'Row 2';
      sheet.cell('A3').value = 'Row 3';

      sheet.deleteRow(1);

      expect(sheet.cell('A1').value).toBe('Row 1');
      expect(sheet.cell('A2').value).toBe('Row 3');
    });

    it('should delete multiple rows', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';
      sheet.cell('A3').value = 'C';
      sheet.cell('A4').value = 'D';

      sheet.deleteRow(1, 2);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('D');
    });

    it('should delete first row', () => {
      sheet.cell('A1').value = 'First';
      sheet.cell('A2').value = 'Second';

      sheet.deleteRow(0);

      expect(sheet.cell('A1').value).toBe('Second');
    });

    it('should preserve data in other columns', () => {
      sheet.cell('A1').value = 'A1';
      sheet.cell('B1').value = 'B1';
      sheet.cell('A2').value = 'A2';
      sheet.cell('B2').value = 'B2';

      sheet.deleteRow(0);

      expect(sheet.cell('A1').value).toBe('A2');
      expect(sheet.cell('B1').value).toBe('B2');
    });
  });

  describe('deleteColumn', () => {
    it('should delete a single column', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.deleteColumn(1);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('C');
    });

    it('should delete multiple columns', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';
      sheet.cell('D1').value = 'D';

      sheet.deleteColumn(1, 2);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('D');
    });
  });

  describe('moveRow', () => {
    it('should move row down', () => {
      sheet.cell('A1').value = 'First';
      sheet.cell('A2').value = 'Second';
      sheet.cell('A3').value = 'Third';

      sheet.moveRow(0, 2);

      expect(sheet.cell('A1').value).toBe('Second');
      expect(sheet.cell('A2').value).toBe('First');
      expect(sheet.cell('A3').value).toBe('Third');
    });

    it('should move row up', () => {
      sheet.cell('A1').value = 'First';
      sheet.cell('A2').value = 'Second';
      sheet.cell('A3').value = 'Third';

      sheet.moveRow(2, 0);

      expect(sheet.cell('A1').value).toBe('Third');
      expect(sheet.cell('A2').value).toBe('First');
      expect(sheet.cell('A3').value).toBe('Second');
    });

    it('should preserve styles when moving', () => {
      sheet.cell('A1').value = 'Styled';
      sheet.cell('A1').style = { font: { bold: true } };
      sheet.cell('A2').value = 'Plain';

      sheet.moveRow(0, 2);

      expect(sheet.cell('A2').style?.font?.bold).toBe(true);
    });
  });

  describe('moveColumn', () => {
    it('should move column right', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.moveColumn(0, 2);

      expect(sheet.cell('A1').value).toBe('B');
      expect(sheet.cell('B1').value).toBe('A');
      expect(sheet.cell('C1').value).toBe('C');
    });

    it('should move column left', () => {
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.moveColumn(2, 0);

      expect(sheet.cell('A1').value).toBe('C');
      expect(sheet.cell('B1').value).toBe('A');
      expect(sheet.cell('C1').value).toBe('B');
    });
  });

  describe('edge cases', () => {
    it('should handle insert with count 0', () => {
      sheet.cell('A1').value = 'Test';

      sheet.insertRow(0, 0);

      expect(sheet.cell('A1').value).toBe('Test');
    });

    it('should handle delete with count 0', () => {
      sheet.cell('A1').value = 'Test';

      sheet.deleteRow(0, 0);

      expect(sheet.cell('A1').value).toBe('Test');
    });

    it('should handle move to same position', () => {
      sheet.cell('A1').value = 'Test';

      sheet.moveRow(0, 0);

      expect(sheet.cell('A1').value).toBe('Test');
    });
  });
});

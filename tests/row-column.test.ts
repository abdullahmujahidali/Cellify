import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Sheet Row/Column Operations', () => {
  describe('insertRow', () => {
    it('should insert a single row', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Row 1';
      sheet.cell('A2').value = 'Row 2';
      sheet.cell('A3').value = 'Row 3';

      sheet.insertRow(1); // Insert at row index 1

      expect(sheet.cell('A1').value).toBe('Row 1');
      expect(sheet.cell('A2').value).toBeNull(); // New empty row
      expect(sheet.cell('A3').value).toBe('Row 2');
      expect(sheet.cell('A4').value).toBe('Row 3');
    });

    it('should insert multiple rows', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';

      sheet.insertRow(1, 3); // Insert 3 rows at index 1

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A5').value).toBe('B'); // Shifted by 3
    });

    it('should insert at the beginning', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';

      sheet.insertRow(0, 2);

      expect(sheet.cell('A3').value).toBe('First');
    });

    it('should preserve cell styles', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Styled';
      sheet.cell('A1').style = { font: { bold: true } };

      sheet.insertRow(0);

      expect(sheet.cell('A2').value).toBe('Styled');
      expect(sheet.cell('A2').style?.font?.bold).toBe(true);
    });

    it('should shift row configurations', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.setRowHeight(1, 30);

      sheet.insertRow(0);

      expect(sheet.getRow(2).height).toBe(30);
    });
  });

  describe('insertColumn', () => {
    it('should insert a single column', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.insertColumn(1); // Insert at column B

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBeNull(); // New empty column
      expect(sheet.cell('C1').value).toBe('B');
      expect(sheet.cell('D1').value).toBe('C');
    });

    it('should insert multiple columns', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';

      sheet.insertColumn(1, 2);

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('D1').value).toBe('B'); // Shifted by 2
    });
  });

  describe('deleteRow', () => {
    it('should delete a single row', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Row 1';
      sheet.cell('A2').value = 'Row 2';
      sheet.cell('A3').value = 'Row 3';

      sheet.deleteRow(1); // Delete row at index 1

      expect(sheet.cell('A1').value).toBe('Row 1');
      expect(sheet.cell('A2').value).toBe('Row 3'); // Shifted up
    });

    it('should delete multiple rows', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';
      sheet.cell('A3').value = 'C';
      sheet.cell('A4').value = 'D';

      sheet.deleteRow(1, 2); // Delete rows 1 and 2

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('A2').value).toBe('D');
    });

    it('should delete first row', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A2').value = 'Second';

      sheet.deleteRow(0);

      expect(sheet.cell('A1').value).toBe('Second');
    });

    it('should preserve data in other columns', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

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
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.deleteColumn(1); // Delete column B

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('C'); // Shifted left
    });

    it('should delete multiple columns', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';
      sheet.cell('D1').value = 'D';

      sheet.deleteColumn(1, 2); // Delete columns B and C

      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('D');
    });
  });

  describe('moveRow', () => {
    it('should move row down', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A2').value = 'Second';
      sheet.cell('A3').value = 'Third';

      sheet.moveRow(0, 2); // Move first row to position 2

      expect(sheet.cell('A1').value).toBe('Second');
      expect(sheet.cell('A2').value).toBe('First');
      expect(sheet.cell('A3').value).toBe('Third');
    });

    it('should move row up', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A2').value = 'Second';
      sheet.cell('A3').value = 'Third';

      sheet.moveRow(2, 0); // Move third row to position 0

      expect(sheet.cell('A1').value).toBe('Third');
      expect(sheet.cell('A2').value).toBe('First');
      expect(sheet.cell('A3').value).toBe('Second');
    });

    it('should preserve styles when moving', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Styled';
      sheet.cell('A1').style = { font: { bold: true } };
      sheet.cell('A2').value = 'Plain';

      sheet.moveRow(0, 2);

      expect(sheet.cell('A2').style?.font?.bold).toBe(true);
    });
  });

  describe('moveColumn', () => {
    it('should move column right', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.moveColumn(0, 2); // Move column A to position C

      expect(sheet.cell('A1').value).toBe('B');
      expect(sheet.cell('B1').value).toBe('A');
      expect(sheet.cell('C1').value).toBe('C');
    });

    it('should move column left', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.moveColumn(2, 0); // Move column C to position A

      expect(sheet.cell('A1').value).toBe('C');
      expect(sheet.cell('B1').value).toBe('A');
      expect(sheet.cell('C1').value).toBe('B');
    });
  });

  describe('edge cases', () => {
    it('should handle insert with count 0', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';

      sheet.insertRow(0, 0);

      expect(sheet.cell('A1').value).toBe('Test');
    });

    it('should handle delete with count 0', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';

      sheet.deleteRow(0, 0);

      expect(sheet.cell('A1').value).toBe('Test');
    });

    it('should handle move to same position', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';

      sheet.moveRow(0, 0);

      expect(sheet.cell('A1').value).toBe('Test');
    });
  });
});

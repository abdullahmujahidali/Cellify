import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Sheet Copy/Paste', () => {
  describe('copyRange', () => {
    it('should copy range to clipboard', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.cell('A2').value = 'World';

      sheet.copyRange('A1:A2');

      expect(sheet.hasClipboard).toBe(true);
    });

    it('should copy values and styles', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';
      sheet.cell('A1').style = { font: { bold: true } };

      sheet.copyRange('A1');
      sheet.pasteRange('B1');

      expect(sheet.cell('B1').value).toBe('Test');
      expect(sheet.cell('B1').style?.font?.bold).toBe(true);
    });
  });

  describe('pasteRange', () => {
    it('should paste at specified location', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('A2').value = 'B';
      sheet.cell('B1').value = 1;
      sheet.cell('B2').value = 2;

      sheet.copyRange('A1:B2');
      sheet.pasteRange('D1');

      expect(sheet.cell('D1').value).toBe('A');
      expect(sheet.cell('D2').value).toBe('B');
      expect(sheet.cell('E1').value).toBe(1);
      expect(sheet.cell('E2').value).toBe(2);
    });

    it('should paste values only when valuesOnly is true', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';
      sheet.cell('A1').style = { font: { bold: true } };

      sheet.copyRange('A1');
      sheet.pasteRange('B1', { valuesOnly: true });

      expect(sheet.cell('B1').value).toBe('Test');
      expect(sheet.cell('B1').style).toBeUndefined();
    });

    it('should paste styles only when stylesOnly is true', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';
      sheet.cell('A1').style = { font: { bold: true } };
      sheet.cell('B1').value = 'Original';

      sheet.copyRange('A1');
      sheet.pasteRange('B1', { stylesOnly: true });

      expect(sheet.cell('B1').value).toBe('Original'); // Unchanged
      expect(sheet.cell('B1').style?.font?.bold).toBe(true);
    });

    it('should transpose when transpose is true', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      // Create a row: A1=1, B1=2, C1=3
      sheet.cell('A1').value = 1;
      sheet.cell('B1').value = 2;
      sheet.cell('C1').value = 3;

      sheet.copyRange('A1:C1');
      sheet.pasteRange('E1', { transpose: true });

      // Should become a column: E1=1, E2=2, E3=3
      expect(sheet.cell('E1').value).toBe(1);
      expect(sheet.cell('E2').value).toBe(2);
      expect(sheet.cell('E3').value).toBe(3);
    });

    it('should do nothing if clipboard is empty', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('B1').value = 'Original';

      // No copy, just paste
      sheet.pasteRange('B1');

      expect(sheet.cell('B1').value).toBe('Original');
    });

    it('should paste using row/col object', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';
      sheet.copyRange('A1');
      sheet.pasteRange({ row: 2, col: 3 }); // D3

      expect(sheet.cell('D3').value).toBe('Test');
    });
  });

  describe('cutRange', () => {
    it('should cut and clear original cells', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Cut me';
      sheet.cell('A1').style = { font: { bold: true } };

      sheet.cutRange('A1');
      sheet.pasteRange('B1');

      expect(sheet.cell('A1').value).toBeNull(); // Cleared
      expect(sheet.cell('B1').value).toBe('Cut me');
      expect(sheet.cell('B1').style?.font?.bold).toBe(true);
    });
  });

  describe('duplicateRange', () => {
    it('should copy and paste in one operation', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.cell('A2').value = 'World';

      sheet.duplicateRange('A1:A2', 'C1');

      expect(sheet.cell('C1').value).toBe('Hello');
      expect(sheet.cell('C2').value).toBe('World');
      // Original still exists
      expect(sheet.cell('A1').value).toBe('Hello');
    });
  });

  describe('clearClipboard', () => {
    it('should clear the clipboard', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';
      sheet.copyRange('A1');

      expect(sheet.hasClipboard).toBe(true);

      sheet.clearClipboard();

      expect(sheet.hasClipboard).toBe(false);
    });
  });

  describe('copy with formulas', () => {
    it('should copy formulas', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').setFormula('=SUM(B1:B10)', 100);

      sheet.copyRange('A1');
      sheet.pasteRange('C1');

      expect(sheet.cell('C1').formula?.formula).toBe('SUM(B1:B10)');
    });
  });
});

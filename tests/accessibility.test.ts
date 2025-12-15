import { describe, it, expect } from 'vitest';
import { Cell } from '../src/core/Cell.js';
import { Sheet } from '../src/core/Sheet.js';
import {
  getCellAccessibility,
  getValueText,
  getSheetAccessibility,
  describeCellPosition,
  describeCellFull,
  createAnnouncement,
  announceNavigation,
  announceSelection,
  announceError,
  announceSuccess,
  getAriaAttributes,
} from '../src/accessibility/helpers.js';

describe('Accessibility Helpers', () => {
  describe('getCellAccessibility', () => {
    it('should identify column header cells', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A1');
      cell.value = 'Name';

      const a11y = getCellAccessibility(cell, sheet, { headerRows: 1 });

      expect(a11y.isHeader).toBe(true);
      expect(a11y.role).toBe('columnheader');
      expect(a11y.scope).toBe('col');
    });

    it('should identify row header cells', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A2');
      cell.value = 'Row 1';

      const a11y = getCellAccessibility(cell, sheet, { headerCols: 1 });

      expect(a11y.isHeader).toBe(true);
      expect(a11y.role).toBe('rowheader');
      expect(a11y.scope).toBe('row');
    });

    it('should identify data cells', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('B2');
      cell.value = 'Data';

      const a11y = getCellAccessibility(cell, sheet, { headerRows: 1, headerCols: 1 });

      expect(a11y.isHeader).toBe(false);
      expect(a11y.role).toBe('gridcell');
      expect(a11y.scope).toBeUndefined();
    });

    it('should generate header references for data cells', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 'Col Header';
      sheet.cell('A2').value = 'Row Header';
      const cell = sheet.cell('B2');
      cell.value = 'Data';

      const a11y = getCellAccessibility(cell, sheet, { headerRows: 1, headerCols: 1 });

      expect(a11y.headers).toContain('cell-0-1'); // Column header
      expect(a11y.headers).toContain('cell-1-0'); // Row header
    });

    it('should include position information', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('C3');

      const a11y = getCellAccessibility(cell, sheet, { includePosition: true });

      expect(a11y.ariaColIndex).toBe(3); // 1-based
      expect(a11y.ariaRowIndex).toBe(3); // 1-based
    });

    it('should handle merged cells', () => {
      const sheet = new Sheet('Test');
      sheet.mergeCells('A1:C3');
      const cell = sheet.cell('A1');

      const a11y = getCellAccessibility(cell, sheet);

      expect(a11y.ariaColSpan).toBe(3);
      expect(a11y.ariaRowSpan).toBe(3);
    });

    it('should mark cells as read-only when sheet is protected', () => {
      const sheet = new Sheet('Test');
      sheet.protect();
      const cell = sheet.cell('A1');
      cell.value = 'Protected';

      const a11y = getCellAccessibility(cell, sheet);

      expect(a11y.ariaReadOnly).toBe(true);
    });

    it('should detect list validation with dropdown', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A1');
      cell.setValidation({
        type: 'list',
        formula1: 'A,B,C',
        showDropDown: true,
      });

      const a11y = getCellAccessibility(cell, sheet);

      expect(a11y.ariaHasPopup).toBe('listbox');
    });
  });

  describe('getValueText', () => {
    it('should return "empty" for null values', () => {
      const cell = new Cell(0, 0);
      expect(getValueText(cell)).toBe('empty');
    });

    it('should return string values as-is', () => {
      const cell = new Cell(0, 0, 'Hello');
      expect(getValueText(cell)).toBe('Hello');
    });

    it('should return boolean as text', () => {
      const cell = new Cell(0, 0, true);
      expect(getValueText(cell)).toBe('true');
    });

    it('should format dates', () => {
      const cell = new Cell(0, 0, new Date('2024-01-15'));
      const text = getValueText(cell);
      expect(text).toContain('2024'); // Date format varies by locale
    });

    it('should describe error values', () => {
      const cell = new Cell(0, 0, '#DIV/0!');
      expect(getValueText(cell)).toBe('division by zero error');
    });

    it('should describe percentage values', () => {
      const cell = new Cell(0, 0, 0.25);
      cell.style = { numberFormat: { formatCode: '0%' } };
      expect(getValueText(cell)).toBe('25 percent');
    });

    it('should describe currency values', () => {
      const cell = new Cell(0, 0, 100);
      cell.style = { numberFormat: { formatCode: '$#,##0.00' } };
      expect(getValueText(cell)).toBe('100.00 dollars');
    });
  });

  describe('getSheetAccessibility', () => {
    it('should include sheet label', () => {
      const sheet = new Sheet('Sales Data');
      const a11y = getSheetAccessibility(sheet);

      expect(a11y.label).toBe('Sales Data');
    });

    it('should include dimensions', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 1;
      sheet.cell('C3').value = 2;

      const a11y = getSheetAccessibility(sheet);

      expect(a11y.ariaRowCount).toBe(3);
      expect(a11y.ariaColCount).toBe(3);
    });

    it('should include header configuration', () => {
      const sheet = new Sheet('Test');
      const a11y = getSheetAccessibility(sheet, { headerRows: 2, headerCols: 1 });

      expect(a11y.headerRowStart).toBe(0);
      expect(a11y.headerRowEnd).toBe(1);
      expect(a11y.headerColStart).toBe(0);
      expect(a11y.headerColEnd).toBe(0);
    });
  });

  describe('describeCellPosition', () => {
    it('should describe cell position in human-readable format', () => {
      expect(describeCellPosition(0, 0)).toBe('Cell A1, row 1, column 1');
      expect(describeCellPosition(4, 2)).toBe('Cell C5, row 5, column 3');
    });
  });

  describe('describeCellFull', () => {
    it('should describe cell with value', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A1');
      cell.value = 'Hello';

      const description = describeCellFull(cell, sheet);

      expect(description).toContain('Cell A1');
      expect(description).toContain('Hello');
    });

    it('should describe empty cell', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A1');

      const description = describeCellFull(cell, sheet);

      expect(description).toContain('empty');
    });

    it('should describe merged cell', () => {
      const sheet = new Sheet('Test');
      sheet.mergeCells('A1:B2');
      const cell = sheet.cell('A1');
      cell.value = 'Merged';

      const description = describeCellFull(cell, sheet);

      expect(description).toContain('merged 2 rows by 2 columns');
    });

    it('should mention formula', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A1');
      cell.setFormula('=SUM(B1:B10)');

      const description = describeCellFull(cell, sheet);

      expect(description).toContain('formula');
      expect(description).toContain('SUM(B1:B10)');
    });

    it('should mention comment', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A1');
      cell.value = 'Data';
      cell.setComment('Important note');

      const description = describeCellFull(cell, sheet);

      expect(description).toContain('has comment');
    });
  });

  describe('announcements', () => {
    it('should create navigation announcement', () => {
      const cell = new Cell(0, 0, 'Hello');
      const announcement = announceNavigation(cell);

      expect(announcement.type).toBe('navigation');
      expect(announcement.priority).toBe('polite');
      expect(announcement.message).toContain('A1');
      expect(announcement.message).toContain('Hello');
    });

    it('should create selection announcement for single cell', () => {
      const announcement = announceSelection(0, 0, 0, 0);

      expect(announcement.type).toBe('selection');
      expect(announcement.message).toContain('Selected cell A1');
    });

    it('should create selection announcement for range', () => {
      const announcement = announceSelection(0, 0, 2, 3);

      expect(announcement.type).toBe('selection');
      expect(announcement.message).toContain('A1');
      expect(announcement.message).toContain('D3');
      expect(announcement.message).toContain('3 rows by 4 columns');
    });

    it('should create error announcement with assertive priority', () => {
      const announcement = announceError('Invalid input');

      expect(announcement.type).toBe('error');
      expect(announcement.priority).toBe('assertive');
      expect(announcement.message).toBe('Invalid input');
    });

    it('should create success announcement', () => {
      const announcement = announceSuccess('Saved successfully');

      expect(announcement.type).toBe('success');
      expect(announcement.priority).toBe('polite');
    });
  });

  describe('getAriaAttributes', () => {
    it('should generate ARIA attributes object', () => {
      const a11y = {
        role: 'gridcell' as const,
        ariaColIndex: 1,
        ariaRowIndex: 1,
        ariaSelected: true,
        ariaReadOnly: false,
      };

      const attrs = getAriaAttributes(a11y);

      expect(attrs['role']).toBe('gridcell');
      expect(attrs['aria-colindex']).toBe(1);
      expect(attrs['aria-rowindex']).toBe(1);
      expect(attrs['aria-selected']).toBe(true);
      expect(attrs['aria-readonly']).toBe(false);
    });

    it('should handle merged cell spans', () => {
      const a11y = {
        ariaColSpan: 3,
        ariaRowSpan: 2,
      };

      const attrs = getAriaAttributes(a11y);

      expect(attrs['aria-colspan']).toBe(3);
      expect(attrs['aria-rowspan']).toBe(2);
    });

    it('should join header references', () => {
      const a11y = {
        headers: ['cell-0-1', 'cell-1-0'],
      };

      const attrs = getAriaAttributes(a11y);

      expect(attrs['headers']).toBe('cell-0-1 cell-1-0');
    });
  });
});

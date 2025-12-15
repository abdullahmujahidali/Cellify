import { describe, it, expect } from 'vitest';
import { Sheet } from '../src/core/Sheet.js';

describe('Sheet', () => {
  describe('constructor', () => {
    it('should create a sheet with a name', () => {
      const sheet = new Sheet('TestSheet');
      expect(sheet.name).toBe('TestSheet');
    });
  });

  describe('cell access', () => {
    it('should get cell by A1 notation', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell('A1');
      expect(cell.row).toBe(0);
      expect(cell.col).toBe(0);
    });

    it('should get cell by row and column', () => {
      const sheet = new Sheet('Test');
      const cell = sheet.cell(0, 0);
      expect(cell.address).toBe('A1');
    });

    it('should return same cell on repeated access', () => {
      const sheet = new Sheet('Test');
      const cell1 = sheet.cell('A1');
      const cell2 = sheet.cell('A1');
      expect(cell1).toBe(cell2);
    });

    it('should return same cell for A1 and (0,0)', () => {
      const sheet = new Sheet('Test');
      const cell1 = sheet.cell('A1');
      const cell2 = sheet.cell(0, 0);
      expect(cell1).toBe(cell2);
    });
  });

  describe('dimensions', () => {
    it('should return null for empty sheet', () => {
      const sheet = new Sheet('Test');
      expect(sheet.dimensions).toBeNull();
    });

    it('should track dimensions as cells are added', () => {
      const sheet = new Sheet('Test');
      sheet.cell('B2').value = 1;
      sheet.cell('D5').value = 2;

      expect(sheet.dimensions).toEqual({
        startRow: 1,
        startCol: 1,
        endRow: 4,
        endCol: 3,
      });
    });

    it('should return correct row and column count', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 1;
      sheet.cell('C3').value = 2;

      expect(sheet.rowCount).toBe(3);
      expect(sheet.columnCount).toBe(3);
    });
  });

  describe('setValues/getValues', () => {
    it('should set values from 2D array', () => {
      const sheet = new Sheet('Test');
      sheet.setValues('A1', [
        [1, 2, 3],
        [4, 5, 6],
      ]);

      expect(sheet.cell('A1').value).toBe(1);
      expect(sheet.cell('C1').value).toBe(3);
      expect(sheet.cell('A2').value).toBe(4);
      expect(sheet.cell('C2').value).toBe(6);
    });

    it('should get values as 2D array', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 1;
      sheet.cell('B1').value = 2;
      sheet.cell('A2').value = 3;
      sheet.cell('B2').value = 4;

      const values = sheet.getValues('A1:B2');
      expect(values).toEqual([
        [1, 2],
        [3, 4],
      ]);
    });

    it('should return null for empty cells in range', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 1;
      sheet.cell('B2').value = 2;

      const values = sheet.getValues('A1:B2');
      expect(values).toEqual([
        [1, null],
        [null, 2],
      ]);
    });
  });

  describe('merge cells', () => {
    it('should merge cells', () => {
      const sheet = new Sheet('Test');
      sheet.mergeCells('A1:B2');

      expect(sheet.merges).toHaveLength(1);
      expect(sheet.cell('A1').isMergeMaster).toBe(true);
      expect(sheet.cell('B1').isMergedSlave).toBe(true);
      expect(sheet.cell('A2').isMergedSlave).toBe(true);
      expect(sheet.cell('B2').isMergedSlave).toBe(true);
    });

    it('should track merged-into reference', () => {
      const sheet = new Sheet('Test');
      sheet.mergeCells('A1:B2');

      expect(sheet.cell('B2').mergedInto).toEqual({ row: 0, col: 0 });
    });

    it('should throw on overlapping merges', () => {
      const sheet = new Sheet('Test');
      sheet.mergeCells('A1:B2');

      expect(() => sheet.mergeCells('B2:C3')).toThrow('overlaps');
    });

    it('should unmerge cells', () => {
      const sheet = new Sheet('Test');
      sheet.mergeCells('A1:B2');
      sheet.unmergeCells('A1:B2');

      expect(sheet.merges).toHaveLength(0);
      expect(sheet.cell('A1').isMergeMaster).toBe(false);
      expect(sheet.cell('B1').isMergedSlave).toBe(false);
    });
  });

  describe('row configuration', () => {
    it('should set row height', () => {
      const sheet = new Sheet('Test');
      sheet.setRowHeight(0, 30);
      expect(sheet.getRow(0).height).toBe(30);
    });

    it('should hide and show rows', () => {
      const sheet = new Sheet('Test');
      sheet.hideRow(0);
      expect(sheet.getRow(0).hidden).toBe(true);

      sheet.showRow(0);
      expect(sheet.getRow(0).hidden).toBe(false);
    });
  });

  describe('column configuration', () => {
    it('should set column width', () => {
      const sheet = new Sheet('Test');
      sheet.setColumnWidth(0, 20);
      expect(sheet.getColumn(0).width).toBe(20);
    });

    it('should hide and show columns', () => {
      const sheet = new Sheet('Test');
      sheet.hideColumn(0);
      expect(sheet.getColumn(0).hidden).toBe(true);

      sheet.showColumn(0);
      expect(sheet.getColumn(0).hidden).toBe(false);
    });
  });

  describe('freeze panes', () => {
    it('should freeze rows and columns', () => {
      const sheet = new Sheet('Test');
      sheet.freeze(2, 1);

      expect(sheet.view.frozenRows).toBe(2);
      expect(sheet.view.frozenCols).toBe(1);
    });

    it('should unfreeze', () => {
      const sheet = new Sheet('Test');
      sheet.freeze(2, 1);
      sheet.unfreeze();

      expect(sheet.view.frozenRows).toBeUndefined();
      expect(sheet.view.frozenCols).toBeUndefined();
    });
  });

  describe('auto filter', () => {
    it('should set auto filter', () => {
      const sheet = new Sheet('Test');
      sheet.setAutoFilter('A1:D10');

      expect(sheet.autoFilter?.range).toEqual({
        startRow: 0,
        startCol: 0,
        endRow: 9,
        endCol: 3,
      });
    });

    it('should remove auto filter', () => {
      const sheet = new Sheet('Test');
      sheet.setAutoFilter('A1:D10');
      sheet.removeAutoFilter();

      expect(sheet.autoFilter).toBeUndefined();
    });
  });

  describe('protection', () => {
    it('should protect sheet', () => {
      const sheet = new Sheet('Test');
      sheet.protect({ password: 'secret' });

      expect(sheet.isProtected).toBe(true);
      expect(sheet.protection?.password).toBe('secret');
    });

    it('should unprotect sheet', () => {
      const sheet = new Sheet('Test');
      sheet.protect();
      sheet.unprotect();

      expect(sheet.isProtected).toBe(false);
    });
  });

  describe('applyStyle', () => {
    it('should apply style to a range', () => {
      const sheet = new Sheet('Test');
      sheet.applyStyle('A1:B2', { font: { bold: true } });

      expect(sheet.cell('A1').style?.font?.bold).toBe(true);
      expect(sheet.cell('B1').style?.font?.bold).toBe(true);
      expect(sheet.cell('A2').style?.font?.bold).toBe(true);
      expect(sheet.cell('B2').style?.font?.bold).toBe(true);
    });
  });

  describe('clearRange', () => {
    it('should clear cells in a range', () => {
      const sheet = new Sheet('Test');
      sheet.setValues('A1', [[1, 2], [3, 4]]);
      sheet.clearRange('A1:B2');

      expect(sheet.cell('A1').value).toBeNull();
      expect(sheet.cell('B2').value).toBeNull();
    });
  });

  describe('iteration', () => {
    it('should iterate over all cells', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 1;
      sheet.cell('B2').value = 2;

      const cells = [...sheet.cells()];
      expect(cells).toHaveLength(2);
    });

    it('should iterate over cells in range', () => {
      const sheet = new Sheet('Test');
      sheet.cell('A1').value = 1;
      sheet.cell('A2').value = 2;
      sheet.cell('B1').value = 3;
      sheet.cell('C1').value = 4;

      const cells = [...sheet.cellsInRange('A1:B2')];
      expect(cells).toHaveLength(3); // A1, A2, B1 (C1 is outside range)
    });
  });
});

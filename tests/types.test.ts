import { describe, it, expect } from 'vitest';
import {
  columnIndexToLetter,
  columnLetterToIndex,
  addressToA1,
  a1ToAddress,
  cellKey,
  parseKey,
  getCellValueType,
} from '../src/types/cell.types.js';
import {
  parseRangeReference,
  rangeToA1,
  rangesOverlap,
  isCellInRange,
  getRangeIntersection,
  getRangeUnion,
  iterateRange,
  getRangeDimensions,
} from '../src/types/range.types.js';
import {
  createSolidFill,
  createBorder,
  createUniformBorders,
} from '../src/types/style.types.js';

describe('Cell Type Utilities', () => {
  describe('columnIndexToLetter', () => {
    it('should convert single letter columns', () => {
      expect(columnIndexToLetter(0)).toBe('A');
      expect(columnIndexToLetter(25)).toBe('Z');
    });

    it('should convert double letter columns', () => {
      expect(columnIndexToLetter(26)).toBe('AA');
      expect(columnIndexToLetter(27)).toBe('AB');
      expect(columnIndexToLetter(51)).toBe('AZ');
      expect(columnIndexToLetter(52)).toBe('BA');
    });

    it('should convert triple letter columns', () => {
      expect(columnIndexToLetter(702)).toBe('AAA');
    });
  });

  describe('columnLetterToIndex', () => {
    it('should convert single letter columns', () => {
      expect(columnLetterToIndex('A')).toBe(0);
      expect(columnLetterToIndex('Z')).toBe(25);
    });

    it('should convert double letter columns', () => {
      expect(columnLetterToIndex('AA')).toBe(26);
      expect(columnLetterToIndex('AB')).toBe(27);
      expect(columnLetterToIndex('AZ')).toBe(51);
      expect(columnLetterToIndex('BA')).toBe(52);
    });

    it('should be case insensitive', () => {
      expect(columnLetterToIndex('a')).toBe(0);
      expect(columnLetterToIndex('aa')).toBe(26);
    });
  });

  describe('addressToA1', () => {
    it('should convert row/col to A1 notation', () => {
      expect(addressToA1(0, 0)).toBe('A1');
      expect(addressToA1(0, 25)).toBe('Z1');
      expect(addressToA1(9, 2)).toBe('C10');
      expect(addressToA1(99, 26)).toBe('AA100');
    });
  });

  describe('a1ToAddress', () => {
    it('should parse A1 notation to row/col', () => {
      expect(a1ToAddress('A1')).toEqual({ row: 0, col: 0 });
      expect(a1ToAddress('Z1')).toEqual({ row: 0, col: 25 });
      expect(a1ToAddress('C10')).toEqual({ row: 9, col: 2 });
      expect(a1ToAddress('AA100')).toEqual({ row: 99, col: 26 });
    });

    it('should handle absolute references', () => {
      expect(a1ToAddress('$A$1')).toEqual({ row: 0, col: 0 });
      expect(a1ToAddress('$B2')).toEqual({ row: 1, col: 1 });
      expect(a1ToAddress('C$3')).toEqual({ row: 2, col: 2 });
    });

    it('should throw on invalid reference', () => {
      expect(() => a1ToAddress('invalid')).toThrow();
      expect(() => a1ToAddress('123')).toThrow();
      expect(() => a1ToAddress('')).toThrow();
    });
  });

  describe('cellKey / parseKey', () => {
    it('should generate and parse cell keys', () => {
      expect(cellKey(0, 0)).toBe('0,0');
      expect(cellKey(10, 5)).toBe('10,5');

      expect(parseKey('0,0')).toEqual({ row: 0, col: 0 });
      expect(parseKey('10,5')).toEqual({ row: 10, col: 5 });
    });

    it('should be reversible', () => {
      const row = 42;
      const col = 17;
      const key = cellKey(row, col);
      const parsed = parseKey(key);
      expect(parsed).toEqual({ row, col });
    });
  });

  describe('getCellValueType', () => {
    it('should identify null values', () => {
      expect(getCellValueType(null)).toBe('null');
      expect(getCellValueType(undefined as unknown as null)).toBe('null');
    });

    it('should identify string values', () => {
      expect(getCellValueType('hello')).toBe('string');
      expect(getCellValueType('')).toBe('string');
    });

    it('should identify number values', () => {
      expect(getCellValueType(42)).toBe('number');
      expect(getCellValueType(3.14)).toBe('number');
      expect(getCellValueType(0)).toBe('number');
    });

    it('should identify boolean values', () => {
      expect(getCellValueType(true)).toBe('boolean');
      expect(getCellValueType(false)).toBe('boolean');
    });

    it('should identify date values', () => {
      expect(getCellValueType(new Date())).toBe('date');
    });

    it('should identify error values', () => {
      expect(getCellValueType('#DIV/0!')).toBe('error');
      expect(getCellValueType('#VALUE!')).toBe('error');
      expect(getCellValueType('#REF!')).toBe('error');
      expect(getCellValueType('#NAME?')).toBe('error');
      expect(getCellValueType('#N/A')).toBe('error');
    });

    it('should identify rich text values', () => {
      expect(getCellValueType({ richText: [{ text: 'Hello' }] })).toBe('string');
    });
  });
});

describe('Range Type Utilities', () => {
  describe('parseRangeReference', () => {
    it('should parse single cell reference', () => {
      expect(parseRangeReference('A1')).toEqual({
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      });
    });

    it('should parse range reference', () => {
      expect(parseRangeReference('A1:C3')).toEqual({
        startRow: 0,
        startCol: 0,
        endRow: 2,
        endCol: 2,
      });
    });

    it('should handle reversed ranges', () => {
      expect(parseRangeReference('C3:A1')).toEqual({
        startRow: 0,
        startCol: 0,
        endRow: 2,
        endCol: 2,
      });
    });

    it('should strip sheet name', () => {
      expect(parseRangeReference('Sheet1!A1:B2')).toEqual({
        startRow: 0,
        startCol: 0,
        endRow: 1,
        endCol: 1,
      });
    });
  });

  describe('rangeToA1', () => {
    it('should convert single cell range', () => {
      expect(rangeToA1({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 })).toBe('A1');
    });

    it('should convert multi-cell range', () => {
      expect(rangeToA1({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 })).toBe('A1:C3');
    });
  });

  describe('rangesOverlap', () => {
    it('should detect overlapping ranges', () => {
      const a = { startRow: 0, startCol: 0, endRow: 2, endCol: 2 };
      const b = { startRow: 1, startCol: 1, endRow: 3, endCol: 3 };
      expect(rangesOverlap(a, b)).toBe(true);
    });

    it('should detect non-overlapping ranges', () => {
      const a = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
      const b = { startRow: 3, startCol: 3, endRow: 4, endCol: 4 };
      expect(rangesOverlap(a, b)).toBe(false);
    });

    it('should detect adjacent ranges as non-overlapping', () => {
      const a = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
      const b = { startRow: 2, startCol: 0, endRow: 3, endCol: 1 };
      expect(rangesOverlap(a, b)).toBe(false);
    });
  });

  describe('isCellInRange', () => {
    const range = { startRow: 1, startCol: 1, endRow: 3, endCol: 3 };

    it('should return true for cells inside range', () => {
      expect(isCellInRange({ row: 2, col: 2 }, range)).toBe(true);
      expect(isCellInRange({ row: 1, col: 1 }, range)).toBe(true);
      expect(isCellInRange({ row: 3, col: 3 }, range)).toBe(true);
    });

    it('should return false for cells outside range', () => {
      expect(isCellInRange({ row: 0, col: 0 }, range)).toBe(false);
      expect(isCellInRange({ row: 4, col: 2 }, range)).toBe(false);
      expect(isCellInRange({ row: 2, col: 0 }, range)).toBe(false);
    });
  });

  describe('getRangeIntersection', () => {
    it('should return intersection of overlapping ranges', () => {
      const a = { startRow: 0, startCol: 0, endRow: 2, endCol: 2 };
      const b = { startRow: 1, startCol: 1, endRow: 3, endCol: 3 };

      expect(getRangeIntersection(a, b)).toEqual({
        startRow: 1,
        startCol: 1,
        endRow: 2,
        endCol: 2,
      });
    });

    it('should return null for non-overlapping ranges', () => {
      const a = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
      const b = { startRow: 3, startCol: 3, endRow: 4, endCol: 4 };

      expect(getRangeIntersection(a, b)).toBeNull();
    });
  });

  describe('getRangeUnion', () => {
    it('should return bounding box of two ranges', () => {
      const a = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
      const b = { startRow: 3, startCol: 0, endRow: 4, endCol: 3 };

      expect(getRangeUnion(a, b)).toEqual({
        startRow: 1,
        startCol: 0,
        endRow: 4,
        endCol: 3,
      });
    });
  });

  describe('iterateRange', () => {
    it('should iterate over all cells in range', () => {
      const range = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
      const cells = [...iterateRange(range)];

      expect(cells).toHaveLength(4);
      expect(cells).toContainEqual({ row: 0, col: 0 });
      expect(cells).toContainEqual({ row: 0, col: 1 });
      expect(cells).toContainEqual({ row: 1, col: 0 });
      expect(cells).toContainEqual({ row: 1, col: 1 });
    });

    it('should iterate in row-major order', () => {
      const range = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
      const cells = [...iterateRange(range)];

      expect(cells[0]).toEqual({ row: 0, col: 0 });
      expect(cells[1]).toEqual({ row: 0, col: 1 });
      expect(cells[2]).toEqual({ row: 1, col: 0 });
      expect(cells[3]).toEqual({ row: 1, col: 1 });
    });
  });

  describe('getRangeDimensions', () => {
    it('should return correct dimensions', () => {
      expect(
        getRangeDimensions({ startRow: 0, startCol: 0, endRow: 2, endCol: 3 })
      ).toEqual({ rows: 3, cols: 4 });

      expect(
        getRangeDimensions({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 })
      ).toEqual({ rows: 1, cols: 1 });
    });
  });
});

describe('Style Type Utilities', () => {
  describe('createSolidFill', () => {
    it('should create solid fill object', () => {
      const fill = createSolidFill('#FF0000');

      expect(fill.type).toBe('pattern');
      expect(fill.pattern).toBe('solid');
      expect(fill.foregroundColor).toBe('#FF0000');
    });
  });

  describe('createBorder', () => {
    it('should create border object with default color', () => {
      const border = createBorder('thin');

      expect(border.style).toBe('thin');
      expect(border.color).toBe('#000000');
    });

    it('should create border object with custom color', () => {
      const border = createBorder('medium', '#FF0000');

      expect(border.style).toBe('medium');
      expect(border.color).toBe('#FF0000');
    });
  });

  describe('createUniformBorders', () => {
    it('should create borders on all sides', () => {
      const borders = createUniformBorders('thin');

      expect(borders.top?.style).toBe('thin');
      expect(borders.right?.style).toBe('thin');
      expect(borders.bottom?.style).toBe('thin');
      expect(borders.left?.style).toBe('thin');
    });

    it('should use same color on all sides', () => {
      const borders = createUniformBorders('medium', '#0000FF');

      expect(borders.top?.color).toBe('#0000FF');
      expect(borders.right?.color).toBe('#0000FF');
      expect(borders.bottom?.color).toBe('#0000FF');
      expect(borders.left?.color).toBe('#0000FF');
    });
  });
});

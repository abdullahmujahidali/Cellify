import { describe, it, expect } from 'vitest';
import { Cell } from '../src/core/Cell.js';

describe('Cell', () => {
  describe('constructor', () => {
    it('should create a cell with row and column', () => {
      const cell = new Cell(0, 0);
      expect(cell.row).toBe(0);
      expect(cell.col).toBe(0);
      expect(cell.value).toBeNull();
    });

    it('should create a cell with initial value', () => {
      const cell = new Cell(0, 0, 'Hello');
      expect(cell.value).toBe('Hello');
    });
  });

  describe('address', () => {
    it('should return A1 notation for cell address', () => {
      expect(new Cell(0, 0).address).toBe('A1');
      expect(new Cell(0, 25).address).toBe('Z1');
      expect(new Cell(0, 26).address).toBe('AA1');
      expect(new Cell(9, 2).address).toBe('C10');
    });
  });

  describe('value', () => {
    it('should set and get string values', () => {
      const cell = new Cell(0, 0);
      cell.value = 'Test';
      expect(cell.value).toBe('Test');
      expect(cell.type).toBe('string');
    });

    it('should set and get number values', () => {
      const cell = new Cell(0, 0);
      cell.value = 42;
      expect(cell.value).toBe(42);
      expect(cell.type).toBe('number');
    });

    it('should set and get boolean values', () => {
      const cell = new Cell(0, 0);
      cell.value = true;
      expect(cell.value).toBe(true);
      expect(cell.type).toBe('boolean');
    });

    it('should set and get date values', () => {
      const cell = new Cell(0, 0);
      const date = new Date('2024-01-01');
      cell.value = date;
      expect(cell.value).toEqual(date);
      expect(cell.type).toBe('date');
    });

    it('should handle null values', () => {
      const cell = new Cell(0, 0, 'test');
      cell.value = null;
      expect(cell.value).toBeNull();
      expect(cell.type).toBe('null');
    });
  });

  describe('formula', () => {
    it('should set a formula', () => {
      const cell = new Cell(0, 0);
      cell.setFormula('=SUM(A1:A10)');
      expect(cell.formula?.formula).toBe('SUM(A1:A10)');
      expect(cell.type).toBe('formula');
    });

    it('should strip leading equals sign', () => {
      const cell = new Cell(0, 0);
      cell.setFormula('=A1+B1');
      expect(cell.formula?.formula).toBe('A1+B1');
    });

    it('should clear formula when value is set directly', () => {
      const cell = new Cell(0, 0);
      cell.setFormula('=A1+B1');
      cell.value = 100;
      expect(cell.formula).toBeUndefined();
      expect(cell.value).toBe(100);
    });

    it('should clear formula with clearFormula()', () => {
      const cell = new Cell(0, 0);
      cell.setFormula('=A1+B1');
      cell.clearFormula();
      expect(cell.formula).toBeUndefined();
    });
  });

  describe('style', () => {
    it('should set and get style', () => {
      const cell = new Cell(0, 0);
      const style = { font: { bold: true } };
      cell.style = style;
      expect(cell.style).toEqual(style);
    });

    it('should apply partial style updates', () => {
      const cell = new Cell(0, 0);
      cell.applyStyle({ font: { bold: true } });
      cell.applyStyle({ font: { italic: true } });
      expect(cell.style?.font?.bold).toBe(true);
      expect(cell.style?.font?.italic).toBe(true);
    });

    it('should merge nested style properties', () => {
      const cell = new Cell(0, 0);
      cell.applyStyle({
        font: { bold: true, size: 14 },
        fill: { type: 'pattern', pattern: 'solid' },
      });
      cell.applyStyle({
        font: { italic: true },
        alignment: { horizontal: 'center' },
      });

      expect(cell.style?.font?.bold).toBe(true);
      expect(cell.style?.font?.size).toBe(14);
      expect(cell.style?.font?.italic).toBe(true);
      expect(cell.style?.fill?.type).toBe('pattern');
      expect(cell.style?.alignment?.horizontal).toBe('center');
    });
  });

  describe('hyperlink', () => {
    it('should set and get hyperlink', () => {
      const cell = new Cell(0, 0);
      cell.setHyperlink('https://example.com', 'Example');
      expect(cell.hyperlink?.target).toBe('https://example.com');
      expect(cell.hyperlink?.tooltip).toBe('Example');
    });

    it('should clear hyperlink', () => {
      const cell = new Cell(0, 0);
      cell.setHyperlink('https://example.com');
      cell.clearHyperlink();
      expect(cell.hyperlink).toBeUndefined();
    });
  });

  describe('comment', () => {
    it('should set and get comment', () => {
      const cell = new Cell(0, 0);
      cell.setComment('This is a note', 'Author');
      expect(cell.comment?.text).toBe('This is a note');
      expect(cell.comment?.author).toBe('Author');
    });

    it('should clear comment', () => {
      const cell = new Cell(0, 0);
      cell.setComment('Note');
      cell.clearComment();
      expect(cell.comment).toBeUndefined();
    });

    it('should set comment via setter with string', () => {
      const cell = new Cell(0, 0);
      cell.comment = 'Simple comment';
      expect(cell.comment?.text).toBe('Simple comment');
    });

    it('should set comment via setter with CellComment object', () => {
      const cell = new Cell(0, 0);
      cell.comment = { text: 'Object comment', author: 'Test Author' };
      expect(cell.comment?.text).toBe('Object comment');
      expect(cell.comment?.author).toBe('Test Author');
    });

    it('should clear comment via setter with undefined', () => {
      const cell = new Cell(0, 0);
      cell.setComment('Note to clear');
      cell.comment = undefined;
      expect(cell.comment).toBeUndefined();
    });

    it('should clear comment via setter with null', () => {
      const cell = new Cell(0, 0);
      cell.setComment('Note to clear');
      cell.comment = null;
      expect(cell.comment).toBeUndefined();
    });
  });

  describe('validation', () => {
    it('should set and get validation', () => {
      const cell = new Cell(0, 0);
      cell.setValidation({
        type: 'whole',
        operator: 'between',
        formula1: 1,
        formula2: 100,
      });
      expect(cell.validation?.type).toBe('whole');
      expect(cell.validation?.operator).toBe('between');
    });

    it('should clear validation', () => {
      const cell = new Cell(0, 0);
      cell.setValidation({ type: 'list', formula1: 'A,B,C' });
      cell.clearValidation();
      expect(cell.validation).toBeUndefined();
    });
  });

  describe('isEmpty', () => {
    it('should return true for empty cell', () => {
      const cell = new Cell(0, 0);
      expect(cell.isEmpty).toBe(true);
    });

    it('should return false for cell with value', () => {
      const cell = new Cell(0, 0, 'test');
      expect(cell.isEmpty).toBe(false);
    });

    it('should return false for cell with style', () => {
      const cell = new Cell(0, 0);
      cell.style = { font: { bold: true } };
      expect(cell.isEmpty).toBe(false);
    });
  });

  describe('clear', () => {
    it('should clear all cell content', () => {
      const cell = new Cell(0, 0, 'test');
      cell.setFormula('=A1');
      cell.style = { font: { bold: true } };
      cell.setHyperlink('https://example.com');
      cell.setComment('Note');
      cell.setValidation({ type: 'whole' });

      cell.clear();

      expect(cell.value).toBeNull();
      expect(cell.formula).toBeUndefined();
      expect(cell.style).toBeUndefined();
      expect(cell.hyperlink).toBeUndefined();
      expect(cell.comment).toBeUndefined();
      expect(cell.validation).toBeUndefined();
    });
  });

  describe('clone', () => {
    it('should create a deep copy of the cell', () => {
      const original = new Cell(0, 0, 'test');
      original.style = { font: { bold: true } };
      original.setHyperlink('https://example.com');

      const clone = original.clone();

      // Should have same values
      expect(clone.value).toBe('test');
      expect(clone.style?.font?.bold).toBe(true);
      expect(clone.hyperlink?.target).toBe('https://example.com');

      // Modifying clone should not affect original
      clone.value = 'modified';
      clone.style!.font!.bold = false;

      expect(original.value).toBe('test');
      expect(original.style?.font?.bold).toBe(true);
    });
  });

  describe('toJSON', () => {
    it('should serialize cell to JSON', () => {
      const cell = new Cell(1, 2, 'Hello');
      cell.style = { font: { bold: true } };

      const json = cell.toJSON();

      expect(json.row).toBe(1);
      expect(json.col).toBe(2);
      expect(json.address).toBe('C2');
      expect(json.value).toBe('Hello');
      expect(json.type).toBe('string');
      expect(json.style).toEqual({ font: { bold: true } });
    });
  });
});

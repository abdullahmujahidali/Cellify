import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Workbook', () => {
  describe('constructor', () => {
    it('should create an empty workbook', () => {
      const workbook = new Workbook();
      expect(workbook.sheetCount).toBe(0);
      expect(workbook.properties.created).toBeInstanceOf(Date);
    });
  });

  describe('sheet management', () => {
    it('should add sheets with auto-generated names', () => {
      const workbook = new Workbook();
      const sheet1 = workbook.addSheet();
      const sheet2 = workbook.addSheet();

      expect(sheet1.name).toBe('Sheet1');
      expect(sheet2.name).toBe('Sheet2');
      expect(workbook.sheetCount).toBe(2);
    });

    it('should add sheets with custom names', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('CustomName');

      expect(sheet.name).toBe('CustomName');
    });

    it('should throw on duplicate sheet names', () => {
      const workbook = new Workbook();
      workbook.addSheet('Test');

      expect(() => workbook.addSheet('Test')).toThrow('already exists');
    });

    it('should throw on invalid sheet names', () => {
      const workbook = new Workbook();

      expect(() => workbook.addSheet('')).toThrow('empty');
      expect(() => workbook.addSheet('a'.repeat(32))).toThrow('31 characters');
      expect(() => workbook.addSheet('Test[1]')).toThrow('invalid characters');
      expect(() => workbook.addSheet('Test/Sheet')).toThrow('invalid characters');
    });

    it('should get sheet by name', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(workbook.getSheet('Test')).toBe(sheet);
      expect(workbook.getSheet('NonExistent')).toBeUndefined();
    });

    it('should get sheet by index', () => {
      const workbook = new Workbook();
      const sheet1 = workbook.addSheet('First');
      const sheet2 = workbook.addSheet('Second');

      expect(workbook.getSheetByIndex(0)).toBe(sheet1);
      expect(workbook.getSheetByIndex(1)).toBe(sheet2);
      expect(workbook.getSheetByIndex(2)).toBeUndefined();
    });

    it('should remove sheet by name', () => {
      const workbook = new Workbook();
      workbook.addSheet('Test');

      expect(workbook.removeSheet('Test')).toBe(true);
      expect(workbook.sheetCount).toBe(0);
      expect(workbook.removeSheet('NonExistent')).toBe(false);
    });

    it('should remove sheet by reference', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(workbook.removeSheet(sheet)).toBe(true);
      expect(workbook.sheetCount).toBe(0);
    });

    it('should rename sheet', () => {
      const workbook = new Workbook();
      workbook.addSheet('OldName');

      expect(workbook.renameSheet('OldName', 'NewName')).toBe(true);
      expect(workbook.getSheet('NewName')).toBeDefined();
      expect(workbook.getSheet('OldName')).toBeUndefined();
    });

    it('should move sheet to new position', () => {
      const workbook = new Workbook();
      workbook.addSheet('First');
      workbook.addSheet('Second');
      workbook.addSheet('Third');

      workbook.moveSheet('Third', 0);

      expect(workbook.getSheetByIndex(0)?.name).toBe('Third');
      expect(workbook.getSheetByIndex(1)?.name).toBe('First');
      expect(workbook.getSheetByIndex(2)?.name).toBe('Second');
    });
  });

  describe('properties', () => {
    it('should set and get title', () => {
      const workbook = new Workbook();
      workbook.title = 'My Workbook';
      expect(workbook.title).toBe('My Workbook');
    });

    it('should set and get author', () => {
      const workbook = new Workbook();
      workbook.author = 'John Doe';
      expect(workbook.author).toBe('John Doe');
    });

    it('should set multiple properties', () => {
      const workbook = new Workbook();
      workbook.setProperties({
        title: 'Test',
        author: 'Author',
        company: 'Company',
      });

      expect(workbook.properties.title).toBe('Test');
      expect(workbook.properties.author).toBe('Author');
      expect(workbook.properties.company).toBe('Company');
    });
  });

  describe('named styles', () => {
    it('should add and get named styles', () => {
      const workbook = new Workbook();
      workbook.addNamedStyle('Header', {
        font: { bold: true, size: 14 },
        fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#CCCCCC' },
      });

      const style = workbook.getNamedStyle('Header');
      expect(style?.name).toBe('Header');
      expect(style?.style.font?.bold).toBe(true);
    });

    it('should remove named styles', () => {
      const workbook = new Workbook();
      workbook.addNamedStyle('Test', { font: { bold: true } });

      expect(workbook.removeNamedStyle('Test')).toBe(true);
      expect(workbook.getNamedStyle('Test')).toBeUndefined();
    });
  });

  describe('defined names', () => {
    it('should add and get defined names', () => {
      const workbook = new Workbook();
      workbook.addSheet('Data');
      workbook.addDefinedName('MyRange', 'Data!$A$1:$B$10');

      const name = workbook.getDefinedName('MyRange');
      expect(name?.formula).toBe('Data!$A$1:$B$10');
    });

    it('should add defined name with options', () => {
      const workbook = new Workbook();
      workbook.addDefinedName('Hidden', 'Sheet1!$A$1', {
        hidden: true,
        comment: 'Internal use',
      });

      const name = workbook.getDefinedName('Hidden');
      expect(name?.hidden).toBe(true);
      expect(name?.comment).toBe('Internal use');
    });
  });

  describe('calculation mode', () => {
    it('should default to auto calculation', () => {
      const workbook = new Workbook();
      expect(workbook.calculationMode).toBe('auto');
    });

    it('should set calculation mode', () => {
      const workbook = new Workbook();
      workbook.calculationMode = 'manual';
      expect(workbook.calculationMode).toBe('manual');
    });
  });

  describe('active sheet', () => {
    it('should get active sheet', () => {
      const workbook = new Workbook();
      const sheet1 = workbook.addSheet('First');
      workbook.addSheet('Second');

      expect(workbook.activeSheet).toBe(sheet1);
    });

    it('should set active sheet by reference', () => {
      const workbook = new Workbook();
      workbook.addSheet('First');
      const sheet2 = workbook.addSheet('Second');

      workbook.setActiveSheet(sheet2);
      expect(workbook.activeSheet).toBe(sheet2);
    });

    it('should set active sheet by name', () => {
      const workbook = new Workbook();
      workbook.addSheet('First');
      const sheet2 = workbook.addSheet('Second');

      workbook.setActiveSheet('Second');
      expect(workbook.activeSheet).toBe(sheet2);
    });

    it('should set active sheet by index', () => {
      const workbook = new Workbook();
      workbook.addSheet('First');
      const sheet2 = workbook.addSheet('Second');

      workbook.setActiveSheet(1);
      expect(workbook.activeSheet).toBe(sheet2);
    });
  });

  describe('toJSON', () => {
    it('should serialize workbook to JSON', () => {
      const workbook = new Workbook();
      workbook.title = 'Test Workbook';
      workbook.addSheet('Sheet1');

      const json = workbook.toJSON();

      expect(json.properties).toBeDefined();
      expect((json.properties as Record<string, unknown>).title).toBe('Test Workbook');
      expect(json.sheets).toHaveLength(1);
    });
  });
});

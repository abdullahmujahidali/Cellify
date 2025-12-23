import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Sheet Search', () => {
  describe('find', () => {
    it('should find cell by exact string match', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.cell('A2').value = 'World';
      sheet.cell('A3').value = 'Hello World';

      const cell = sheet.find('Hello');
      expect(cell).toBeDefined();
      expect(cell?.value).toBe('Hello');
    });

    it('should find cell by partial match', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello World';

      const cell = sheet.find('World');
      expect(cell).toBeDefined();
      expect(cell?.value).toBe('Hello World');
    });

    it('should find cell by number', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 100;
      sheet.cell('A2').value = 200;

      const cell = sheet.find(100);
      expect(cell).toBeDefined();
      expect(cell?.value).toBe(100);
    });

    it('should find cell by regex', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'test123';
      sheet.cell('A2').value = 'hello';

      const cell = sheet.find(/\d+/);
      expect(cell).toBeDefined();
      expect(cell?.value).toBe('test123');
    });

    it('should return undefined when not found', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';

      const cell = sheet.find('NotFound');
      expect(cell).toBeUndefined();
    });

    it('should be case insensitive by default', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'HELLO';

      const cell = sheet.find('hello');
      expect(cell).toBeDefined();
    });

    it('should respect matchCase option', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'HELLO';
      sheet.cell('A2').value = 'hello';

      const cell = sheet.find('hello', { matchCase: true });
      expect(cell?.value).toBe('hello');
    });

    it('should respect matchCell option', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello World';
      sheet.cell('A2').value = 'Hello';

      const cell = sheet.find('Hello', { matchCell: true });
      expect(cell?.value).toBe('Hello');
    });
  });

  describe('findAll', () => {
    it('should find all matching cells', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.cell('A2').value = 'World';
      sheet.cell('A3').value = 'Hello Again';

      const cells = sheet.findAll('Hello');
      expect(cells.length).toBe(2);
    });

    it('should return empty array when no matches', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';

      const cells = sheet.findAll('NotFound');
      expect(cells.length).toBe(0);
    });

    it('should search in formulas when specified', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 100;
      sheet.cell('A2').setFormula('=SUM(A1:A1)', 100);

      const cells = sheet.findAll('SUM', { searchIn: 'formulas' });
      expect(cells.length).toBe(1);
      expect(cells[0].address).toBe('A2');
    });

    it('should search in both values and formulas', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'SUM';
      sheet.cell('A2').setFormula('=SUM(A1:A1)', 0);

      const cells = sheet.findAll('SUM', { searchIn: 'both' });
      expect(cells.length).toBe(2);
    });

    it('should respect range option', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'test';
      sheet.cell('A2').value = 'test';
      sheet.cell('A3').value = 'test';

      const cells = sheet.findAll('test', { range: 'A1:A2' });
      expect(cells.length).toBe(2);
    });
  });

  describe('replace', () => {
    it('should replace first occurrence', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello World';
      sheet.cell('A2').value = 'Hello Again';

      const cell = sheet.replace('Hello', 'Hi');
      expect(cell?.value).toBe('Hi World');
      expect(sheet.cell('A2').value).toBe('Hello Again'); // Unchanged
    });

    it('should replace with regex', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'test123';

      sheet.replace(/\d+/, '456');
      expect(sheet.cell('A1').value).toBe('test456');
    });

    it('should return undefined if not found', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';

      const cell = sheet.replace('NotFound', 'X');
      expect(cell).toBeUndefined();
    });
  });

  describe('replaceAll', () => {
    it('should replace all occurrences', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello World';
      sheet.cell('A2').value = 'Hello Again';

      const cells = sheet.replaceAll('Hello', 'Hi');
      expect(cells.length).toBe(2);
      expect(sheet.cell('A1').value).toBe('Hi World');
      expect(sheet.cell('A2').value).toBe('Hi Again');
    });

    it('should replace numbers', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 100;
      sheet.cell('A2').value = 100;
      sheet.cell('A3').value = 200;

      const cells = sheet.replaceAll(100, 150);
      expect(cells.length).toBe(2);
      expect(sheet.cell('A1').value).toBe(150);
      expect(sheet.cell('A2').value).toBe(150);
      expect(sheet.cell('A3').value).toBe(200);
    });
  });
});

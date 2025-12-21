import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';

describe('Sheet Undo/Redo', () => {
  describe('basic undo', () => {
    it('should undo a single value change', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      expect(sheet.cell('A1').value).toBe('Hello');

      sheet.undo();
      expect(sheet.cell('A1').value).toBeNull();
    });

    it('should undo multiple value changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A1').value = 'Second';
      sheet.cell('A1').value = 'Third';

      expect(sheet.cell('A1').value).toBe('Third');

      sheet.undo();
      expect(sheet.cell('A1').value).toBe('Second');

      sheet.undo();
      expect(sheet.cell('A1').value).toBe('First');

      sheet.undo();
      expect(sheet.cell('A1').value).toBeNull();
    });

    it('should undo style changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';
      sheet.cell('A1').style = { font: { bold: true } };

      expect(sheet.cell('A1').style?.font?.bold).toBe(true);

      sheet.undo();
      expect(sheet.cell('A1').style).toBeUndefined();
    });

    it('should return false when nothing to undo', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(sheet.undo()).toBe(false);
    });
  });

  describe('basic redo', () => {
    it('should redo an undone change', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.undo();
      expect(sheet.cell('A1').value).toBeNull();

      sheet.redo();
      expect(sheet.cell('A1').value).toBe('Hello');
    });

    it('should redo multiple undone changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A1').value = 'Second';

      sheet.undo();
      sheet.undo();

      expect(sheet.cell('A1').value).toBeNull();

      sheet.redo();
      expect(sheet.cell('A1').value).toBe('First');

      sheet.redo();
      expect(sheet.cell('A1').value).toBe('Second');
    });

    it('should return false when nothing to redo', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(sheet.redo()).toBe(false);
    });

    it('should clear redo stack on new change', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A1').value = 'Second';

      sheet.undo(); // Back to First
      expect(sheet.canRedo).toBe(true);

      sheet.cell('A1').value = 'New'; // New change clears redo
      expect(sheet.canRedo).toBe(false);
    });
  });

  describe('canUndo and canRedo', () => {
    it('should report canUndo correctly', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(sheet.canUndo).toBe(false);

      sheet.cell('A1').value = 'Hello';
      expect(sheet.canUndo).toBe(true);

      sheet.undo();
      expect(sheet.canUndo).toBe(false);
    });

    it('should report canRedo correctly', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(sheet.canRedo).toBe(false);

      sheet.cell('A1').value = 'Hello';
      expect(sheet.canRedo).toBe(false);

      sheet.undo();
      expect(sheet.canRedo).toBe(true);

      sheet.redo();
      expect(sheet.canRedo).toBe(false);
    });
  });

  describe('undoCount and redoCount', () => {
    it('should track undo count', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(sheet.undoCount).toBe(0);

      sheet.cell('A1').value = 'First';
      expect(sheet.undoCount).toBe(1);

      sheet.cell('A1').value = 'Second';
      expect(sheet.undoCount).toBe(2);

      sheet.undo();
      expect(sheet.undoCount).toBe(1);
    });

    it('should track redo count', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A1').value = 'Second';

      expect(sheet.redoCount).toBe(0);

      sheet.undo();
      expect(sheet.redoCount).toBe(1);

      sheet.undo();
      expect(sheet.redoCount).toBe(2);

      sheet.redo();
      expect(sheet.redoCount).toBe(1);
    });
  });

  describe('clearHistory', () => {
    it('should clear undo and redo history', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'First';
      sheet.cell('A1').value = 'Second';
      sheet.undo();

      expect(sheet.canUndo).toBe(true);
      expect(sheet.canRedo).toBe(true);

      sheet.clearHistory();

      expect(sheet.canUndo).toBe(false);
      expect(sheet.canRedo).toBe(false);
    });
  });

  describe('setMaxUndoHistory', () => {
    it('should limit undo history', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.setMaxUndoHistory(3);

      sheet.cell('A1').value = 'One';
      sheet.cell('A1').value = 'Two';
      sheet.cell('A1').value = 'Three';
      sheet.cell('A1').value = 'Four';
      sheet.cell('A1').value = 'Five';

      expect(sheet.undoCount).toBe(3); // Only last 3 kept

      sheet.undo();
      expect(sheet.cell('A1').value).toBe('Four');

      sheet.undo();
      expect(sheet.cell('A1').value).toBe('Three');

      sheet.undo();
      expect(sheet.cell('A1').value).toBe('Two');

      // Can't undo further
      expect(sheet.canUndo).toBe(false);
    });
  });

  describe('batch operations', () => {
    it('should undo batch as single operation', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.batch(() => {
        sheet.cell('A1').value = 'Hello';
        sheet.cell('B1').value = 'World';
        sheet.cell('C1').value = '!';
      });

      expect(sheet.cell('A1').value).toBe('Hello');
      expect(sheet.cell('B1').value).toBe('World');
      expect(sheet.cell('C1').value).toBe('!');

      // Single undo should revert all three
      sheet.undo();

      expect(sheet.cell('A1').value).toBeNull();
      expect(sheet.cell('B1').value).toBeNull();
      expect(sheet.cell('C1').value).toBeNull();
    });

    it('should redo batch as single operation', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.batch(() => {
        sheet.cell('A1').value = 'Hello';
        sheet.cell('B1').value = 'World';
      });

      sheet.undo();
      sheet.redo();

      expect(sheet.cell('A1').value).toBe('Hello');
      expect(sheet.cell('B1').value).toBe('World');
    });

    it('should count batch as single undo step', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.batch(() => {
        sheet.cell('A1').value = 'One';
        sheet.cell('A2').value = 'Two';
        sheet.cell('A3').value = 'Three';
      });

      expect(sheet.undoCount).toBe(1);
    });
  });

  describe('multiple cells', () => {
    it('should undo changes to different cells independently', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'B';
      sheet.cell('C1').value = 'C';

      sheet.undo(); // Undo C
      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBe('B');
      expect(sheet.cell('C1').value).toBeNull();

      sheet.undo(); // Undo B
      expect(sheet.cell('A1').value).toBe('A');
      expect(sheet.cell('B1').value).toBeNull();

      sheet.undo(); // Undo A
      expect(sheet.cell('A1').value).toBeNull();
    });
  });

  describe('undo/redo with events disabled', () => {
    it('should not track changes in undo history when events are disabled', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.setEventsEnabled(false);
      sheet.cell('A1').value = 'Hello';
      sheet.setEventsEnabled(true);

      // No undo available because events were disabled
      expect(sheet.canUndo).toBe(false);
    });
  });
});

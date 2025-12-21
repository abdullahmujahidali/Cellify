import { describe, it, expect, vi } from 'vitest';
import { Workbook } from '../src/core/Workbook.js';
import type { CellChangeEvent, CellStyleChangeEvent, CellAddedEvent, CellDeletedEvent } from '../src/types/event.types.js';

describe('Sheet Event System', () => {
  describe('cellChange events', () => {
    it('should emit cellChange event when cell value changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      // First set a value so the cell exists
      sheet.cell('A1').value = 'Hello';

      // Subscribe after initial value is set
      sheet.on('cellChange', handler);
      sheet.cell('A1').value = 'World';

      expect(handler).toHaveBeenCalledTimes(1);
      const event = handler.mock.calls[0][0] as CellChangeEvent;
      expect(event.type).toBe('cellChange');
      expect(event.address).toBe('A1');
      expect(event.oldValue).toBe('Hello');
      expect(event.newValue).toBe('World');
      expect(event.sheetName).toBe('Test');
    });

    it('should not emit event when value is unchanged', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.cell('A1').value = 'Hello';
      sheet.on('cellChange', handler);
      sheet.cell('A1').value = 'Hello'; // Same value

      expect(handler).not.toHaveBeenCalled();
    });

    it('should emit cellChange event for formula changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.on('cellChange', handler);
      sheet.cell('A1').setFormula('=SUM(B1:B10)');

      expect(handler).toHaveBeenCalledTimes(1);
      const event = handler.mock.calls[0][0] as CellChangeEvent;
      expect(event.type).toBe('cellChange');
    });
  });

  describe('cellStyleChange events', () => {
    it('should emit cellStyleChange event when style is set', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.cell('A1').value = 'Test';
      sheet.on('cellStyleChange', handler);
      sheet.cell('A1').style = { font: { bold: true } };

      expect(handler).toHaveBeenCalledTimes(1);
      const event = handler.mock.calls[0][0] as CellStyleChangeEvent;
      expect(event.type).toBe('cellStyleChange');
      expect(event.address).toBe('A1');
      expect(event.oldStyle).toBeUndefined();
      expect(event.newStyle?.font?.bold).toBe(true);
    });

    it('should emit cellStyleChange event when applyStyle is called', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.cell('A1').value = 'Test';
      sheet.on('cellStyleChange', handler);
      sheet.cell('A1').applyStyle({ font: { italic: true } });

      expect(handler).toHaveBeenCalledTimes(1);
    });
  });

  describe('cellAdded events', () => {
    it('should emit cellAdded event when new cell is created', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.on('cellAdded', handler);
      sheet.cell('A1').value = 'Hello';
      sheet.cell('B2').value = 'World';

      expect(handler).toHaveBeenCalledTimes(2);

      const event1 = handler.mock.calls[0][0] as CellAddedEvent;
      expect(event1.type).toBe('cellAdded');
      expect(event1.address).toBe('A1');

      const event2 = handler.mock.calls[1][0] as CellAddedEvent;
      expect(event2.address).toBe('B2');
    });

    it('should not emit cellAdded for existing cells', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';

      const handler = vi.fn();
      sheet.on('cellAdded', handler);

      sheet.cell('A1').value = 'Updated';

      expect(handler).not.toHaveBeenCalled();
    });
  });

  describe('cellDeleted events', () => {
    it('should emit cellDeleted event when cell is deleted', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';

      const handler = vi.fn();
      sheet.on('cellDeleted', handler);

      sheet.deleteCell('A1');

      expect(handler).toHaveBeenCalledTimes(1);
      const event = handler.mock.calls[0][0] as CellDeletedEvent;
      expect(event.type).toBe('cellDeleted');
      expect(event.address).toBe('A1');
      expect(event.value).toBe('Hello');
    });
  });

  describe('wildcard listener', () => {
    it('should receive all events with * listener', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.on('*', handler);

      sheet.cell('A1').value = 'Hello'; // cellAdded + cellChange (null->Hello)
      sheet.cell('A1').value = 'World'; // cellChange (Hello->World)
      sheet.cell('A1').style = { font: { bold: true } }; // cellStyleChange

      expect(handler).toHaveBeenCalledTimes(4);
      expect(handler.mock.calls[0][0].type).toBe('cellAdded');
      expect(handler.mock.calls[1][0].type).toBe('cellChange'); // null -> Hello
      expect(handler.mock.calls[2][0].type).toBe('cellChange'); // Hello -> World
      expect(handler.mock.calls[3][0].type).toBe('cellStyleChange');
    });
  });

  describe('event unsubscription', () => {
    it('should stop receiving events after off()', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      // Set initial value before subscribing
      sheet.cell('A1').value = 'Hello';

      sheet.on('cellChange', handler);
      sheet.cell('A1').value = 'World'; // 1 event

      sheet.off('cellChange', handler);
      sheet.cell('A1').value = 'Again'; // No event (unsubscribed)

      expect(handler).toHaveBeenCalledTimes(1);
    });
  });

  describe('events enabled/disabled', () => {
    it('should not emit events when disabled', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.on('cellChange', handler);
      sheet.setEventsEnabled(false);

      sheet.cell('A1').value = 'Hello';
      sheet.cell('A1').value = 'World';

      expect(handler).not.toHaveBeenCalled();
    });

    it('should resume emitting events when re-enabled', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler = vi.fn();

      sheet.on('cellChange', handler);
      sheet.setEventsEnabled(false);
      sheet.cell('A1').value = 'Hello';

      sheet.setEventsEnabled(true);
      sheet.cell('A1').value = 'World';

      expect(handler).toHaveBeenCalledTimes(1);
    });

    it('should report eventsEnabled state', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      expect(sheet.eventsEnabled).toBe(true);
      sheet.setEventsEnabled(false);
      expect(sheet.eventsEnabled).toBe(false);
    });
  });

  describe('change tracking', () => {
    it('should track value changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello'; // Change from null to Hello
      sheet.cell('A1').value = 'World'; // Change from Hello to World

      const changes = sheet.getChanges();
      expect(changes.length).toBe(2); // Both changes tracked
      expect(changes[0].type).toBe('value');
      expect(changes[0].address).toBe('A1');
      expect(changes[0].oldValue).toBeNull();
      expect(changes[0].newValue).toBe('Hello');
      expect(changes[1].oldValue).toBe('Hello');
      expect(changes[1].newValue).toBe('World');
    });

    it('should track style changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Test';
      sheet.cell('A1').style = { font: { bold: true } };

      const changes = sheet.getChanges();
      const styleChange = changes.find((c) => c.type === 'style');
      expect(styleChange).toBeDefined();
      expect(styleChange?.newStyle?.font?.bold).toBe(true);
    });

    it('should track delete changes', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.deleteCell('A1');

      const changes = sheet.getChanges();
      const deleteChange = changes.find((c) => c.type === 'delete');
      expect(deleteChange).toBeDefined();
      expect(deleteChange?.oldValue).toBe('Hello');
    });

    it('should clear changes on commit', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.cell('A1').value = 'World';

      expect(sheet.changeCount).toBe(2); // Two value changes
      sheet.commitChanges();
      expect(sheet.changeCount).toBe(0);
      expect(sheet.getChanges().length).toBe(0);
    });

    it('should have unique change IDs', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');

      sheet.cell('A1').value = 'Hello';
      sheet.cell('A1').value = 'World';
      sheet.cell('B1').value = 'Test';
      sheet.cell('B1').value = 'Again';

      const changes = sheet.getChanges();
      const ids = changes.map((c) => c.id);
      const uniqueIds = new Set(ids);
      expect(uniqueIds.size).toBe(ids.length);
    });
  });

  describe('multiple listeners', () => {
    it('should notify all listeners for same event', () => {
      const workbook = new Workbook();
      const sheet = workbook.addSheet('Test');
      const handler1 = vi.fn();
      const handler2 = vi.fn();

      sheet.cell('A1').value = 'Hello';

      sheet.on('cellChange', handler1);
      sheet.on('cellChange', handler2);

      sheet.cell('A1').value = 'World';

      expect(handler1).toHaveBeenCalledTimes(1);
      expect(handler2).toHaveBeenCalledTimes(1);
    });
  });
});

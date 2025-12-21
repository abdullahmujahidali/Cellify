---
sidebar_position: 10
---

# Events, Change Tracking & Undo/Redo

Cellify provides an event system for tracking changes to sheets. This enables real-time sync, undo/redo, and integration with collaboration libraries like Yjs or Liveblocks.

## Subscribing to Events

Use `sheet.on()` to listen for changes:

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

// Listen for cell value changes
sheet.on('cellChange', (event) => {
  console.log(`Cell ${event.address} changed`);
  console.log(`  Old value: ${event.oldValue}`);
  console.log(`  New value: ${event.newValue}`);
});

// Listen for style changes
sheet.on('cellStyleChange', (event) => {
  console.log(`Cell ${event.address} style changed`);
});

// Listen for new cells
sheet.on('cellAdded', (event) => {
  console.log(`New cell created at ${event.address}`);
});

// Listen for deleted cells
sheet.on('cellDeleted', (event) => {
  console.log(`Cell ${event.address} deleted, had value: ${event.value}`);
});
```

## Wildcard Listener

Use `'*'` to listen to all events:

```typescript
sheet.on('*', (event) => {
  console.log(`Event: ${event.type} on ${event.address}`);
});
```

## Unsubscribing

Use `sheet.off()` to remove a listener:

```typescript
const handler = (event) => console.log(event);

sheet.on('cellChange', handler);
// ... later
sheet.off('cellChange', handler);
```

## Change Tracking

Cellify tracks all changes for sync purposes:

```typescript
// Make changes
sheet.cell('A1').value = 'Hello';
sheet.cell('B1').value = 'World';
sheet.cell('A1').style = { font: { bold: true } };

// Get all changes since last commit
const changes = sheet.getChanges();
console.log(`${changes.length} changes pending`);

// Each change has:
// - id: Unique identifier
// - type: 'value' | 'style' | 'formula' | 'delete'
// - address: Cell address (A1 notation)
// - row, col: Cell coordinates
// - oldValue, newValue: For value changes
// - oldStyle, newStyle: For style changes
// - timestamp: When the change occurred

// Sync to server
await fetch('/api/sync', {
  method: 'POST',
  body: JSON.stringify({ changes }),
});

// Clear change buffer after successful sync
sheet.commitChanges();
```

## Disabling Events

For bulk operations, disable events to improve performance:

```typescript
sheet.setEventsEnabled(false);

// Bulk import - no events fired
for (let i = 0; i < 10000; i++) {
  sheet.cell(i, 0).value = i;
}

sheet.setEventsEnabled(true);
```

Check event state with `sheet.eventsEnabled`.

## Undo/Redo

Cellify provides built-in undo/redo functionality:

```typescript
import { Workbook } from 'cellify';

const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

sheet.cell('A1').value = 'Hello';
sheet.cell('A1').value = 'World';

// Undo last change
sheet.undo(); // A1 is now 'Hello'
sheet.undo(); // A1 is now null

// Redo undone changes
sheet.redo(); // A1 is now 'Hello'
sheet.redo(); // A1 is now 'World'
```

### Checking Undo/Redo State

```typescript
if (sheet.canUndo) {
  sheet.undo();
}

if (sheet.canRedo) {
  sheet.redo();
}

console.log(`${sheet.undoCount} undo steps available`);
console.log(`${sheet.redoCount} redo steps available`);
```

### Batch Operations

Group multiple changes into a single undo step:

```typescript
sheet.batch(() => {
  sheet.cell('A1').value = 'Hello';
  sheet.cell('B1').value = 'World';
  sheet.cell('C1').value = '!';
});

// Single undo reverts all three changes
sheet.undo();
```

### Managing History

```typescript
// Clear all undo/redo history
sheet.clearHistory();

// Limit history size (default: 100)
sheet.setMaxUndoHistory(50);
```

## Event Types

### CellChangeEvent

```typescript
interface CellChangeEvent {
  type: 'cellChange';
  sheetName: string;
  address: string; // A1 notation
  row: number;
  col: number;
  oldValue: CellValue;
  newValue: CellValue;
  timestamp: number;
}
```

### CellStyleChangeEvent

```typescript
interface CellStyleChangeEvent {
  type: 'cellStyleChange';
  sheetName: string;
  address: string;
  row: number;
  col: number;
  oldStyle: CellStyle | undefined;
  newStyle: CellStyle | undefined;
  timestamp: number;
}
```

### CellAddedEvent

```typescript
interface CellAddedEvent {
  type: 'cellAdded';
  sheetName: string;
  address: string;
  row: number;
  col: number;
  timestamp: number;
}
```

### CellDeletedEvent

```typescript
interface CellDeletedEvent {
  type: 'cellDeleted';
  sheetName: string;
  address: string;
  row: number;
  col: number;
  value: CellValue;
  timestamp: number;
}
```

## Integration with Yjs

Example of syncing with [Yjs](https://yjs.dev/) for real-time collaboration:

```typescript
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { Workbook } from 'cellify';

const ydoc = new Y.Doc();
const ymap = ydoc.getMap('spreadsheet');
const provider = new WebsocketProvider('ws://localhost:1234', 'room', ydoc);

const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

// Send local changes to Yjs
sheet.on('cellChange', (event) => {
  ymap.set(event.address, {
    value: event.newValue,
    timestamp: event.timestamp,
  });
});

// Apply remote changes from Yjs
ymap.observe((event) => {
  sheet.setEventsEnabled(false); // Prevent echo
  event.changes.keys.forEach((change, key) => {
    if (change.action === 'add' || change.action === 'update') {
      const data = ymap.get(key);
      sheet.cell(key).value = data.value;
    }
  });
  sheet.setEventsEnabled(true);
});
```

## Integration with Liveblocks

Example with [Liveblocks](https://liveblocks.io/):

```typescript
import { createClient } from '@liveblocks/client';
import { Workbook } from 'cellify';

const client = createClient({ publicApiKey: 'pk_...' });
const { room } = client.enterRoom('spreadsheet-room');
const storage = await room.getStorage();

const workbook = new Workbook();
const sheet = workbook.addSheet('Data');

// Send local changes
sheet.on('cellChange', (event) => {
  storage.root.set(event.address, event.newValue);
});

// Receive remote changes
room.subscribe(storage.root, () => {
  sheet.setEventsEnabled(false);
  // Apply updates...
  sheet.setEventsEnabled(true);
});
```

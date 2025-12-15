---
sidebar_position: 7
---

# Building Interactive UIs

This guide shows how to build interactive spreadsheet UIs using Cellify. The demo application (`demo/index.html`) serves as a complete reference implementation.

## Overview

Cellify provides the data layer for spreadsheet operations. Building interactive features like cell editing, context menus, copy/paste, and filtering requires combining Cellify's APIs with standard DOM manipulation.

## Cell Selection and Navigation

### Selecting Cells

```typescript
// Track selected cell
let selectedRow: number | null = null;
let selectedCol: number | null = null;

function selectCell(row: number, col: number) {
  // Clear previous selection
  document.querySelectorAll('td.selected').forEach(td =>
    td.classList.remove('selected')
  );

  selectedRow = row;
  selectedCol = col;

  // Highlight new selection
  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (td) {
    td.classList.add('selected');
  }
}

// Add click handler
table.addEventListener('click', (e) => {
  const td = e.target.closest('td[data-row][data-col]');
  if (td) {
    selectCell(
      parseInt(td.dataset.row),
      parseInt(td.dataset.col)
    );
  }
});
```

### Keyboard Navigation

```typescript
document.addEventListener('keydown', (e) => {
  if (selectedRow === null || selectedCol === null) return;

  switch (e.key) {
    case 'ArrowUp':
      e.preventDefault();
      selectCell(selectedRow - 1, selectedCol);
      break;
    case 'ArrowDown':
      e.preventDefault();
      selectCell(selectedRow + 1, selectedCol);
      break;
    case 'ArrowLeft':
      e.preventDefault();
      selectCell(selectedRow, selectedCol - 1);
      break;
    case 'ArrowRight':
      e.preventDefault();
      selectCell(selectedRow, selectedCol + 1);
      break;
    case 'Tab':
      e.preventDefault();
      selectCell(selectedRow, selectedCol + (e.shiftKey ? -1 : 1));
      break;
  }
});
```

## Cell Editing

### Starting Edit Mode

```typescript
let isEditing = false;

function startEditing(row: number, col: number) {
  if (isEditing) return;

  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (!td) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.getCell(row, col);

  // Get current value or formula
  let editValue = '';
  if (cell?.formula) {
    editValue = '=' + cell.formula.formula;
  } else if (cell?.value !== null && cell?.value !== undefined) {
    editValue = String(cell.value);
  }

  isEditing = true;
  td.classList.add('editing');

  // Create input element
  const input = document.createElement('input');
  input.type = 'text';
  input.value = editValue;
  td.innerHTML = '';
  td.appendChild(input);
  input.focus();
  input.select();

  // Handle save/cancel
  input.addEventListener('blur', () => saveEdit(row, col, input.value));
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      saveEdit(row, col, input.value);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      cancelEdit(row, col);
    }
  });
}

// Double-click to edit
table.addEventListener('dblclick', (e) => {
  const td = e.target.closest('td[data-row][data-col]');
  if (td) {
    startEditing(parseInt(td.dataset.row), parseInt(td.dataset.col));
  }
});

// Enter key to edit
document.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && selectedRow !== null && !isEditing) {
    e.preventDefault();
    startEditing(selectedRow, selectedCol);
  }
});
```

### Saving Edits

```typescript
function saveEdit(row: number, col: number, newValue: string) {
  if (!isEditing) return;

  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (!td) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(row, col);

  // Parse input
  const trimmed = newValue.trim();

  if (trimmed.startsWith('=')) {
    // Formula
    cell.setFormula(trimmed.slice(1));
  } else if (trimmed === '') {
    // Empty
    cell.clear();
  } else if (!isNaN(Number(trimmed))) {
    // Number
    cell.value = Number(trimmed);
  } else {
    // String
    cell.value = trimmed;
  }

  // Update display
  isEditing = false;
  td.classList.remove('editing');
  td.textContent = formatCellValue(cell.value);
}
```

## Undo Functionality

### Implementing Undo Stack

```typescript
interface UndoEntry {
  row: number;
  col: number;
  oldValue: any;
  newValue: any;
}

const undoStack: UndoEntry[] = [];
const MAX_UNDO = 50;

function pushUndo(row: number, col: number, oldValue: any, newValue: any) {
  undoStack.push({ row, col, oldValue, newValue });
  if (undoStack.length > MAX_UNDO) {
    undoStack.shift();
  }
}

function undo() {
  if (undoStack.length === 0) return;

  const entry = undoStack.pop()!;
  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(entry.row, entry.col);

  if (entry.oldValue !== undefined && entry.oldValue !== null) {
    cell.value = entry.oldValue;
  } else {
    cell.clear();
  }

  updateCellDisplay(entry.row, entry.col);
}

// Keyboard shortcut
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
    e.preventDefault();
    undo();
  }
});
```

## Copy, Cut, and Paste

### Clipboard Implementation

```typescript
interface ClipboardData {
  row: number;
  col: number;
  value: any;
  formula?: string;
  style?: CellStyle;
  comment?: CellComment;
  isCut: boolean;
}

let clipboard: ClipboardData | null = null;

function copyCell() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.getCell(selectedRow, selectedCol);

  clipboard = {
    row: selectedRow,
    col: selectedCol,
    value: cell?.value,
    formula: cell?.formula?.formula,
    style: cell?.style ? JSON.parse(JSON.stringify(cell.style)) : undefined,
    comment: cell?.comment ? JSON.parse(JSON.stringify(cell.comment)) : undefined,
    isCut: false
  };
}

function cutCell() {
  copyCell();
  if (clipboard) {
    clipboard.isCut = true;
    // Visual indicator for cut cell
    const td = document.querySelector(`td[data-row="${selectedRow}"][data-col="${selectedCol}"]`);
    if (td) {
      td.style.opacity = '0.5';
    }
  }
}

function pasteCell() {
  if (!clipboard || selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  // Paste value or formula
  if (clipboard.formula) {
    cell.setFormula(clipboard.formula);
  } else if (clipboard.value !== undefined) {
    cell.value = clipboard.value;
  }

  // Paste style
  if (clipboard.style) {
    cell.style = JSON.parse(JSON.stringify(clipboard.style));
  }

  // Paste comment
  if (clipboard.comment) {
    const text = typeof clipboard.comment.text === 'string'
      ? clipboard.comment.text
      : clipboard.comment.text.plainText;
    cell.setComment(text, clipboard.comment.author);
  }

  // Clear source if cut
  if (clipboard.isCut) {
    const sourceCell = sheet.cell(clipboard.row, clipboard.col);
    sourceCell.clear();
    updateCellDisplay(clipboard.row, clipboard.col);
    clipboard = null;
  }

  updateCellDisplay(selectedRow, selectedCol);
}

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
  if (e.ctrlKey || e.metaKey) {
    if (e.key === 'c') {
      e.preventDefault();
      copyCell();
    } else if (e.key === 'x') {
      e.preventDefault();
      cutCell();
    } else if (e.key === 'v') {
      e.preventDefault();
      pasteCell();
    }
  }
});
```

## Context Menu

### Creating a Context Menu

```html
<div class="context-menu" id="contextMenu">
  <div class="context-menu-item" data-action="copy">Copy</div>
  <div class="context-menu-item" data-action="cut">Cut</div>
  <div class="context-menu-item" data-action="paste">Paste</div>
  <div class="context-menu-divider"></div>
  <div class="context-menu-item" data-action="bold">Bold</div>
  <div class="context-menu-item" data-action="italic">Italic</div>
  <div class="context-menu-divider"></div>
  <div class="context-menu-item" data-action="add-comment">Add Comment</div>
  <div class="context-menu-item" data-action="clear">Clear Cell</div>
</div>
```

```css
.context-menu {
  position: fixed;
  background: white;
  border: 1px solid #ddd;
  border-radius: 8px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.15);
  min-width: 160px;
  z-index: 1000;
  display: none;
}

.context-menu.visible {
  display: block;
}

.context-menu-item {
  padding: 8px 16px;
  cursor: pointer;
}

.context-menu-item:hover {
  background: #f3f4f6;
}

.context-menu-divider {
  height: 1px;
  background: #e5e7eb;
  margin: 4px 0;
}
```

```typescript
const contextMenu = document.getElementById('contextMenu')!;

function showContextMenu(e: MouseEvent, row: number, col: number) {
  e.preventDefault();
  selectCell(row, col);

  contextMenu.style.left = e.clientX + 'px';
  contextMenu.style.top = e.clientY + 'px';
  contextMenu.classList.add('visible');
}

function hideContextMenu() {
  contextMenu.classList.remove('visible');
}

// Show on right-click
table.addEventListener('contextmenu', (e) => {
  const td = (e.target as Element).closest('td[data-row][data-col]');
  if (td) {
    showContextMenu(
      e,
      parseInt(td.dataset.row!),
      parseInt(td.dataset.col!)
    );
  }
});

// Hide on click outside
document.addEventListener('click', (e) => {
  if (!contextMenu.contains(e.target as Node)) {
    hideContextMenu();
  }
});

// Handle menu actions
contextMenu.addEventListener('click', (e) => {
  const item = (e.target as Element).closest('.context-menu-item');
  if (!item) return;

  const action = item.dataset.action;

  switch (action) {
    case 'copy': copyCell(); break;
    case 'cut': cutCell(); break;
    case 'paste': pasteCell(); break;
    case 'bold': toggleBold(); break;
    case 'italic': toggleItalic(); break;
    case 'add-comment': showCommentDialog(); break;
    case 'clear': clearCell(); break;
  }

  hideContextMenu();
});
```

## Cell Formatting

### Toggle Bold/Italic

```typescript
function toggleBold() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};
  if (!cell.style.font) cell.style.font = {};

  cell.style.font.bold = !cell.style.font.bold;
  updateCellDisplay(selectedRow, selectedCol);
}

function toggleItalic() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};
  if (!cell.style.font) cell.style.font = {};

  cell.style.font.italic = !cell.style.font.italic;
  updateCellDisplay(selectedRow, selectedCol);
}

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
  if (e.ctrlKey || e.metaKey) {
    if (e.key === 'b') {
      e.preventDefault();
      toggleBold();
    } else if (e.key === 'i') {
      e.preventDefault();
      toggleItalic();
    }
  }
});
```

### Setting Colors

```typescript
function setFillColor(color: string | null) {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};

  if (color) {
    cell.style.fill = {
      type: 'pattern',
      pattern: 'solid',
      foregroundColor: color
    };
  } else {
    delete cell.style.fill;
  }

  updateCellDisplay(selectedRow, selectedCol);
}

function setTextColor(color: string) {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};
  if (!cell.style.font) cell.style.font = {};

  cell.style.font.color = color;
  updateCellDisplay(selectedRow, selectedCol);
}
```

## Comments

### Adding and Managing Comments

```typescript
function showCommentDialog() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.getCell(selectedRow, selectedCol);

  // Get existing comment text
  let currentText = '';
  if (cell?.comment) {
    currentText = typeof cell.comment.text === 'string'
      ? cell.comment.text
      : cell.comment.text.plainText || '';
  }

  const text = prompt('Enter comment:', currentText);
  if (text === null) return;

  if (text.trim() === '') {
    deleteComment();
    return;
  }

  const author = prompt('Author (optional):') || undefined;

  const targetCell = sheet.cell(selectedRow, selectedCol);
  targetCell.setComment(text, author);
  updateCellDisplay(selectedRow, selectedCol);
}

function deleteComment() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);
  cell.comment = undefined;
  updateCellDisplay(selectedRow, selectedCol);
}
```

### Comment Indicator

```css
/* Yellow triangle indicator */
td.has-comment::before {
  content: '';
  position: absolute;
  top: 0;
  right: 0;
  width: 0;
  height: 0;
  border-style: solid;
  border-width: 0 8px 8px 0;
  border-color: transparent #F59E0B transparent transparent;
}
```

## Column Filters

### Filter Implementation

```typescript
interface FilterState {
  [colIndex: number]: Set<string>;
}

const activeFilters: FilterState = {};

function showFilterDropdown(colIndex: number) {
  const sheet = workbook.sheets[currentSheetIndex];
  const dims = sheet.dimensions;
  if (!dims) return;

  // Collect unique values
  const values = new Map<string, number>();
  for (let r = dims.startRow; r <= dims.endRow; r++) {
    const cell = sheet.getCell(r, colIndex);
    let value = cell?.value;

    if (value instanceof Date) {
      value = value.toLocaleDateString();
    } else if (value === null || value === undefined) {
      value = '(Empty)';
    } else {
      value = String(value);
    }

    values.set(value, (values.get(value) || 0) + 1);
  }

  // Build dropdown UI with checkboxes for each value
  // ... render filter dropdown
}

function applyFilters() {
  const sheet = workbook.sheets[currentSheetIndex];
  const tbody = document.querySelector('tbody')!;
  const rows = tbody.querySelectorAll('tr');

  rows.forEach(tr => {
    const rowIndex = parseInt(tr.dataset.row!);
    let visible = true;

    // Check each active filter
    for (const [colIndex, allowedValues] of Object.entries(activeFilters)) {
      const cell = sheet.getCell(rowIndex, parseInt(colIndex));
      let value = formatCellValue(cell?.value);

      if (!allowedValues.has(value)) {
        visible = false;
        break;
      }
    }

    tr.classList.toggle('filtered-out', !visible);
  });
}
```

```css
tr.filtered-out {
  display: none !important;
}
```

## Updating Cell Display

### Refresh Cell After Changes

```typescript
function updateCellDisplay(row: number, col: number) {
  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (!td) return;

  const sheet = workbook.sheets[currentSheetIndex];
  const cell = sheet.getCell(row, col);

  // Update value
  td.textContent = formatCellValue(cell?.value);

  // Update inline styles from cell style
  td.style.cssText = cellStyleToCss(cell?.style);

  // Update comment indicator
  td.classList.toggle('has-comment', !!cell?.comment);

  // Update title/tooltip
  if (cell?.comment) {
    const text = typeof cell.comment.text === 'string'
      ? cell.comment.text
      : cell.comment.text.plainText;
    td.title = `Comment: ${text}`;
  } else {
    td.title = '';
  }
}

function formatCellValue(value: any): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toLocaleDateString();
  if (typeof value === 'number') return value.toLocaleString();
  return String(value);
}

function cellStyleToCss(style?: CellStyle): string {
  if (!style) return '';

  let css = '';

  if (style.font) {
    if (style.font.bold) css += 'font-weight: bold;';
    if (style.font.italic) css += 'font-style: italic;';
    if (style.font.color) css += `color: ${style.font.color};`;
  }

  if (style.fill?.type === 'pattern' && style.fill.pattern === 'solid') {
    css += `background-color: ${style.fill.foregroundColor};`;
  }

  if (style.alignment) {
    if (style.alignment.horizontal) {
      css += `text-align: ${style.alignment.horizontal};`;
    }
  }

  return css;
}
```

## Complete Example

See the full implementation in [`demo/index.html`](https://github.com/user/cellify/blob/main/demo/index.html) which includes:

- Cell selection and navigation
- Inline editing with Enter/Escape
- Undo/redo stack
- Right-click context menu
- Copy/cut/paste
- Cell formatting (colors, bold, italic)
- Comment management
- Column filters
- Column/row resizing

The demo serves as a reference for building your own spreadsheet UI with Cellify.

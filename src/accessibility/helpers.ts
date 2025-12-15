import type { Cell } from '../core/Cell.js';
import type { Sheet } from '../core/Sheet.js';
import type {
  CellAccessibility,
  SheetAccessibility,
  Announcement,
  AnnounceType,
} from './types.js';
import { addressToA1 } from '../types/cell.types.js';

/**
 * Generate accessibility metadata for a cell
 */
export function getCellAccessibility(
  cell: Cell,
  sheet: Sheet,
  options: {
    headerRows?: number;
    headerCols?: number;
    includePosition?: boolean;
  } = {}
): CellAccessibility {
  const { headerRows = 0, headerCols = 0, includePosition = true } = options;
  const a11y: CellAccessibility = {};

  // Determine if this is a header cell
  const isRowHeader = cell.col < headerCols;
  const isColHeader = cell.row < headerRows;
  a11y.isHeader = isRowHeader || isColHeader;

  // Set scope for header cells
  if (isColHeader && isRowHeader) {
    // Corner cell - could be either, typically column
    a11y.scope = 'col';
  } else if (isColHeader) {
    a11y.scope = 'col';
  } else if (isRowHeader) {
    a11y.scope = 'row';
  }

  // Set role
  if (isColHeader) {
    a11y.role = 'columnheader';
  } else if (isRowHeader) {
    a11y.role = 'rowheader';
  } else {
    a11y.role = 'gridcell';
  }

  // Generate header references for data cells
  if (!a11y.isHeader && (headerRows > 0 || headerCols > 0)) {
    a11y.headers = [];

    // Add column header reference
    if (headerRows > 0) {
      for (let r = 0; r < headerRows; r++) {
        a11y.headers.push(`cell-${r}-${cell.col}`);
      }
    }

    // Add row header reference
    if (headerCols > 0) {
      for (let c = 0; c < headerCols; c++) {
        a11y.headers.push(`cell-${cell.row}-${c}`);
      }
    }
  }

  // Position information
  if (includePosition) {
    a11y.ariaColIndex = cell.col + 1; // ARIA uses 1-based indexing
    a11y.ariaRowIndex = cell.row + 1;
  }

  // Handle merged cells
  if (cell.merge) {
    a11y.ariaColSpan = cell.merge.endCol - cell.merge.startCol + 1;
    a11y.ariaRowSpan = cell.merge.endRow - cell.merge.startRow + 1;
  }

  // Read-only if cell has no validation allowing input
  // or if sheet is protected and cell is locked
  const protection = sheet.protection;
  const cellProtection = cell.style?.protection;
  if (protection?.sheet && cellProtection?.locked !== false) {
    a11y.ariaReadOnly = true;
  }

  // Handle validation
  if (cell.validation) {
    if (cell.validation.showErrorMessage) {
      // Cell has validation - could be invalid
      // Actual invalid state would be determined by evaluating the value
    }
    if (cell.validation.type === 'list' && cell.validation.showDropDown) {
      a11y.ariaHasPopup = 'listbox';
    }
  }

  // Generate value text for formatted numbers
  const valueText = getValueText(cell);
  if (valueText && valueText !== String(cell.value)) {
    a11y.ariaValueText = valueText;
  }

  return a11y;
}

/**
 * Generate human-readable value text for a cell
 */
export function getValueText(cell: Cell): string | undefined {
  const value = cell.value;
  if (value === null || value === undefined) {
    return 'empty';
  }

  const format = cell.style?.numberFormat;

  // Handle different value types
  if (typeof value === 'boolean') {
    return value ? 'true' : 'false';
  }

  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  if (typeof value === 'number') {
    // Check for percentage format
    if (format?.formatCode?.includes('%')) {
      return `${(value * 100).toFixed(0)} percent`;
    }

    // Check for currency format
    if (format?.formatCode?.includes('$')) {
      return `${value.toFixed(2)} dollars`;
    }

    // Check for accounting format
    if (format?.category === 'accounting') {
      return `${value.toFixed(2)} in accounting format`;
    }

    return String(value);
  }

  // Handle error values
  if (typeof value === 'string' && value.startsWith('#')) {
    const errorDescriptions: Record<string, string> = {
      '#NULL!': 'null error',
      '#DIV/0!': 'division by zero error',
      '#VALUE!': 'value error',
      '#REF!': 'reference error',
      '#NAME?': 'name error',
      '#NUM!': 'number error',
      '#N/A': 'not available',
      '#GETTING_DATA': 'loading data',
    };
    return errorDescriptions[value] || 'error';
  }

  // Handle rich text
  if (typeof value === 'object' && 'richText' in value) {
    return value.richText.map((run) => run.text).join('');
  }

  return String(value);
}

/**
 * Generate accessibility metadata for a sheet
 */
export function getSheetAccessibility(
  sheet: Sheet,
  options: {
    headerRows?: number;
    headerCols?: number;
  } = {}
): SheetAccessibility {
  const { headerRows = 0, headerCols = 0 } = options;
  const dimensions = sheet.dimensions;

  const a11y: SheetAccessibility = {
    label: sheet.name,
    ariaMultiSelectable: true,
  };

  if (dimensions) {
    a11y.ariaRowCount = dimensions.endRow - dimensions.startRow + 1;
    a11y.ariaColCount = dimensions.endCol - dimensions.startCol + 1;
  }

  if (headerRows > 0) {
    a11y.headerRowStart = 0;
    a11y.headerRowEnd = headerRows - 1;
  }

  if (headerCols > 0) {
    a11y.headerColStart = 0;
    a11y.headerColEnd = headerCols - 1;
  }

  return a11y;
}

/**
 * Generate a cell description for screen readers
 */
export function describeCellPosition(row: number, col: number): string {
  const address = addressToA1(row, col);
  return `Cell ${address}, row ${row + 1}, column ${col + 1}`;
}

/**
 * Generate a cell description including value
 */
export function describeCellFull(cell: Cell, _sheet: Sheet): string {
  const position = describeCellPosition(cell.row, cell.col);
  const valueText = getValueText(cell);

  let description = `${position}, ${valueText || 'empty'}`;

  // Add merge info
  if (cell.isMergeMaster && cell.merge) {
    const cols = cell.merge.endCol - cell.merge.startCol + 1;
    const rows = cell.merge.endRow - cell.merge.startRow + 1;
    description += `, merged ${rows} rows by ${cols} columns`;
  }

  // Add formula info
  if (cell.formula) {
    description += `, formula: ${cell.formula.formula}`;
  }

  // Add comment info
  if (cell.comment) {
    description += ', has comment';
  }

  // Add validation info
  if (cell.validation) {
    description += ', has data validation';
  }

  return description;
}

/**
 * Create an announcement for screen readers
 */
export function createAnnouncement(
  message: string,
  type: AnnounceType,
  priority: 'polite' | 'assertive' = 'polite'
): Announcement {
  return { message, type, priority };
}

/**
 * Generate navigation announcement
 */
export function announceNavigation(cell: Cell): Announcement {
  const position = describeCellPosition(cell.row, cell.col);
  const value = getValueText(cell) || 'empty';
  return createAnnouncement(`${position}, ${value}`, 'navigation', 'polite');
}

/**
 * Generate selection announcement
 */
export function announceSelection(
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): Announcement {
  if (startRow === endRow && startCol === endCol) {
    return createAnnouncement(
      `Selected cell ${addressToA1(startRow, startCol)}`,
      'selection',
      'polite'
    );
  }

  const rows = endRow - startRow + 1;
  const cols = endCol - startCol + 1;
  const startAddr = addressToA1(startRow, startCol);
  const endAddr = addressToA1(endRow, endCol);

  return createAnnouncement(
    `Selected range ${startAddr} to ${endAddr}, ${rows} rows by ${cols} columns`,
    'selection',
    'polite'
  );
}

/**
 * Generate error announcement
 */
export function announceError(message: string): Announcement {
  return createAnnouncement(message, 'error', 'assertive');
}

/**
 * Generate success announcement
 */
export function announceSuccess(message: string): Announcement {
  return createAnnouncement(message, 'success', 'polite');
}

/**
 * Generate ARIA attributes object for rendering
 */
export function getAriaAttributes(a11y: CellAccessibility): Record<string, string | number | boolean> {
  const attrs: Record<string, string | number | boolean> = {};

  if (a11y.role) attrs['role'] = a11y.role;
  if (a11y.ariaLabel) attrs['aria-label'] = a11y.ariaLabel;
  if (a11y.ariaDescribedBy) attrs['aria-describedby'] = a11y.ariaDescribedBy;
  if (a11y.ariaSelected !== undefined) attrs['aria-selected'] = a11y.ariaSelected;
  if (a11y.ariaReadOnly !== undefined) attrs['aria-readonly'] = a11y.ariaReadOnly;
  if (a11y.ariaRequired !== undefined) attrs['aria-required'] = a11y.ariaRequired;
  if (a11y.ariaInvalid !== undefined) attrs['aria-invalid'] = a11y.ariaInvalid;
  if (a11y.ariaValueText) attrs['aria-valuetext'] = a11y.ariaValueText;
  if (a11y.ariaHasPopup !== undefined) attrs['aria-haspopup'] = a11y.ariaHasPopup;
  if (a11y.ariaExpanded !== undefined) attrs['aria-expanded'] = a11y.ariaExpanded;
  if (a11y.ariaLevel !== undefined) attrs['aria-level'] = a11y.ariaLevel;
  if (a11y.ariaPosInSet !== undefined) attrs['aria-posinset'] = a11y.ariaPosInSet;
  if (a11y.ariaSetSize !== undefined) attrs['aria-setsize'] = a11y.ariaSetSize;
  if (a11y.ariaColIndex !== undefined) attrs['aria-colindex'] = a11y.ariaColIndex;
  if (a11y.ariaRowIndex !== undefined) attrs['aria-rowindex'] = a11y.ariaRowIndex;
  if (a11y.ariaColSpan !== undefined) attrs['aria-colspan'] = a11y.ariaColSpan;
  if (a11y.ariaRowSpan !== undefined) attrs['aria-rowspan'] = a11y.ariaRowSpan;
  if (a11y.headers?.length) attrs['headers'] = a11y.headers.join(' ');
  if (a11y.scope) attrs['scope'] = a11y.scope;

  return attrs;
}

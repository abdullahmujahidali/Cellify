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
 * Build accessibility metadata for a specific cell based on sheet context and header configuration.
 *
 * Generates ARIA-related properties such as role and scope, header references for data cells,
 * 1-based ARIA row/column indices, span values for merged cells, read-only state, popup hints for list validations,
 * and a human-readable value text when different from the raw value.
 *
 * @param cell - The cell to produce accessibility metadata for
 * @param sheet - The sheet containing the cell; used for protection and dimension/contextual checks
 * @param options.headerRows - Number of top rows treated as column headers (default: 0)
 * @param options.headerCols - Number of left columns treated as row headers (default: 0)
 * @param options.includePosition - If true, include `ariaRowIndex` and `ariaColIndex` (default: true)
 * @returns A CellAccessibility object containing computed ARIA attributes and related accessibility hints
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
 * Produce a human-readable textual representation of a cell's value.
 *
 * The result normalizes common value types for spoken/output-friendly text:
 * - `null` or `undefined` -> "empty"
 * - booleans -> "true" or "false"
 * - Date -> locale-formatted date string
 * - numbers -> plain numeric string, or formatted as a percentage, currency, or accounting when the cell's number format indicates those
 * - error codes beginning with `#` -> a descriptive error phrase when known (e.g. "#DIV/0!" -> "division by zero error"), otherwise "error"
 * - rich text objects with a `richText` array -> concatenated run text
 * - fallback -> stringified value
 *
 * @returns A human-readable string describing the cell's value, or `undefined` only when no representation can be produced.
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
 * Builds accessibility metadata for a sheet, including label, counts, and header ranges.
 *
 * @param sheet - Sheet to derive accessibility information from (uses sheet.name and sheet.dimensions when available)
 * @param options - Optional settings for header regions
 * @param options.headerRows - Number of top rows treated as header rows (if > 0, `headerRowStart` and `headerRowEnd` will be set)
 * @param options.headerCols - Number of left columns treated as header columns (if > 0, `headerColStart` and `headerColEnd` will be set)
 * @returns An object describing sheet-level accessibility: `label` (sheet name), `ariaMultiSelectable`, optional `ariaRowCount`/`ariaColCount` derived from dimensions, and optional header start/end indices for rows and columns
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
 * Create a human-readable description of a cell's position.
 *
 * @returns A string containing the cell's A1 address and its 1-based row and column numbers (for example, "Cell A1, row 1, column 1").
 */
export function describeCellPosition(row: number, col: number): string {
  const address = addressToA1(row, col);
  return `Cell ${address}, row ${row + 1}, column ${col + 1}`;
}

/**
 * Produce a spoken description of a cell's position, value, and metadata.
 *
 * @returns A string describing the cell's A1 address and row/column indices, the cell value or 'empty', and, when present, merge dimensions, formula text, comment presence, and data validation.
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
 * Build an announcement object for screen readers.
 *
 * @param message - The text to announce
 * @param type - The announcement category (e.g., 'navigation', 'selection', 'error', 'success')
 * @param priority - The ARIA live priority for the announcement; defaults to 'polite'
 * @returns An Announcement object containing the provided message, type, and priority
 */
export function createAnnouncement(
  message: string,
  type: AnnounceType,
  priority: 'polite' | 'assertive' = 'polite'
): Announcement {
  return { message, type, priority };
}

/**
 * Create a navigation announcement describing a cell's position and readable value.
 *
 * @returns An Announcement containing a message of the form "Cell {A1}, row {n}, column {m}, {value}", with type `navigation` and priority `polite`.
 */
export function announceNavigation(cell: Cell): Announcement {
  const position = describeCellPosition(cell.row, cell.col);
  const value = getValueText(cell) || 'empty';
  return createAnnouncement(`${position}, ${value}`, 'navigation', 'polite');
}

/**
 * Create an announcement describing a single cell or rectangular range selection for screen readers.
 *
 * @param startRow - Zero-based index of the first (top) row in the selection
 * @param startCol - Zero-based index of the first (left) column in the selection
 * @param endRow - Zero-based index of the last (bottom) row in the selection
 * @param endCol - Zero-based index of the last (right) column in the selection
 * @returns An Announcement containing a human-readable message for the selected cell or range
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
 * Create an accessibility announcement representing an error.
 *
 * @param message - The error message to announce
 * @returns An Announcement with type `'error'` and priority `'assertive'`
 */
export function announceError(message: string): Announcement {
  return createAnnouncement(message, 'error', 'assertive');
}

/**
 * Create a success announcement for screen readers.
 *
 * @param message - The announcement text to be spoken
 * @returns An `Announcement` with type `'success'` and priority `'polite'`
 */
export function announceSuccess(message: string): Announcement {
  return createAnnouncement(message, 'success', 'polite');
}

/**
 * Convert a CellAccessibility descriptor into a flat map of DOM ARIA attributes for rendering.
 *
 * @param a11y - Accessibility metadata for a cell
 * @returns An object mapping attribute names (e.g., `role`, `aria-label`, `aria-rowindex`) to their values; when `headers` is present it is returned as a single space-separated string
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
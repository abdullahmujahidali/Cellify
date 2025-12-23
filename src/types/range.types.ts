import type { CellAddress, CellValue } from './cell.types.js';
import type { CellStyle } from './style.types.js';

/**
 * Range definition - a rectangular area of cells
 */
export interface RangeDefinition {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/**
 * Range reference in A1 notation (e.g., "A1:B10", "Sheet1!A1:B10")
 */
export type RangeReference = string;

/**
 * Parse a range reference like "A1:B10" or "A1"
 */
export function parseRangeReference(ref: string): RangeDefinition {
  // Remove sheet name if present
  const refWithoutSheet = ref.includes('!') ? ref.split('!')[1] : ref;

  // Check if it's a single cell or a range
  if (refWithoutSheet.includes(':')) {
    const [start, end] = refWithoutSheet.split(':');
    const startAddr = parseA1Reference(start);
    const endAddr = parseA1Reference(end);

    return {
      startRow: Math.min(startAddr.row, endAddr.row),
      startCol: Math.min(startAddr.col, endAddr.col),
      endRow: Math.max(startAddr.row, endAddr.row),
      endCol: Math.max(startAddr.col, endAddr.col),
    };
  }

  // Single cell
  const addr = parseA1Reference(refWithoutSheet);
  return {
    startRow: addr.row,
    startCol: addr.col,
    endRow: addr.row,
    endCol: addr.col,
  };
}

/**
 * Parse a single A1 reference
 */
function parseA1Reference(ref: string): CellAddress {
  const match = ref.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }

  let col = 0;
  const colStr = match[1].toUpperCase();
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }

  return {
    row: parseInt(match[2], 10) - 1,
    col: col - 1,
  };
}

/**
 * Convert range definition to A1 notation
 */
export function rangeToA1(range: RangeDefinition): string {
  const startCol = columnIndexToLetter(range.startCol);
  const endCol = columnIndexToLetter(range.endCol);

  if (
    range.startRow === range.endRow &&
    range.startCol === range.endCol
  ) {
    return `${startCol}${range.startRow + 1}`;
  }

  return `${startCol}${range.startRow + 1}:${endCol}${range.endRow + 1}`;
}

/**
 * Convert column index to letter
 */
function columnIndexToLetter(index: number): string {
  let letter = '';
  let temp = index;

  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }

  return letter;
}

/**
 * Check if two ranges overlap
 */
export function rangesOverlap(a: RangeDefinition, b: RangeDefinition): boolean {
  return !(
    a.endRow < b.startRow ||
    a.startRow > b.endRow ||
    a.endCol < b.startCol ||
    a.startCol > b.endCol
  );
}

/**
 * Check if a cell is within a range
 */
export function isCellInRange(cell: CellAddress, range: RangeDefinition): boolean {
  return (
    cell.row >= range.startRow &&
    cell.row <= range.endRow &&
    cell.col >= range.startCol &&
    cell.col <= range.endCol
  );
}

/**
 * Get the intersection of two ranges, or null if they don't overlap
 */
export function getRangeIntersection(
  a: RangeDefinition,
  b: RangeDefinition
): RangeDefinition | null {
  if (!rangesOverlap(a, b)) {
    return null;
  }

  return {
    startRow: Math.max(a.startRow, b.startRow),
    startCol: Math.max(a.startCol, b.startCol),
    endRow: Math.min(a.endRow, b.endRow),
    endCol: Math.min(a.endCol, b.endCol),
  };
}

/**
 * Get the union (bounding box) of two ranges
 */
export function getRangeUnion(a: RangeDefinition, b: RangeDefinition): RangeDefinition {
  return {
    startRow: Math.min(a.startRow, b.startRow),
    startCol: Math.min(a.startCol, b.startCol),
    endRow: Math.max(a.endRow, b.endRow),
    endCol: Math.max(a.endCol, b.endCol),
  };
}

/**
 * Iterator for all cells in a range
 */
export function* iterateRange(range: RangeDefinition): Generator<CellAddress> {
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      yield { row, col };
    }
  }
}

/**
 * Get dimensions of a range
 */
export function getRangeDimensions(range: RangeDefinition): { rows: number; cols: number } {
  return {
    rows: range.endRow - range.startRow + 1,
    cols: range.endCol - range.startCol + 1,
  };
}

/**
 * Conditional formatting rule types
 */
export type ConditionalFormatType =
  | 'cellIs'
  | 'containsText'
  | 'timePeriod'
  | 'aboveAverage'
  | 'top10'
  | 'uniqueValues'
  | 'duplicateValues'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'expression';

/**
 * Conditional formatting rule
 */
export interface ConditionalFormatRule {
  type: ConditionalFormatType;
  priority: number;
  ranges: RangeDefinition[];
  stopIfTrue?: boolean;

  // For cellIs, containsText, expression
  operator?: string;
  formula?: string | string[];

  // Style to apply when condition is met
  style?: CellStyle;

  // For color scales
  colorScale?: {
    minColor: string;
    midColor?: string;
    maxColor: string;
    minType?: 'min' | 'num' | 'percent' | 'percentile' | 'formula';
    midType?: 'num' | 'percent' | 'percentile' | 'formula';
    maxType?: 'max' | 'num' | 'percent' | 'percentile' | 'formula';
    minValue?: number | string;
    midValue?: number | string;
    maxValue?: number | string;
  };

  // For data bars
  dataBar?: {
    color: string;
    showValue?: boolean;
    minLength?: number;
    maxLength?: number;
  };

  // For icon sets
  iconSet?: {
    iconSet: string;
    reverse?: boolean;
    showValue?: boolean;
  };
}

/**
 * Auto filter definition
 */
export interface AutoFilter {
  range: RangeDefinition;
  columns?: AutoFilterColumn[];
}

/**
 * Paste options for sheet.pasteRange() method
 */
export interface PasteOptions {
  /** Paste only values, ignore styles */
  valuesOnly?: boolean;
  /** Paste only styles, ignore values */
  stylesOnly?: boolean;
  /** Transpose rows and columns when pasting */
  transpose?: boolean;
}

/**
 * Search options for sheet.find() and sheet.findAll() methods
 */
export interface SearchOptions {
  /** The search query (string or number) */
  query?: string | number;
  /** Regular expression to match */
  regex?: RegExp;
  /** Match case when searching (default: false) */
  matchCase?: boolean;
  /** Match entire cell content (default: false, matches partial) */
  matchCell?: boolean;
  /** Where to search: 'values', 'formulas', or 'both' (default: 'values') */
  searchIn?: 'values' | 'formulas' | 'both';
  /** Limit search to specific range */
  range?: string | RangeDefinition;
}

/**
 * Filter criteria for sheet.filter() method
 */
export interface FilterCriteria {
  // Equality
  equals?: string | number | boolean | null;
  notEquals?: string | number | boolean | null;

  // String operations (case-insensitive)
  contains?: string;
  notContains?: string;
  startsWith?: string;
  endsWith?: string;

  // Numeric operations
  greaterThan?: number;
  greaterThanOrEqual?: number;
  lessThan?: number;
  lessThanOrEqual?: number;
  between?: [number, number];
  notBetween?: [number, number];

  // Value list
  in?: (string | number | boolean | null)[];
  notIn?: (string | number | boolean | null)[];

  // Empty checks
  isEmpty?: boolean;
  isNotEmpty?: boolean;

  // Custom function
  custom?: (value: CellValue) => boolean;
}

/**
 * Auto filter column configuration
 */
export interface AutoFilterColumn {
  columnIndex: number;
  filterType?: 'custom' | 'top10' | 'dynamic' | 'color';
  values?: (string | number | boolean)[];
  customFilters?: {
    operator: string;
    value: string | number;
  }[];
}

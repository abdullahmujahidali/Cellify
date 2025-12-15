import type { CellStyle } from './style.types.js';

/**
 * Supported cell value types
 */
export type CellValueType = 'string' | 'number' | 'boolean' | 'date' | 'error' | 'formula' | 'null';

/**
 * Excel error types
 */
export type CellErrorType =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA';

/**
 * Rich text run - a segment of text with its own formatting
 */
export interface RichTextRun {
  text: string;
  font?: {
    name?: string;
    size?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
  };
}

/**
 * Rich text value - text with multiple formatting runs
 */
export interface RichTextValue {
  richText: RichTextRun[];
}

/**
 * Hyperlink definition
 */
export interface CellHyperlink {
  target: string; // URL, file path, or internal reference
  tooltip?: string;
  display?: string; // Display text (if different from cell value)
}

/**
 * Cell comment/note
 */
export interface CellComment {
  text: string | RichTextValue;
  author?: string;
  visible?: boolean; // Whether comment is always visible
}

/**
 * Data validation types
 */
export type ValidationType =
  | 'whole'
  | 'decimal'
  | 'list'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom';

/**
 * Data validation operator
 */
export type ValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual';

/**
 * Data validation error style
 */
export type ValidationErrorStyle = 'stop' | 'warning' | 'information';

/**
 * Cell data validation rule
 */
export interface CellValidation {
  type: ValidationType;
  operator?: ValidationOperator;
  formula1?: string | number | Date; // First value or formula
  formula2?: string | number | Date; // Second value (for between/notBetween)
  allowBlank?: boolean;
  showDropDown?: boolean; // For list validation
  showInputMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  showErrorMessage?: boolean;
  errorStyle?: ValidationErrorStyle;
  errorTitle?: string;
  errorMessage?: string;
}

/**
 * Primitive cell values
 */
export type PrimitiveCellValue = string | number | boolean | Date | null;

/**
 * All possible cell values including rich text
 */
export type CellValue = PrimitiveCellValue | RichTextValue | CellErrorType;

/**
 * Formula with optional cached result
 */
export interface CellFormula {
  formula: string; // Formula text without leading '='
  result?: CellValue; // Cached result
  sharedIndex?: number; // For shared formulas
}

/**
 * Cell address in A1 notation
 */
export interface CellAddress {
  row: number; // 0-based
  col: number; // 0-based
}

/**
 * Cell reference string (e.g., "A1", "$B$2")
 */
export type CellReference = string;

/**
 * Cell merge definition
 */
export interface MergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/**
 * Cell data - the complete representation of a cell
 */
export interface CellData {
  value: CellValue;
  type: CellValueType;
  formula?: CellFormula;
  style?: CellStyle;
  hyperlink?: CellHyperlink;
  comment?: CellComment;
  validation?: CellValidation;
  // Merge info - only set on the top-left cell of a merge
  merge?: MergeRange;
  // If this cell is part of a merge but not the master cell
  mergedInto?: CellAddress;
}

/**
 * Sparse cell storage - only stores cells that have data
 * Key format: "row,col" e.g., "0,0" for A1
 */
export type CellStorage = Map<string, CellData>;

/**
 * Convert column index to letter (0 -> A, 25 -> Z, 26 -> AA)
 */
export function columnIndexToLetter(index: number): string {
  let letter = '';
  let temp = index;

  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }

  return letter;
}

/**
 * Convert column letter to index (A -> 0, Z -> 25, AA -> 26)
 */
export function columnLetterToIndex(letter: string): number {
  let index = 0;
  const upper = letter.toUpperCase();

  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }

  return index - 1;
}

/**
 * Convert cell address to A1 notation
 */
export function addressToA1(row: number, col: number): string {
  return `${columnIndexToLetter(col)}${row + 1}`;
}

/**
 * Parse A1 notation to cell address
 */
export function a1ToAddress(a1: string): CellAddress {
  const match = a1.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) {
    throw new Error(`Invalid cell reference: ${a1}`);
  }

  return {
    col: columnLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

/**
 * Generate storage key from row and column
 */
export function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

/**
 * Parse storage key to row and column
 */
export function parseKey(key: string): CellAddress {
  const [row, col] = key.split(',').map(Number);
  return { row, col };
}

/**
 * Determine the value type of a cell value
 */
export function getCellValueType(value: CellValue): CellValueType {
  if (value === null || value === undefined) {
    return 'null';
  }
  if (typeof value === 'string') {
    // Check if it's an error
    if (
      value.startsWith('#') &&
      ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'].includes(
        value as CellErrorType
      )
    ) {
      return 'error';
    }
    return 'string';
  }
  if (typeof value === 'number') {
    return 'number';
  }
  if (typeof value === 'boolean') {
    return 'boolean';
  }
  if (value instanceof Date) {
    return 'date';
  }
  if (typeof value === 'object' && 'richText' in value) {
    return 'string'; // Rich text is treated as string type
  }
  return 'null';
}

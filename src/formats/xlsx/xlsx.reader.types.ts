/**
 * XLSX Import types and options
 */

import type { Workbook } from '../../core/Workbook.js';

/**
 * Progress phase during import
 */
export type XlsxImportPhase =
  | 'unzip'
  | 'sharedStrings'
  | 'styles'
  | 'sheets'
  | 'properties';

/**
 * Progress callback for long-running imports
 */
export type XlsxProgressCallback = (
  phase: XlsxImportPhase,
  current: number,
  total: number
) => void;

/**
 * Options for XLSX import
 */
export interface XlsxImportOptions {
  /**
   * Which sheets to import
   * - 'all': Import all sheets (default)
   * - string[]: Import sheets by name
   * - number[]: Import sheets by index (0-based)
   * @default 'all'
   */
  sheets?: 'all' | string[] | number[];

  /**
   * Import formula text from cells
   * @default true
   */
  importFormulas?: boolean;

  /**
   * Import cell styles (fonts, colors, borders, etc.)
   * @default true
   */
  importStyles?: boolean;

  /**
   * Import merged cell ranges
   * @default true
   */
  importMergedCells?: boolean;

  /**
   * Import column widths and row heights
   * @default true
   */
  importDimensions?: boolean;

  /**
   * Import freeze pane settings
   * @default true
   */
  importFreezePanes?: boolean;

  /**
   * Import document properties (title, author, etc.)
   * @default true
   */
  importProperties?: boolean;

  /**
   * Import cell comments/notes
   * @default true
   */
  importComments?: boolean;

  /**
   * Maximum rows to import per sheet (0 = unlimited)
   * @default 0
   */
  maxRows?: number;

  /**
   * Maximum columns to import per sheet (0 = unlimited)
   * @default 0
   */
  maxCols?: number;

  /**
   * Progress callback for monitoring import progress
   */
  onProgress?: XlsxProgressCallback;
}

/**
 * Default import options
 */
export const DEFAULT_XLSX_IMPORT_OPTIONS: Required<Omit<XlsxImportOptions, 'onProgress'>> = {
  sheets: 'all',
  importFormulas: true,
  importStyles: true,
  importMergedCells: true,
  importDimensions: true,
  importFreezePanes: true,
  importProperties: true,
  importComments: true,
  maxRows: 0,
  maxCols: 0,
};

/**
 * Warning during import (non-fatal issues)
 */
export interface XlsxImportWarning {
  /** Warning code for programmatic handling */
  code: string;
  /** Human-readable warning message */
  message: string;
  /** Location in workbook (e.g., "Sheet1!A1") */
  location?: string;
}

/**
 * Import statistics
 */
export interface XlsxImportStats {
  /** Number of sheets imported */
  sheetCount: number;
  /** Total cells with values */
  totalCells: number;
  /** Cells containing formulas */
  formulaCells: number;
  /** Merged cell ranges */
  mergedRanges: number;
  /** Import duration in milliseconds */
  durationMs: number;
}

/**
 * Result of XLSX import
 */
export interface XlsxImportResult {
  /** The imported workbook */
  workbook: Workbook;
  /** Import statistics */
  stats: XlsxImportStats;
  /** Non-fatal warnings encountered during import */
  warnings: XlsxImportWarning[];
}

/**
 * Internal parse context for reading XLSX
 */
export interface XlsxParseContext {
  /** Import options */
  options: Required<Omit<XlsxImportOptions, 'onProgress'>>;
  /** Progress callback (if provided) */
  onProgress?: XlsxProgressCallback;
  /** Shared strings table from sharedStrings.xml */
  sharedStrings: string[];
  /** Number format map: numFmtId → formatCode */
  numberFormats: Map<number, string>;
  /** Cell XF styles: index → style info */
  cellXfs: CellXfInfo[];
  /** Font definitions */
  fonts: FontInfo[];
  /** Fill definitions */
  fills: FillInfo[];
  /** Border definitions */
  borders: BorderInfo[];
  /** Sheet info from workbook.xml */
  sheetInfos: SheetInfo[];
  /** Warnings collected during parsing */
  warnings: XlsxImportWarning[];
  /** Statistics */
  stats: XlsxImportStats;
}

/**
 * Sheet info from workbook.xml
 */
export interface SheetInfo {
  name: string;
  sheetId: number;
  rId: string;
}

/**
 * Cell XF info for style lookup
 */
export interface CellXfInfo {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyNumberFormat?: boolean;
  applyAlignment?: boolean;
  alignment?: {
    horizontal?: string;
    vertical?: string;
    wrapText?: boolean;
    textRotation?: number;
  };
}

/**
 * Font info from styles.xml
 */
export interface FontInfo {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  size?: number;
  color?: string;
  name?: string;
}

/**
 * Fill info from styles.xml
 */
export interface FillInfo {
  patternType?: string;
  fgColor?: string;
  bgColor?: string;
}

/**
 * Border side info
 */
export interface BorderSideInfo {
  style?: string;
  color?: string;
}

/**
 * Border info from styles.xml
 */
export interface BorderInfo {
  left?: BorderSideInfo;
  right?: BorderSideInfo;
  top?: BorderSideInfo;
  bottom?: BorderSideInfo;
}

/**
 * Built-in date format IDs in Excel
 * These IDs indicate the cell contains a date/time value
 */
export const DATE_FORMAT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, // Standard date/time formats
  45, 46, 47, // Time formats
  // Note: Custom formats (id >= 164) need pattern inspection
]);

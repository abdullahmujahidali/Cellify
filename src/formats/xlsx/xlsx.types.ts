/**
 * XLSX format types and constants
 */

import type { Workbook } from '../../core/Workbook.js';
import type { CellStyle, CellAlignment } from '../../types/style.types.js';

/**
 * Options for XLSX export
 */
export interface XlsxExportOptions {
  /**
   * Compression level 0-9 (0 = no compression, 9 = max)
   * @default 6
   */
  compressionLevel?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

  /**
   * Include document properties (title, author, etc.)
   * @default true
   */
  includeProperties?: boolean;

  /**
   * Application name for metadata
   * @default 'Cellify'
   */
  application?: string;

  /**
   * Use shared strings table for text values (reduces file size)
   * @default true
   */
  useSharedStrings?: boolean;

  /**
   * Default column width in characters
   * @default 8.43
   */
  defaultColumnWidth?: number;

  /**
   * Default row height in points
   * @default 15
   */
  defaultRowHeight?: number;
}

/**
 * Default export options
 */
export const DEFAULT_XLSX_OPTIONS: Required<XlsxExportOptions> = {
  compressionLevel: 6,
  includeProperties: true,
  application: 'Cellify',
  useSharedStrings: true,
  defaultColumnWidth: 8.43,
  defaultRowHeight: 15,
};

/**
 * Internal cell format record (maps to cellXfs in styles.xml)
 */
export interface CellXf {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignment?: CellAlignment;
  protection?: { locked?: boolean; hidden?: boolean };
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  applyNumberFormat?: boolean;
  applyProtection?: boolean;
}

/**
 * Internal build context passed between generators
 */
export interface XlsxBuildContext {
  workbook: Workbook;
  options: Required<XlsxExportOptions>;

  /** Shared strings table */
  sharedStrings: SharedStringsTable;

  /** Style registry for deduplication */
  styleRegistry: StyleRegistry;

  /** Sheet metadata for relationships */
  sheets: Array<{
    name: string;
    sheetId: number;
    rId: string;
    target: string;
  }>;

  /** Relationship ID for styles.xml */
  stylesRId: string;

  /** Relationship ID for sharedStrings.xml (if used) */
  sharedStringsRId?: string;
}

/**
 * Interface for SharedStringsTable (implemented in xlsx.strings.ts)
 */
export interface SharedStringsTable {
  add(value: string): number;
  readonly count: number;
  readonly uniqueCount: number;
  generateXml(): string;
}

/**
 * Interface for StyleRegistry (implemented in xlsx.styles.ts)
 */
export interface StyleRegistry {
  registerStyle(style: CellStyle | undefined): number;
  generateStylesXml(): string;
}

/**
 * OOXML namespaces
 */
export const NS = {
  spreadsheetml: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  relationships: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  packageRels: 'http://schemas.openxmlformats.org/package/2006/relationships',
} as const;

/**
 * Relationship types
 */
export const REL_TYPES = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  worksheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  sharedStrings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  vmlDrawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
} as const;

/**
 * Content types for OOXML parts
 */
export const CONTENT_TYPES = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  coreProperties: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProperties: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vmlDrawing: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
} as const;

/**
 * Built-in Excel number format IDs
 * IDs 0-163 are reserved for built-in formats
 */
export const BUILTIN_NUM_FMT_IDS: Record<string, number> = {
  'General': 0,
  '0': 1,
  '0.00': 2,
  '#,##0': 3,
  '#,##0.00': 4,
  '0%': 9,
  '0.00%': 10,
  '0.00E+00': 11,
  '# ?/?': 12,
  '# ??/??': 13,
  'mm-dd-yy': 14,
  'd-mmm-yy': 15,
  'd-mmm': 16,
  'mmm-yy': 17,
  'h:mm AM/PM': 18,
  'h:mm:ss AM/PM': 19,
  'h:mm': 20,
  'h:mm:ss': 21,
  'm/d/yy h:mm': 22,
  '#,##0 ;(#,##0)': 37,
  '#,##0 ;[Red](#,##0)': 38,
  '#,##0.00;(#,##0.00)': 39,
  '#,##0.00;[Red](#,##0.00)': 40,
  'mm:ss': 45,
  '[h]:mm:ss': 46,
  'mmss.0': 47,
  '##0.0E+0': 48,
  '@': 49,
  'yyyy-mm-dd': 14, // Common alias
};

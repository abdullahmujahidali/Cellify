/**
 * CSV format options for import and export
 */

/**
 * Options for CSV export
 */
export interface CsvExportOptions {
  /**
   * Field delimiter character
   * @default ','
   */
  delimiter?: string;

  /**
   * Row delimiter (line ending)
   * @default '\r\n' (CRLF for Excel compatibility)
   */
  rowDelimiter?: string;

  /**
   * Quote character for fields containing special characters
   * @default '"'
   */
  quoteChar?: string;

  /**
   * Always quote all fields, not just those requiring quotes
   * @default false
   */
  quoteAllFields?: boolean;

  /**
   * Range to export (e.g., 'A1:D10'). If not specified, exports used range.
   */
  range?: string;

  /**
   * How to handle null/undefined values
   * @default ''
   */
  nullValue?: string;

  /**
   * Date format string for date values
   * @default 'ISO' (yyyy-mm-dd)
   */
  dateFormat?: 'ISO' | 'locale' | string;

  /**
   * Include BOM (Byte Order Mark) for UTF-8
   * Helps Excel recognize UTF-8 encoding
   * @default false
   */
  includeBom?: boolean;

  /**
   * Encoding for the output
   * @default 'utf-8'
   */
  encoding?: 'utf-8' | 'utf-16' | 'ascii';
}

/**
 * Options for CSV import
 */
export interface CsvImportOptions {
  /**
   * Field delimiter character
   * @default ',' (auto-detected if not specified)
   */
  delimiter?: string;

  /**
   * Quote character
   * @default '"'
   */
  quoteChar?: string;

  /**
   * Skip empty lines
   * @default true
   */
  skipEmptyLines?: boolean;

  /**
   * Trim whitespace from values
   * @default true
   */
  trimValues?: boolean;

  /**
   * First row contains headers (column names)
   * @default false
   */
  hasHeaders?: boolean;

  /**
   * Sheet name to create/use
   * @default 'Sheet1'
   */
  sheetName?: string;

  /**
   * Starting cell for import
   * @default 'A1'
   */
  startCell?: string;

  /**
   * Attempt to detect and parse numbers
   * @default true
   */
  detectNumbers?: boolean;

  /**
   * Attempt to detect and parse dates
   * @default false
   */
  detectDates?: boolean;

  /**
   * Date formats to try when detecting dates
   * @default ['yyyy-mm-dd', 'mm/dd/yyyy', 'dd/mm/yyyy']
   */
  dateFormats?: string[];

  /**
   * Maximum number of rows to import (0 = unlimited)
   * @default 0
   */
  maxRows?: number;

  /**
   * Comment character - lines starting with this are skipped
   */
  commentChar?: string;

  /**
   * Callback for progress reporting
   */
  onProgress?: (rowsProcessed: number, totalRows?: number) => void;
}

/**
 * Result of CSV import operation
 */
export interface CsvImportResult {
  /**
   * Number of rows imported
   */
  rowCount: number;

  /**
   * Number of columns imported
   */
  columnCount: number;

  /**
   * Headers if hasHeaders was true
   */
  headers?: string[];

  /**
   * Any warnings during import
   */
  warnings?: string[];
}

/**
 * Default export options
 */
export const DEFAULT_CSV_EXPORT_OPTIONS: Required<Omit<CsvExportOptions, 'range'>> = {
  delimiter: ',',
  rowDelimiter: '\r\n',
  quoteChar: '"',
  quoteAllFields: false,
  nullValue: '',
  dateFormat: 'ISO',
  includeBom: false,
  encoding: 'utf-8',
};

/**
 * Default import options
 */
export const DEFAULT_CSV_IMPORT_OPTIONS: Required<Omit<CsvImportOptions, 'commentChar' | 'onProgress' | 'dateFormats'>> & {
  dateFormats: string[];
} = {
  delimiter: ',',
  quoteChar: '"',
  skipEmptyLines: true,
  trimValues: true,
  hasHeaders: false,
  sheetName: 'Sheet1',
  startCell: 'A1',
  detectNumbers: true,
  detectDates: false,
  dateFormats: ['yyyy-mm-dd', 'mm/dd/yyyy', 'dd/mm/yyyy'],
  maxRows: 0,
};

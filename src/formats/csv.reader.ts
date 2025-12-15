/**
 * CSV Reader - Import CSV data into sheets
 */

import { Sheet } from '../core/Sheet.js';
import { Workbook } from '../core/Workbook.js';
import type { CellValue } from '../types/cell.types.js';
import type { CsvImportOptions, CsvImportResult } from './csv.types.js';
import { DEFAULT_CSV_IMPORT_OPTIONS } from './csv.types.js';
import { a1ToAddress } from '../types/cell.types.js';

/**
 * Parse CSV string and import into a new workbook
 *
 * @param csv - CSV string to parse
 * @param options - Import options
 * @returns Workbook containing the imported data
 *
 * @example
 * ```typescript
 * const csv = `Name,Age,City
 * Alice,30,New York
 * Bob,25,Los Angeles`;
 *
 * const workbook = csvToWorkbook(csv, { hasHeaders: true });
 * const sheet = workbook.getSheet('Sheet1');
 * console.log(sheet.cell('A1').value); // 'Name'
 * ```
 */
export function csvToWorkbook(csv: string, options: CsvImportOptions = {}): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addSheet(options.sheetName ?? DEFAULT_CSV_IMPORT_OPTIONS.sheetName);

  csvToSheet(csv, sheet, options);

  return workbook;
}

/**
 * Parse CSV string and import into an existing sheet
 *
 * @param csv - CSV string to parse
 * @param sheet - Sheet to import data into
 * @param options - Import options
 * @returns Import result with statistics
 */
export function csvToSheet(
  csv: string,
  sheet: Sheet,
  options: CsvImportOptions = {}
): CsvImportResult {
  const opts = { ...DEFAULT_CSV_IMPORT_OPTIONS, ...options };

  // Detect delimiter if not explicitly specified in options
  const delimiter = options.delimiter ?? detectDelimiter(csv);

  // Parse CSV into rows
  const rows = parseCsv(csv, delimiter, opts.quoteChar, opts.skipEmptyLines, opts.commentChar);

  // Apply max rows limit
  const limitedRows = opts.maxRows > 0 ? rows.slice(0, opts.maxRows) : rows;

  // Get starting position
  const startAddr = a1ToAddress(opts.startCell);
  let startRow = startAddr.row;
  const startCol = startAddr.col;

  const result: CsvImportResult = {
    rowCount: 0,
    columnCount: 0,
    warnings: [],
  };

  // Extract headers if present
  if (opts.hasHeaders && limitedRows.length > 0) {
    result.headers = limitedRows[0].map((v) => String(v));
  }

  // Import rows
  let maxCols = 0;

  for (let i = 0; i < limitedRows.length; i++) {
    const row = limitedRows[i];
    maxCols = Math.max(maxCols, row.length);

    for (let j = 0; j < row.length; j++) {
      let value: CellValue = row[j];

      // Trim values if requested
      if (opts.trimValues && typeof value === 'string') {
        value = value.trim();
      }

      // Detect and convert types
      if (typeof value === 'string') {
        value = convertValue(value, opts);
      }

      // Set cell value
      sheet.cell(startRow + i, startCol + j).value = value;
    }

    // Report progress
    if (opts.onProgress) {
      opts.onProgress(i + 1, limitedRows.length);
    }
  }

  result.rowCount = limitedRows.length;
  result.columnCount = maxCols;

  return result;
}

/**
 * Parse CSV from a Uint8Array buffer
 *
 * @param buffer - CSV data as Uint8Array
 * @param options - Import options
 * @returns Workbook containing the imported data
 */
export function csvBufferToWorkbook(buffer: Uint8Array, options: CsvImportOptions = {}): Workbook {
  const csv = new TextDecoder('utf-8').decode(buffer);
  // Remove BOM if present
  const cleanCsv = csv.charCodeAt(0) === 0xfeff ? csv.slice(1) : csv;
  return csvToWorkbook(cleanCsv, options);
}

/**
 * Detect the most likely delimiter in a CSV string
 */
function detectDelimiter(csv: string): string {
  const delimiters = [',', ';', '\t', '|'];
  const firstLine = csv.split(/\r?\n/)[0] || '';

  let bestDelimiter = ',';
  let bestCount = 0;

  for (const delim of delimiters) {
    // Count occurrences outside of quotes
    let count = 0;
    let inQuotes = false;

    for (const char of firstLine) {
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === delim && !inQuotes) {
        count++;
      }
    }

    if (count > bestCount) {
      bestCount = count;
      bestDelimiter = delim;
    }
  }

  return bestDelimiter;
}

/**
 * Parse CSV string into array of rows
 *
 * Implements RFC 4180 parsing with support for:
 * - Quoted fields
 * - Escaped quotes (doubled)
 * - Multi-line fields
 * - Custom delimiters
 */
function parseCsv(
  csv: string,
  delimiter: string,
  quoteChar: string,
  skipEmptyLines: boolean,
  commentChar?: string
): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;
  let i = 0;

  // Normalize line endings
  const normalized = csv.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  while (i < normalized.length) {
    const char = normalized[i];
    const nextChar = normalized[i + 1];

    if (inQuotes) {
      if (char === quoteChar) {
        if (nextChar === quoteChar) {
          // Escaped quote
          currentField += quoteChar;
          i += 2;
          continue;
        } else {
          // End of quoted field
          inQuotes = false;
          i++;
          continue;
        }
      } else {
        currentField += char;
        i++;
        continue;
      }
    }

    // Not in quotes
    if (char === quoteChar) {
      inQuotes = true;
      i++;
      continue;
    }

    if (char === delimiter) {
      currentRow.push(currentField);
      currentField = '';
      i++;
      continue;
    }

    if (char === '\n') {
      currentRow.push(currentField);

      // Check for comment or empty line
      const isComment = commentChar && currentRow[0]?.startsWith(commentChar);
      const isEmpty = skipEmptyLines && currentRow.length === 1 && currentRow[0] === '';

      if (!isComment && !isEmpty) {
        rows.push(currentRow);
      }

      currentRow = [];
      currentField = '';
      i++;
      continue;
    }

    currentField += char;
    i++;
  }

  // Handle last field/row
  if (currentField || currentRow.length > 0) {
    currentRow.push(currentField);

    const isComment = commentChar && currentRow[0]?.startsWith(commentChar);
    const isEmpty = skipEmptyLines && currentRow.length === 1 && currentRow[0] === '';

    if (!isComment && !isEmpty) {
      rows.push(currentRow);
    }
  }

  return rows;
}

/**
 * Convert a string value to appropriate type
 */
function convertValue(
  value: string,
  options: typeof DEFAULT_CSV_IMPORT_OPTIONS
): CellValue {
  // Empty string
  if (value === '') {
    return null;
  }

  // Boolean detection
  const lowerValue = value.toLowerCase();
  if (lowerValue === 'true') return true;
  if (lowerValue === 'false') return false;

  // Number detection
  if (options.detectNumbers) {
    const num = parseNumber(value);
    if (num !== null) {
      return num;
    }
  }

  // Date detection
  if (options.detectDates) {
    const date = parseDate(value, options.dateFormats);
    if (date !== null) {
      return date;
    }
  }

  return value;
}

/**
 * Try to parse a string as a number
 */
function parseNumber(value: string): number | null {
  // Remove thousands separators and handle different decimal formats
  const cleaned = value.trim();

  // Skip if empty or has letters (except e for scientific notation)
  if (!cleaned || /[a-df-zA-DF-Z]/.test(cleaned)) {
    return null;
  }

  // Handle percentage
  if (cleaned.endsWith('%')) {
    const num = parseFloat(cleaned.slice(0, -1));
    if (!isNaN(num)) {
      return num / 100;
    }
    return null;
  }

  // Handle currency symbols
  const withoutCurrency = cleaned.replace(/^[$€£¥₹]|[$€£¥₹]$/g, '');

  // Handle thousands separators (1,000 or 1.000)
  const normalized = withoutCurrency.replace(/,/g, '');

  const num = parseFloat(normalized);

  if (isNaN(num)) {
    return null;
  }

  // Verify it's actually a number representation
  // Avoid converting things like "123abc" which parseFloat accepts
  if (!/^-?\d*\.?\d+(?:[eE][+-]?\d+)?$/.test(normalized)) {
    return null;
  }

  return num;
}

/**
 * Try to parse a string as a date
 */
function parseDate(value: string, formats: string[]): Date | null {
  const trimmed = value.trim();

  // Try ISO format first
  const isoDate = new Date(trimmed);
  if (!isNaN(isoDate.getTime()) && trimmed.includes('-')) {
    return isoDate;
  }

  // Try each format
  for (const format of formats) {
    const date = parseDateFormat(trimmed, format);
    if (date) {
      return date;
    }
  }

  return null;
}

/**
 * Parse date with specific format
 */
function parseDateFormat(value: string, format: string): Date | null {
  // Simple format parsing
  const formatParts = format.toLowerCase().split(/[^a-z]+/);
  const valueParts = value.split(/[^0-9]+/);

  if (formatParts.length !== valueParts.length) {
    return null;
  }

  let year = 0;
  let month = 0;
  let day = 0;

  for (let i = 0; i < formatParts.length; i++) {
    const fp = formatParts[i];
    const vp = parseInt(valueParts[i], 10);

    if (isNaN(vp)) {
      return null;
    }

    if (fp === 'yyyy' || fp === 'yy') {
      year = fp === 'yy' ? (vp > 50 ? 1900 + vp : 2000 + vp) : vp;
    } else if (fp === 'mm' || fp === 'm') {
      month = vp - 1; // JS months are 0-indexed
    } else if (fp === 'dd' || fp === 'd') {
      day = vp;
    }
  }

  if (year && month >= 0 && day) {
    const date = new Date(year, month, day);
    // Validate the date is real
    if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
      return date;
    }
  }

  return null;
}

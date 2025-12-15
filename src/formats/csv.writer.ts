/**
 * CSV Writer - Export sheets to CSV format
 */

import type { Sheet } from '../core/Sheet.js';
import type { CellValue } from '../types/cell.types.js';
import type { CsvExportOptions } from './csv.types.js';
import { DEFAULT_CSV_EXPORT_OPTIONS } from './csv.types.js';
import { parseRangeReference } from '../types/range.types.js';

/**
 * UTF-8 BOM character sequence
 */
const UTF8_BOM = '\uFEFF';

/**
 * Export a sheet to CSV string
 *
 * @param sheet - The sheet to export
 * @param options - Export options
 * @returns CSV string
 *
 * @example
 * ```typescript
 * const csv = sheetToCsv(sheet);
 * console.log(csv);
 * // "Name","Age","City"
 * // "Alice",30,"New York"
 * // "Bob",25,"Los Angeles"
 * ```
 */
export function sheetToCsv(sheet: Sheet, options: CsvExportOptions = {}): string {
  const opts = { ...DEFAULT_CSV_EXPORT_OPTIONS, ...options };
  const { delimiter, rowDelimiter, quoteChar, quoteAllFields, includeBom } = opts;

  // Determine range to export
  const range = options.range
    ? parseRangeReference(options.range)
    : sheet.dimensions;

  if (!range) {
    // Empty sheet
    return includeBom ? UTF8_BOM : '';
  }

  const rows: string[] = [];

  // Build CSV rows
  for (let rowIdx = range.startRow; rowIdx <= range.endRow; rowIdx++) {
    const rowValues: string[] = [];

    for (let colIdx = range.startCol; colIdx <= range.endCol; colIdx++) {
      const cell = sheet.getCell(rowIdx, colIdx);
      const value = cell?.value ?? null;
      const formatted = formatValue(value, opts);
      const escaped = escapeField(formatted, delimiter, quoteChar, quoteAllFields);
      rowValues.push(escaped);
    }

    rows.push(rowValues.join(delimiter));
  }

  const csv = rows.join(rowDelimiter);
  return includeBom ? UTF8_BOM + csv : csv;
}

/**
 * Export a sheet to CSV and return as Uint8Array (for file writing)
 *
 * @param sheet - The sheet to export
 * @param options - Export options
 * @returns CSV as Uint8Array
 */
export function sheetToCsvBuffer(sheet: Sheet, options: CsvExportOptions = {}): Uint8Array {
  const csv = sheetToCsv(sheet, options);
  return new TextEncoder().encode(csv);
}

/**
 * Export multiple sheets to separate CSV strings
 *
 * @param sheets - Array of sheets to export
 * @param options - Export options (applied to all sheets)
 * @returns Map of sheet name to CSV string
 */
export function sheetsToCsv(
  sheets: Sheet[],
  options: CsvExportOptions = {}
): Map<string, string> {
  const result = new Map<string, string>();

  for (const sheet of sheets) {
    result.set(sheet.name, sheetToCsv(sheet, options));
  }

  return result;
}

/**
 * Format a cell value for CSV output
 */
function formatValue(value: CellValue, options: typeof DEFAULT_CSV_EXPORT_OPTIONS): string {
  if (value === null || value === undefined) {
    return options.nullValue;
  }

  if (typeof value === 'string') {
    return value;
  }

  if (typeof value === 'number') {
    // Handle special numbers
    if (Number.isNaN(value)) {
      return 'NaN';
    }
    if (!Number.isFinite(value)) {
      return value > 0 ? 'Infinity' : '-Infinity';
    }
    return String(value);
  }

  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }

  if (value instanceof Date) {
    return formatDate(value, options.dateFormat);
  }

  // Rich text - extract plain text
  if (typeof value === 'object' && 'richText' in value) {
    return value.richText.map((run) => run.text).join('');
  }

  return String(value);
}

/**
 * Format a date value
 */
function formatDate(date: Date, format: string): string {
  if (format === 'ISO') {
    return date.toISOString().split('T')[0]; // yyyy-mm-dd
  }

  if (format === 'locale') {
    return date.toLocaleDateString();
  }

  // Custom format
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');

  return format
    .replace('yyyy', String(year))
    .replace('yy', String(year).slice(-2))
    .replace('mm', month)
    .replace('dd', day)
    .replace('HH', hours)
    .replace('MM', minutes)
    .replace('SS', seconds);
}

/**
 * Escape a field value for CSV
 *
 * Follows RFC 4180:
 * - Fields containing delimiter, quote, or newline must be quoted
 * - Quote characters within quoted fields must be doubled
 */
function escapeField(
  value: string,
  delimiter: string,
  quoteChar: string,
  quoteAll: boolean
): string {
  // Check if quoting is needed
  const needsQuoting =
    quoteAll ||
    value.includes(delimiter) ||
    value.includes(quoteChar) ||
    value.includes('\n') ||
    value.includes('\r');

  if (!needsQuoting) {
    return value;
  }

  // Escape quote characters by doubling them
  const escaped = value.replace(new RegExp(escapeRegex(quoteChar), 'g'), quoteChar + quoteChar);

  return quoteChar + escaped + quoteChar;
}

/**
 * Escape special regex characters
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

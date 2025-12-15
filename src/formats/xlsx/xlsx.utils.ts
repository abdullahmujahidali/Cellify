/**
 * Utility functions for XLSX export
 */

/**
 * Excel epoch (December 31, 1899)
 * Excel's date system starts at January 1, 1900 = day 1
 * So Dec 31, 1899 = day 0 (epoch)
 */
const EXCEL_EPOCH = Date.UTC(1899, 11, 31); // Dec 31, 1899 UTC

/**
 * Milliseconds per day
 */
const MS_PER_DAY = 86400000;

/**
 * Excel's 1900 leap year bug cutoff
 * Excel incorrectly treats 1900 as a leap year (Feb 29, 1900 = day 60)
 * This bug was inherited from Lotus 1-2-3 for compatibility
 * Any date on or after March 1, 1900 (day 61) needs adjustment
 */
const LEAP_YEAR_BUG_CUTOFF = 60;

/**
 * Convert JavaScript Date to Excel serial date number
 *
 * Excel stores dates as floating-point numbers where:
 * - Integer part = days since December 30, 1899
 * - Fractional part = time of day
 *
 * @param date - JavaScript Date object
 * @returns Excel serial date number
 *
 * @example
 * dateToExcelSerial(new Date('2024-01-01')) // Returns 45292
 * dateToExcelSerial(new Date('1900-03-01')) // Returns 61 (with leap year bug adjustment)
 */
export function dateToExcelSerial(date: Date): number {
  // Use getTime() directly - it's already UTC milliseconds
  const utcTime = date.getTime();

  // Calculate days since Excel epoch
  let serial = (utcTime - EXCEL_EPOCH) / MS_PER_DAY;

  // Account for Excel's 1900 leap year bug
  // Dates on or after March 1, 1900 need to be incremented by 1
  if (serial >= LEAP_YEAR_BUG_CUTOFF) {
    serial += 1;
  }

  return serial;
}

/**
 * Convert Excel serial date to JavaScript Date
 *
 * @param serial - Excel serial date number
 * @returns JavaScript Date object
 */
export function excelSerialToDate(serial: number): Date {
  // Adjust for leap year bug
  if (serial >= LEAP_YEAR_BUG_CUTOFF) {
    serial -= 1;
  }

  const ms = EXCEL_EPOCH + serial * MS_PER_DAY;
  return new Date(ms);
}

/**
 * Convert Cellify column width (in characters) to Excel column width
 *
 * Excel column width formula:
 * width = Truncate(({NumChars} * {MaxDigitWidth} + {5 pixel padding}) / {MaxDigitWidth} * 256) / 256
 *
 * For Calibri 11pt (default Excel font), MaxDigitWidth ≈ 7 pixels
 *
 * @param charWidth - Width in character units
 * @returns Excel column width value
 */
export function toExcelColumnWidth(charWidth: number): number {
  const MAX_DIGIT_WIDTH = 7;
  return Math.round(((charWidth * MAX_DIGIT_WIDTH + 5) / MAX_DIGIT_WIDTH) * 256) / 256;
}

/**
 * Convert Excel column width to character width
 *
 * @param excelWidth - Excel column width value
 * @returns Width in character units
 */
export function fromExcelColumnWidth(excelWidth: number): number {
  const MAX_DIGIT_WIDTH = 7;
  return (excelWidth * MAX_DIGIT_WIDTH - 5) / MAX_DIGIT_WIDTH;
}

/**
 * Convert column index (0-based) to Excel column letter
 *
 * @param index - 0-based column index
 * @returns Excel column letter (A, B, ..., Z, AA, AB, ...)
 *
 * @example
 * columnIndexToLetter(0)  // 'A'
 * columnIndexToLetter(25) // 'Z'
 * columnIndexToLetter(26) // 'AA'
 */
export function columnIndexToLetter(index: number): string {
  let result = '';
  let n = index;

  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }

  return result;
}

/**
 * Convert row and column indices to A1-style cell reference
 *
 * @param row - 0-based row index
 * @param col - 0-based column index
 * @returns A1-style reference (e.g., "A1", "B2", "AA100")
 */
export function cellRef(row: number, col: number): string {
  return columnIndexToLetter(col) + (row + 1);
}

/**
 * Convert a range to A1-style reference
 *
 * @param startRow - 0-based start row
 * @param startCol - 0-based start column
 * @param endRow - 0-based end row
 * @param endCol - 0-based end column
 * @returns A1-style range reference (e.g., "A1:C10")
 */
export function rangeRef(startRow: number, startCol: number, endRow: number, endCol: number): string {
  return `${cellRef(startRow, startCol)}:${cellRef(endRow, endCol)}`;
}

/**
 * Get the span string for a row (e.g., "1:5" for columns A-E)
 *
 * @param minCol - 0-based minimum column index
 * @param maxCol - 0-based maximum column index
 * @returns Excel span string (1-based)
 */
export function getRowSpans(minCol: number, maxCol: number): string {
  return `${minCol + 1}:${maxCol + 1}`;
}

/**
 * Determine if a cell value should use a date number format
 *
 * @param value - Cell value
 * @returns True if value is a Date
 */
export function isDateValue(value: unknown): value is Date {
  return value instanceof Date;
}

/**
 * Get rich text as plain string
 *
 * @param richText - Rich text value object
 * @returns Plain text string
 */
export function richTextToString(richText: { richText: Array<{ text: string }> }): string {
  return richText.richText.map((run) => run.text).join('');
}

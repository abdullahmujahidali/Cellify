/**
 * XLSX Format Module
 *
 * Export workbooks to Excel (.xlsx) format with full styling support.
 */

// Types
export type { XlsxExportOptions } from './xlsx.types.js';
export { DEFAULT_XLSX_OPTIONS } from './xlsx.types.js';

// Writer
export { workbookToXlsx, workbookToXlsxBlob } from './xlsx.writer.js';

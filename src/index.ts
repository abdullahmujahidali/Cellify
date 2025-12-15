/**
 * Cellify - A lightweight, affordable spreadsheet library
 *
 * @example
 * ```typescript
 * import { Workbook } from 'cellify';
 *
 * const workbook = new Workbook();
 * const sheet = workbook.addSheet('Data');
 *
 * // Set cell values
 * sheet.cell('A1').value = 'Hello';
 * sheet.cell('B1').value = 'World';
 *
 * // Apply styles
 * sheet.cell('A1').applyStyle({
 *   font: { bold: true, size: 14 },
 *   fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FFFF00' }
 * });
 *
 * // Merge cells
 * sheet.mergeCells('A1:B1');
 * ```
 */

// Core classes
export { Cell, Sheet, Workbook } from './core/index.js';
export type {
  RowConfig,
  ColumnConfig,
  SheetView,
  PageSetup,
  SheetProtection,
  WorkbookProperties,
  DefinedName,
  CalculationMode,
  WorkbookView,
} from './core/index.js';

// All types
export * from './types/index.js';

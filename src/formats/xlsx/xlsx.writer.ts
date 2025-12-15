/**
 * XLSX Writer - Export workbooks to Excel format
 *
 * Generates valid OOXML (.xlsx) files that can be opened in:
 * - Microsoft Excel
 * - LibreOffice Calc
 * - Google Sheets
 * - Apple Numbers
 */

import { zipSync, strToU8 } from 'fflate';
import type { Zippable } from 'fflate';
import type { Workbook } from '../../core/Workbook.js';
import type { XlsxExportOptions, XlsxBuildContext } from './xlsx.types.js';
import { DEFAULT_XLSX_OPTIONS } from './xlsx.types.js';
import { SharedStringsTable } from './xlsx.strings.js';
import { StyleRegistry } from './xlsx.styles.js';
import {
  generateContentTypes,
  generateRootRels,
  generateWorkbook,
  generateWorkbookRels,
  generateWorksheet,
  generateCoreProperties,
  generateAppProperties,
} from './xlsx.parts.js';

/**
 * Export a workbook to XLSX format
 *
 * @param workbook - The workbook to export
 * @param options - Export options
 * @returns Uint8Array containing the XLSX file data
 *
 * @example
 * ```typescript
 * import { Workbook, workbookToXlsx } from 'cellify';
 *
 * const workbook = new Workbook();
 * const sheet = workbook.addSheet('Data');
 * sheet.cell('A1').value = 'Hello';
 * sheet.cell('B1').value = 42;
 *
 * const xlsxData = workbookToXlsx(workbook);
 * // Write to file or send to client
 * ```
 */
export function workbookToXlsx(workbook: Workbook, options: XlsxExportOptions = {}): Uint8Array {
  const opts = { ...DEFAULT_XLSX_OPTIONS, ...options };

  // Ensure workbook has at least one sheet
  if (workbook.sheetCount === 0) {
    workbook.addSheet('Sheet1');
  }

  // Initialize build context
  const ctx: XlsxBuildContext = {
    workbook,
    options: opts,
    sharedStrings: new SharedStringsTable(),
    styleRegistry: new StyleRegistry(),
    sheets: [],
    stylesRId: '',
  };

  // Assign relationship IDs
  assignRelationshipIds(ctx);

  // First pass: collect all strings and register all styles
  // This needs to happen before generating XML so all indices are known
  collectStringsAndStyles(ctx);

  // Build ZIP structure
  const files: Zippable = {};

  // Required parts
  files['[Content_Types].xml'] = strToU8(generateContentTypes(ctx));
  files['_rels/.rels'] = strToU8(generateRootRels(ctx));
  files['xl/workbook.xml'] = strToU8(generateWorkbook(ctx));
  files['xl/_rels/workbook.xml.rels'] = strToU8(generateWorkbookRels(ctx));
  files['xl/styles.xml'] = strToU8(ctx.styleRegistry.generateStylesXml());

  // Shared strings (if any)
  if (ctx.sharedStrings.count > 0) {
    files['xl/sharedStrings.xml'] = strToU8(ctx.sharedStrings.generateXml());
  }

  // Worksheets
  const sheets = workbook.sheets;
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(generateWorksheet(sheet, ctx));
  }

  // Document properties (optional)
  if (opts.includeProperties) {
    files['docProps/core.xml'] = strToU8(generateCoreProperties(ctx));
    files['docProps/app.xml'] = strToU8(generateAppProperties(ctx));
  }

  // Create ZIP archive
  return zipSync(files, { level: opts.compressionLevel });
}

/**
 * Export workbook to XLSX and return as Blob
 * Useful for browser downloads
 *
 * @param workbook - The workbook to export
 * @param options - Export options
 * @returns Blob containing the XLSX file
 *
 * @example
 * ```typescript
 * const blob = workbookToXlsxBlob(workbook);
 * const url = URL.createObjectURL(blob);
 *
 * const a = document.createElement('a');
 * a.href = url;
 * a.download = 'spreadsheet.xlsx';
 * a.click();
 * ```
 */
export function workbookToXlsxBlob(workbook: Workbook, options?: XlsxExportOptions): Blob {
  const data = workbookToXlsx(workbook, options);
  // Create a new ArrayBuffer from the Uint8Array for Blob compatibility
  const buffer = new ArrayBuffer(data.length);
  new Uint8Array(buffer).set(data);
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Assign relationship IDs for all parts
 */
function assignRelationshipIds(ctx: XlsxBuildContext): void {
  let rId = 1;

  // Assign sheet rIds first
  const sheets = ctx.workbook.sheets;
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    ctx.sheets.push({
      name: sheet.name,
      sheetId: i + 1,
      rId: `rId${rId++}`,
      target: `worksheets/sheet${i + 1}.xml`,
    });
  }

  // Styles
  ctx.stylesRId = `rId${rId++}`;

  // Shared strings (if enabled)
  if (ctx.options.useSharedStrings) {
    ctx.sharedStringsRId = `rId${rId++}`;
  }
}

/**
 * First pass: collect all strings and register all styles
 *
 * This ensures all shared string indices and style indices are assigned
 * before we generate the actual XML content.
 */
function collectStringsAndStyles(ctx: XlsxBuildContext): void {
  const sheets = ctx.workbook.sheets;

  for (const sheet of sheets) {
    for (const cell of sheet.cells()) {
      // Skip merged slave cells
      if (cell.isMergedSlave) continue;

      // Register style
      if (cell.style) {
        ctx.styleRegistry.registerStyle(cell.style);
      }

      // Collect strings
      if (ctx.options.useSharedStrings) {
        const value = cell.value;

        if (typeof value === 'string') {
          ctx.sharedStrings.add(value);
        } else if (value && typeof value === 'object' && 'richText' in value) {
          // Rich text - convert to plain string for shared strings
          const plainText = value.richText.map((run: { text: string }) => run.text).join('');
          ctx.sharedStrings.add(plainText);
        }

        // Also collect formula results if they're strings
        if (cell.formula?.result && typeof cell.formula.result === 'string') {
          ctx.sharedStrings.add(cell.formula.result);
        }
      }
    }
  }
}

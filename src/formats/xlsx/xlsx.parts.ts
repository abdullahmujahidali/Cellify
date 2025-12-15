/**
 * XML Part Generators for XLSX export
 *
 * Generates the individual XML files that make up an XLSX package.
 */

import type { Sheet } from '../../core/Sheet.js';
import type { Cell } from '../../core/Cell.js';
import type { CellValue } from '../../types/cell.types.js';
import type { XlsxBuildContext } from './xlsx.types.js';
import { NS, REL_TYPES, CONTENT_TYPES } from './xlsx.types.js';
import { XML_DECLARATION, escapeXml, sanitizeXmlString, isExcelError } from './xlsx.xml.js';
import { dateToExcelSerial, cellRef, rangeRef, toExcelColumnWidth, richTextToString } from './xlsx.utils.js';

/**
 * Generate [Content_Types].xml
 */
export function generateContentTypes(ctx: XlsxBuildContext): string {
  const parts: string[] = [XML_DECLARATION];

  parts.push(`<Types xmlns="${NS.contentTypes}">`);

  // Default extensions
  parts.push('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
  parts.push('<Default Extension="xml" ContentType="application/xml"/>');

  // Workbook
  parts.push(`<Override PartName="/xl/workbook.xml" ContentType="${CONTENT_TYPES.workbook}"/>`);

  // Worksheets
  for (let i = 0; i < ctx.sheets.length; i++) {
    parts.push(`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CONTENT_TYPES.worksheet}"/>`);
  }

  // Styles
  parts.push(`<Override PartName="/xl/styles.xml" ContentType="${CONTENT_TYPES.styles}"/>`);

  // Shared strings (if used)
  if (ctx.sharedStrings.count > 0) {
    parts.push(`<Override PartName="/xl/sharedStrings.xml" ContentType="${CONTENT_TYPES.sharedStrings}"/>`);
  }

  // Document properties
  if (ctx.options.includeProperties) {
    parts.push(`<Override PartName="/docProps/core.xml" ContentType="${CONTENT_TYPES.coreProperties}"/>`);
    parts.push(`<Override PartName="/docProps/app.xml" ContentType="${CONTENT_TYPES.extendedProperties}"/>`);
  }

  parts.push('</Types>');

  return parts.join('\n');
}

/**
 * Generate _rels/.rels (package relationships)
 */
export function generateRootRels(ctx: XlsxBuildContext): string {
  const parts: string[] = [XML_DECLARATION];

  parts.push(`<Relationships xmlns="${NS.packageRels}">`);

  parts.push(`<Relationship Id="rId1" Type="${REL_TYPES.officeDocument}" Target="xl/workbook.xml"/>`);

  if (ctx.options.includeProperties) {
    parts.push(`<Relationship Id="rId2" Type="${REL_TYPES.coreProperties}" Target="docProps/core.xml"/>`);
    parts.push(`<Relationship Id="rId3" Type="${REL_TYPES.extendedProperties}" Target="docProps/app.xml"/>`);
  }

  parts.push('</Relationships>');

  return parts.join('\n');
}

/**
 * Generate xl/workbook.xml
 */
export function generateWorkbook(ctx: XlsxBuildContext): string {
  const parts: string[] = [XML_DECLARATION];

  parts.push(`<workbook xmlns="${NS.spreadsheetml}" xmlns:r="${NS.r}">`);

  // Sheets
  parts.push('<sheets>');
  for (const sheet of ctx.sheets) {
    parts.push(`<sheet name="${escapeXml(sheet.name)}" sheetId="${sheet.sheetId}" r:id="${sheet.rId}"/>`);
  }
  parts.push('</sheets>');

  parts.push('</workbook>');

  return parts.join('\n');
}

/**
 * Generate xl/_rels/workbook.xml.rels
 */
export function generateWorkbookRels(ctx: XlsxBuildContext): string {
  const parts: string[] = [XML_DECLARATION];

  parts.push(`<Relationships xmlns="${NS.packageRels}">`);

  // Worksheet relationships
  for (const sheet of ctx.sheets) {
    parts.push(`<Relationship Id="${sheet.rId}" Type="${REL_TYPES.worksheet}" Target="${sheet.target}"/>`);
  }

  // Styles
  parts.push(`<Relationship Id="${ctx.stylesRId}" Type="${REL_TYPES.styles}" Target="styles.xml"/>`);

  // Shared strings
  if (ctx.sharedStringsRId) {
    parts.push(`<Relationship Id="${ctx.sharedStringsRId}" Type="${REL_TYPES.sharedStrings}" Target="sharedStrings.xml"/>`);
  }

  parts.push('</Relationships>');

  return parts.join('\n');
}

/**
 * Generate xl/worksheets/sheetN.xml
 */
export function generateWorksheet(sheet: Sheet, ctx: XlsxBuildContext): string {
  const parts: string[] = [XML_DECLARATION];

  parts.push(`<worksheet xmlns="${NS.spreadsheetml}" xmlns:r="${NS.r}">`);

  // Sheet views (for freeze panes, grid lines, etc.)
  parts.push(generateSheetViews(sheet));

  // Sheet format properties
  parts.push(generateSheetFormatPr(ctx));

  // Column definitions
  const colsXml = generateCols(sheet, ctx);
  if (colsXml) {
    parts.push(colsXml);
  }

  // Sheet data (rows and cells)
  parts.push(generateSheetData(sheet, ctx));

  // Merge cells
  const mergesXml = generateMergeCells(sheet);
  if (mergesXml) {
    parts.push(mergesXml);
  }

  // Auto filter
  const autoFilterXml = generateAutoFilter(sheet);
  if (autoFilterXml) {
    parts.push(autoFilterXml);
  }

  parts.push('</worksheet>');

  return parts.join('\n');
}

/**
 * Generate sheetViews element
 */
function generateSheetViews(sheet: Sheet): string {
  const view = sheet.view;
  const parts: string[] = ['<sheetViews>'];

  let sheetViewAttrs = 'tabSelected="1" workbookViewId="0"';

  if (view.showGridLines === false) {
    sheetViewAttrs += ' showGridLines="0"';
  }
  if (view.showRowColHeaders === false) {
    sheetViewAttrs += ' showRowColHeaders="0"';
  }
  if (view.zoomScale && view.zoomScale !== 100) {
    sheetViewAttrs += ` zoomScale="${view.zoomScale}"`;
  }

  parts.push(`<sheetView ${sheetViewAttrs}>`);

  // Freeze panes
  if (view.frozenRows || view.frozenCols) {
    const frozenRows = view.frozenRows ?? 0;
    const frozenCols = view.frozenCols ?? 0;

    // Determine active pane
    let activePane: string;
    if (frozenRows > 0 && frozenCols > 0) {
      activePane = 'bottomRight';
    } else if (frozenRows > 0) {
      activePane = 'bottomLeft';
    } else {
      activePane = 'topRight';
    }

    const topLeftCell = cellRef(frozenRows, frozenCols);

    parts.push(
      `<pane xSplit="${frozenCols}" ySplit="${frozenRows}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`
    );
    parts.push(`<selection pane="${activePane}" activeCell="${topLeftCell}" sqref="${topLeftCell}"/>`);
  }

  parts.push('</sheetView>');
  parts.push('</sheetViews>');

  return parts.join('\n');
}

/**
 * Generate sheetFormatPr element
 */
function generateSheetFormatPr(ctx: XlsxBuildContext): string {
  return `<sheetFormatPr defaultRowHeight="${ctx.options.defaultRowHeight}" defaultColWidth="${ctx.options.defaultColumnWidth}"/>`;
}

/**
 * Generate cols element for column widths
 */
function generateCols(sheet: Sheet, ctx: XlsxBuildContext): string | null {
  const columns = sheet.columns;
  if (columns.size === 0) {
    return null;
  }

  const parts: string[] = ['<cols>'];

  // Sort column indices
  const sortedCols = [...columns.keys()].sort((a, b) => a - b);

  for (const colIdx of sortedCols) {
    const config = columns.get(colIdx)!;
    const colNum = colIdx + 1; // Excel uses 1-based

    const attrs: string[] = [`min="${colNum}"`, `max="${colNum}"`];

    if (config.width !== undefined) {
      attrs.push(`width="${toExcelColumnWidth(config.width)}"`, 'customWidth="1"');
    } else {
      attrs.push(`width="${ctx.options.defaultColumnWidth}"`);
    }

    if (config.hidden) {
      attrs.push('hidden="1"');
    }

    parts.push(`<col ${attrs.join(' ')}/>`);
  }

  parts.push('</cols>');

  return parts.join('\n');
}

/**
 * Generate sheetData element with rows and cells
 */
function generateSheetData(sheet: Sheet, ctx: XlsxBuildContext): string {
  const dims = sheet.dimensions;
  if (!dims) {
    return '<sheetData/>';
  }

  const parts: string[] = ['<sheetData>'];

  // Group cells by row
  const rowCells = new Map<number, Cell[]>();

  for (const cell of sheet.cells()) {
    // Skip merged slave cells
    if (cell.isMergedSlave) continue;

    let row = rowCells.get(cell.row);
    if (!row) {
      row = [];
      rowCells.set(cell.row, row);
    }
    row.push(cell);
  }

  // Generate rows in order
  const sortedRows = [...rowCells.keys()].sort((a, b) => a - b);

  for (const rowIdx of sortedRows) {
    const cells = rowCells.get(rowIdx)!;
    cells.sort((a, b) => a.col - b.col);

    // Row attributes
    const rowConfig = sheet.getRow(rowIdx);
    const rowNum = rowIdx + 1; // Excel uses 1-based

    const rowAttrs: string[] = [`r="${rowNum}"`];

    if (rowConfig.height !== undefined) {
      rowAttrs.push(`ht="${rowConfig.height}"`, 'customHeight="1"');
    }
    if (rowConfig.hidden) {
      rowAttrs.push('hidden="1"');
    }

    parts.push(`<row ${rowAttrs.join(' ')}>`);

    // Generate cells
    for (const cell of cells) {
      parts.push(generateCell(cell, ctx));
    }

    parts.push('</row>');
  }

  parts.push('</sheetData>');

  return parts.join('\n');
}

/**
 * Generate a single cell element
 */
function generateCell(cell: Cell, ctx: XlsxBuildContext): string {
  const ref = cell.address;
  const style = cell.style;
  const formula = cell.formula;
  const value = cell.value;

  // Get style index
  const styleIdx = ctx.styleRegistry.registerStyle(style);

  // Build cell attributes
  const attrs: string[] = [`r="${ref}"`];
  if (styleIdx > 0) {
    attrs.push(`s="${styleIdx}"`);
  }

  // Formula cell
  if (formula) {
    const formulaXml = `<f>${escapeXml(formula.formula)}</f>`;

    if (formula.result !== undefined) {
      const resultXml = formatCellValue(formula.result, ctx);
      if (resultXml.type) {
        attrs.push(`t="${resultXml.type}"`);
      }
      return `<c ${attrs.join(' ')}>${formulaXml}${resultXml.xml}</c>`;
    }

    return `<c ${attrs.join(' ')}>${formulaXml}</c>`;
  }

  // Empty cell with style
  if (value === null || value === undefined) {
    return `<c ${attrs.join(' ')}/>`;
  }

  // Value cell
  const valueResult = formatCellValue(value, ctx);
  if (valueResult.type) {
    attrs.push(`t="${valueResult.type}"`);
  }

  return `<c ${attrs.join(' ')}>${valueResult.xml}</c>`;
}

/**
 * Format a cell value for XML output
 */
function formatCellValue(
  value: CellValue,
  ctx: XlsxBuildContext
): { type?: string; xml: string } {
  if (value === null || value === undefined) {
    return { xml: '' };
  }

  // Number
  if (typeof value === 'number') {
    if (Number.isNaN(value)) {
      return { type: 'e', xml: '<v>#NUM!</v>' };
    }
    if (!Number.isFinite(value)) {
      return { type: 'e', xml: '<v>#NUM!</v>' };
    }
    return { xml: `<v>${value}</v>` };
  }

  // Boolean
  if (typeof value === 'boolean') {
    return { type: 'b', xml: `<v>${value ? 1 : 0}</v>` };
  }

  // Date
  if (value instanceof Date) {
    const serial = dateToExcelSerial(value);
    return { xml: `<v>${serial}</v>` };
  }

  // String
  if (typeof value === 'string') {
    // Check for Excel error values
    if (isExcelError(value)) {
      return { type: 'e', xml: `<v>${escapeXml(value)}</v>` };
    }

    // Use shared strings if enabled
    if (ctx.options.useSharedStrings) {
      const idx = ctx.sharedStrings.add(value);
      return { type: 's', xml: `<v>${idx}</v>` };
    }

    // Inline string
    return { type: 'inlineStr', xml: `<is><t>${sanitizeXmlString(value)}</t></is>` };
  }

  // Rich text
  if (typeof value === 'object' && 'richText' in value) {
    const plainText = richTextToString(value);
    if (ctx.options.useSharedStrings) {
      const idx = ctx.sharedStrings.add(plainText);
      return { type: 's', xml: `<v>${idx}</v>` };
    }
    return { type: 'inlineStr', xml: `<is><t>${sanitizeXmlString(plainText)}</t></is>` };
  }

  return { xml: '' };
}

/**
 * Generate mergeCells element
 */
function generateMergeCells(sheet: Sheet): string | null {
  const merges = sheet.merges;
  if (merges.length === 0) {
    return null;
  }

  const parts: string[] = [`<mergeCells count="${merges.length}">`];

  for (const merge of merges) {
    const ref = rangeRef(merge.startRow, merge.startCol, merge.endRow, merge.endCol);
    parts.push(`<mergeCell ref="${ref}"/>`);
  }

  parts.push('</mergeCells>');

  return parts.join('\n');
}

/**
 * Generate autoFilter element
 */
function generateAutoFilter(sheet: Sheet): string | null {
  const autoFilter = sheet.autoFilter;
  if (!autoFilter) {
    return null;
  }

  const ref = rangeRef(
    autoFilter.range.startRow,
    autoFilter.range.startCol,
    autoFilter.range.endRow,
    autoFilter.range.endCol
  );

  return `<autoFilter ref="${ref}"/>`;
}

/**
 * Generate docProps/core.xml
 */
export function generateCoreProperties(ctx: XlsxBuildContext): string {
  const wb = ctx.workbook;
  const parts: string[] = [XML_DECLARATION];

  parts.push(
    `<cp:coreProperties xmlns:cp="${NS.coreProperties}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}">`
  );

  if (wb.properties.title) {
    parts.push(`<dc:title>${escapeXml(wb.properties.title)}</dc:title>`);
  }
  if (wb.properties.author) {
    parts.push(`<dc:creator>${escapeXml(wb.properties.author)}</dc:creator>`);
  }
  if (wb.properties.lastModifiedBy) {
    parts.push(`<cp:lastModifiedBy>${escapeXml(wb.properties.lastModifiedBy)}</cp:lastModifiedBy>`);
  }
  if (wb.properties.created) {
    parts.push(`<dcterms:created xsi:type="dcterms:W3CDTF">${wb.properties.created.toISOString()}</dcterms:created>`);
  }
  if (wb.properties.modified) {
    parts.push(`<dcterms:modified xsi:type="dcterms:W3CDTF">${wb.properties.modified.toISOString()}</dcterms:modified>`);
  }

  parts.push('</cp:coreProperties>');

  return parts.join('\n');
}

/**
 * Generate docProps/app.xml
 */
export function generateAppProperties(ctx: XlsxBuildContext): string {
  const parts: string[] = [XML_DECLARATION];

  parts.push(`<Properties xmlns="${NS.extendedProperties}">`);
  parts.push(`<Application>${escapeXml(ctx.options.application)}</Application>`);
  parts.push('</Properties>');

  return parts.join('\n');
}

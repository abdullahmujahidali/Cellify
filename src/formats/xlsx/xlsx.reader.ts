/**
 * XLSX Reader - Import Excel files into Workbooks
 *
 * Parses OOXML (.xlsx) files and creates Cellify Workbook objects.
 */

import { unzipSync } from 'fflate';
import { Workbook } from '../../core/Workbook.js';
import type { Sheet } from '../../core/Sheet.js';
import type { CellStyle, Border, CellBorders, BorderStyle as BorderStyleType } from '../../types/style.types.js';
import type {
  XlsxImportOptions,
  XlsxImportResult,
  XlsxParseContext,
  SheetInfo,
  CellXfInfo,
  FontInfo,
  FillInfo,
  BorderInfo,
  BorderSideInfo,
} from './xlsx.reader.types.js';
import { DEFAULT_XLSX_IMPORT_OPTIONS, DATE_FORMAT_IDS } from './xlsx.reader.types.js';
import {
  parseElement,
  parseElements,
  getAttr,
  getTextContent,
  parseCellRef,
  resolveRelPath,
  unescapeXml,
} from './xlsx.parser.js';
import { excelSerialToDate } from './xlsx.utils.js';
import { a1ToAddress } from '../../types/cell.types.js';

/**
 * Import an XLSX file into a Workbook
 *
 * @param buffer - XLSX file data as Uint8Array
 * @param options - Import options
 * @returns Import result with workbook, stats, and warnings
 *
 * @example
 * ```typescript
 * import { xlsxToWorkbook } from 'cellify';
 *
 * const response = await fetch('spreadsheet.xlsx');
 * const buffer = new Uint8Array(await response.arrayBuffer());
 *
 * const { workbook, stats, warnings } = xlsxToWorkbook(buffer);
 * console.log(`Imported ${stats.sheetCount} sheets, ${stats.totalCells} cells`);
 * ```
 */
export function xlsxToWorkbook(
  buffer: Uint8Array,
  options: XlsxImportOptions = {}
): XlsxImportResult {
  const startTime = performance.now();

  const opts = { ...DEFAULT_XLSX_IMPORT_OPTIONS, ...options };

  // Initialize parse context
  const ctx: XlsxParseContext = {
    options: opts,
    onProgress: options.onProgress,
    sharedStrings: [],
    numberFormats: new Map(),
    cellXfs: [],
    fonts: [],
    fills: [],
    borders: [],
    sheetInfos: [],
    warnings: [],
    stats: {
      sheetCount: 0,
      totalCells: 0,
      formulaCells: 0,
      mergedRanges: 0,
      durationMs: 0,
    },
  };

  // Unzip the XLSX file
  ctx.onProgress?.('unzip', 0, 1);
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(buffer);
  } catch (error) {
    throw new Error(`Failed to unzip XLSX file: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
  ctx.onProgress?.('unzip', 1, 1);

  // Helper to read file as string
  const readFile = (path: string): string | undefined => {
    const data = files[path];
    if (!data) return undefined;
    return new TextDecoder('utf-8').decode(data);
  };

  // Parse workbook relationships to find sheet targets
  const workbookRelsXml = readFile('xl/_rels/workbook.xml.rels');
  const sheetTargets = new Map<string, string>();
  if (workbookRelsXml) {
    const rels = parseElements(workbookRelsXml, 'Relationship');
    for (const rel of rels) {
      const rId = getAttr(rel, 'Id');
      const target = getAttr(rel, 'Target');
      if (rId && target) {
        sheetTargets.set(rId, target);
      }
    }
  }

  // Parse workbook.xml to get sheet list
  const workbookXml = readFile('xl/workbook.xml');
  if (!workbookXml) {
    throw new Error('Invalid XLSX: missing xl/workbook.xml');
  }

  const sheetElements = parseElements(workbookXml, 'sheet');
  for (const sheetEl of sheetElements) {
    const name = getAttr(sheetEl, 'name') || 'Sheet';
    const sheetId = parseInt(getAttr(sheetEl, 'sheetId') || '0', 10);
    const rId = getAttr(sheetEl, 'r:id') || getAttr(sheetEl, 'rId') || '';

    ctx.sheetInfos.push({ name, sheetId, rId });
  }

  // Parse shared strings
  ctx.onProgress?.('sharedStrings', 0, 1);
  const sharedStringsXml = readFile('xl/sharedStrings.xml');
  if (sharedStringsXml) {
    parseSharedStrings(sharedStringsXml, ctx);
  }
  ctx.onProgress?.('sharedStrings', 1, 1);

  // Parse styles
  ctx.onProgress?.('styles', 0, 1);
  const stylesXml = readFile('xl/styles.xml');
  if (stylesXml && opts.importStyles) {
    parseStyles(stylesXml, ctx);
  }
  ctx.onProgress?.('styles', 1, 1);

  // Create workbook
  const workbook = new Workbook();

  // Determine which sheets to import
  const sheetsToImport = filterSheets(ctx.sheetInfos, opts.sheets);

  // Parse document properties
  if (opts.importProperties) {
    ctx.onProgress?.('properties', 0, 1);
    const corePropsXml = readFile('docProps/core.xml');
    if (corePropsXml) {
      parseProperties(corePropsXml, workbook);
    }
    ctx.onProgress?.('properties', 1, 1);
  }

  // Parse each worksheet
  ctx.onProgress?.('sheets', 0, sheetsToImport.length);
  for (let i = 0; i < sheetsToImport.length; i++) {
    const sheetInfo = sheetsToImport[i];
    const target = sheetTargets.get(sheetInfo.rId) || `worksheets/sheet${sheetInfo.sheetId}.xml`;
    const sheetPath = resolveRelPath('xl/', target);

    const sheetXml = readFile(sheetPath);
    if (!sheetXml) {
      ctx.warnings.push({
        code: 'MISSING_SHEET',
        message: `Sheet "${sheetInfo.name}" not found at ${sheetPath}`,
      });
      continue;
    }

    const sheet = workbook.addSheet(sheetInfo.name);
    parseWorksheet(sheetXml, sheet, ctx);
    ctx.stats.sheetCount++;

    const sheetRelsPath = sheetPath.replace(/\.xml$/, '.xml.rels').replace('worksheets/', 'worksheets/_rels/');
    const sheetRelsXml = readFile(sheetRelsPath);
    const hyperlinkTargets = new Map<string, string>(); // rId -> URL

    if (sheetRelsXml) {
      const rels = parseElements(sheetRelsXml, 'Relationship');
      for (const rel of rels) {
        const rId = getAttr(rel, 'Id');
        const type = getAttr(rel, 'Type');
        const target = getAttr(rel, 'Target');

        if (type?.includes('hyperlink') && rId && target) {
          hyperlinkTargets.set(rId, target);
        }
      }
    }

    if (opts.importComments) {
      const sheetIndex = sheetInfo.sheetId;
      const commentPaths = [
        `xl/comments${sheetIndex}.xml`,
        `xl/comments${i + 1}.xml`,
      ];

      if (sheetRelsXml) {
        const rels = parseElements(sheetRelsXml, 'Relationship');
        for (const rel of rels) {
          const type = getAttr(rel, 'Type');
          if (type?.includes('comments')) {
            const target = getAttr(rel, 'Target');
            if (target) {
              const commentPath = resolveRelPath(sheetPath.substring(0, sheetPath.lastIndexOf('/') + 1), target);
              commentPaths.unshift(commentPath);
            }
          }
        }
      }

      for (const commentPath of commentPaths) {
        const commentsXml = readFile(commentPath);
        if (commentsXml) {
          parseComments(commentsXml, sheet);
          break;
        }
      }
    }

    // Parse hyperlinks
    if (opts.importHyperlinks) {
      parseHyperlinks(sheetXml, sheet, hyperlinkTargets);
    }

    ctx.onProgress?.('sheets', i + 1, sheetsToImport.length);
  }

  // Calculate duration
  ctx.stats.durationMs = Math.round(performance.now() - startTime);

  return {
    workbook,
    stats: ctx.stats,
    warnings: ctx.warnings,
  };
}

/**
 * Import XLSX from Blob (browser convenience)
 *
 * @param blob - XLSX file as Blob
 * @param options - Import options
 * @returns Promise resolving to import result
 */
export async function xlsxBlobToWorkbook(
  blob: Blob,
  options?: XlsxImportOptions
): Promise<XlsxImportResult> {
  const buffer = new Uint8Array(await blob.arrayBuffer());
  return xlsxToWorkbook(buffer, options);
}

/**
 * Filter sheets based on import options
 */
function filterSheets(
  allSheets: SheetInfo[],
  filter: 'all' | string[] | number[]
): SheetInfo[] {
  if (filter === 'all') {
    return allSheets;
  }

  if (typeof filter[0] === 'number') {
    // Filter by index
    const indices = filter as number[];
    return allSheets.filter((_, i) => indices.includes(i));
  }

  // Filter by name
  const names = filter as string[];
  return allSheets.filter((s) => names.includes(s.name));
}

/**
 * Parse shared strings table
 */
function parseSharedStrings(xml: string, ctx: XlsxParseContext): void {
  const siElements = parseElements(xml, 'si');

  for (const si of siElements) {
    // Check for plain text
    const tContent = getTextContent(si.inner, 't');
    if (tContent !== undefined) {
      ctx.sharedStrings.push(tContent);
      continue;
    }

    // Check for rich text (concatenate all runs)
    const rElements = parseElements(si.inner, 'r');
    if (rElements.length > 0) {
      let text = '';
      for (const r of rElements) {
        const rText = getTextContent(r.inner, 't');
        if (rText) text += rText;
      }
      ctx.sharedStrings.push(text);
      continue;
    }

    // Empty or unknown format
    ctx.sharedStrings.push('');
  }
}

/**
 * Parse styles.xml
 */
function parseStyles(xml: string, ctx: XlsxParseContext): void {
  // Parse number formats
  const numFmts = parseElements(xml, 'numFmt');
  for (const fmt of numFmts) {
    const id = parseInt(getAttr(fmt, 'numFmtId') || '0', 10);
    const code = getAttr(fmt, 'formatCode') || '';
    ctx.numberFormats.set(id, code);
  }

  // Parse fonts
  const fontElements = parseElements(xml, 'font');
  for (const fontEl of fontElements) {
    const font: FontInfo = {};

    if (parseElement(fontEl.inner, 'b')) font.bold = true;
    if (parseElement(fontEl.inner, 'i')) font.italic = true;
    if (parseElement(fontEl.inner, 'u')) font.underline = true;
    if (parseElement(fontEl.inner, 'strike')) font.strike = true;

    const sz = parseElement(fontEl.inner, 'sz');
    if (sz) font.size = parseFloat(getAttr(sz, 'val') || '11');

    const color = parseElement(fontEl.inner, 'color');
    if (color) {
      const rgb = getAttr(color, 'rgb');
      if (rgb) font.color = '#' + rgb.slice(2); // Remove alpha
    }

    const name = parseElement(fontEl.inner, 'name');
    if (name) font.name = getAttr(name, 'val');

    ctx.fonts.push(font);
  }

  // Parse fills
  const fillElements = parseElements(xml, 'fill');
  for (const fillEl of fillElements) {
    const fill: FillInfo = {};

    const pattern = parseElement(fillEl.inner, 'patternFill');
    if (pattern) {
      fill.patternType = getAttr(pattern, 'patternType');

      const fgColor = parseElement(pattern.inner, 'fgColor');
      if (fgColor) {
        const rgb = getAttr(fgColor, 'rgb');
        if (rgb) fill.fgColor = '#' + rgb.slice(2);
      }

      const bgColor = parseElement(pattern.inner, 'bgColor');
      if (bgColor) {
        const rgb = getAttr(bgColor, 'rgb');
        if (rgb) fill.bgColor = '#' + rgb.slice(2);
      }
    }

    ctx.fills.push(fill);
  }

  // Parse borders
  const borderElements = parseElements(xml, 'border');
  for (const borderEl of borderElements) {
    const border: BorderInfo = {};

    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      const sideEl = parseElement(borderEl.inner, side);
      if (sideEl) {
        const sideInfo: BorderSideInfo = {};
        sideInfo.style = getAttr(sideEl, 'style');

        const color = parseElement(sideEl.inner, 'color');
        if (color) {
          const rgb = getAttr(color, 'rgb');
          if (rgb) sideInfo.color = '#' + rgb.slice(2);
        }

        border[side] = sideInfo;
      }
    }

    ctx.borders.push(border);
  }

  // Parse cellXfs (cell formats)
  // Find cellXfs section
  const cellXfsSection = parseElement(xml, 'cellXfs');
  if (cellXfsSection) {
    const cellXfElements = parseElements(cellXfsSection.inner, 'xf');
    for (const xf of cellXfElements) {
      const info: CellXfInfo = {
        fontId: parseInt(getAttr(xf, 'fontId') || '0', 10),
        fillId: parseInt(getAttr(xf, 'fillId') || '0', 10),
        borderId: parseInt(getAttr(xf, 'borderId') || '0', 10),
        numFmtId: parseInt(getAttr(xf, 'numFmtId') || '0', 10),
        applyFont: getAttr(xf, 'applyFont') === '1',
        applyFill: getAttr(xf, 'applyFill') === '1',
        applyBorder: getAttr(xf, 'applyBorder') === '1',
        applyNumberFormat: getAttr(xf, 'applyNumberFormat') === '1',
        applyAlignment: getAttr(xf, 'applyAlignment') === '1',
      };

      const alignment = parseElement(xf.inner, 'alignment');
      if (alignment) {
        info.alignment = {
          horizontal: getAttr(alignment, 'horizontal') as 'left' | 'center' | 'right' | undefined,
          vertical: getAttr(alignment, 'vertical') as 'top' | 'middle' | 'bottom' | undefined,
          wrapText: getAttr(alignment, 'wrapText') === '1',
          textRotation: parseInt(getAttr(alignment, 'textRotation') || '0', 10) || undefined,
        };
      }

      ctx.cellXfs.push(info);
    }
  }
}

/**
 * Parse document properties
 */
function parseProperties(xml: string, workbook: Workbook): void {
  const title = getTextContent(xml, 'dc:title') || getTextContent(xml, 'title');
  const creator = getTextContent(xml, 'dc:creator') || getTextContent(xml, 'creator');

  if (title) workbook.properties.title = title;
  if (creator) workbook.properties.author = creator;
}

/**
 * Parse worksheet XML
 */
function parseWorksheet(xml: string, sheet: Sheet, ctx: XlsxParseContext): void {
  const opts = ctx.options;

  // Parse dimensions (column widths)
  if (opts.importDimensions) {
    const colElements = parseElements(xml, 'col');
    for (const col of colElements) {
      const min = parseInt(getAttr(col, 'min') || '0', 10) - 1;
      const max = parseInt(getAttr(col, 'max') || '0', 10) - 1;
      const width = parseFloat(getAttr(col, 'width') || '0');
      const customWidth = getAttr(col, 'customWidth') === '1';

      if (customWidth && width > 0) {
        for (let c = min; c <= max; c++) {
          sheet.setColumnWidth(c, width);
        }
      }
    }
  }

  // Parse rows and cells
  const rowElements = parseElements(xml, 'row');

  for (const rowEl of rowElements) {
    const rowIndex = parseInt(getAttr(rowEl, 'r') || '0', 10) - 1;

    // Check row limit
    if (opts.maxRows > 0 && rowIndex >= opts.maxRows) continue;

    // Parse row height
    if (opts.importDimensions) {
      const ht = parseFloat(getAttr(rowEl, 'ht') || '0');
      const customHeight = getAttr(rowEl, 'customHeight') === '1';
      if (customHeight && ht > 0) {
        sheet.setRowHeight(rowIndex, ht);
      }
    }

    // Parse cells
    const cellElements = parseElements(rowEl.inner, 'c');

    for (const cellEl of cellElements) {
      const ref = getAttr(cellEl, 'r');
      if (!ref) continue;

      const { row, col } = parseCellRef(ref);

      // Check column limit
      if (opts.maxCols > 0 && col >= opts.maxCols) continue;

      // Get cell type and value
      const cellType = getAttr(cellEl, 't');
      const styleIndex = parseInt(getAttr(cellEl, 's') || '0', 10);

      // Get value and formula
      const valueStr = getTextContent(cellEl.inner, 'v');
      const formulaStr = opts.importFormulas ? getTextContent(cellEl.inner, 'f') : undefined;

      // Parse cell value based on type
      let value = parseCellValue(cellType, valueStr, cellEl.inner, ctx, styleIndex);

      // Get the cell and set value
      const cell = sheet.cell(row, col);

      if (formulaStr) {
        cell.setFormula(formulaStr, value);
        ctx.stats.formulaCells++;
        // Only count toward totalCells if there's a cached result
        if (value !== undefined && value !== null) {
          ctx.stats.totalCells++;
        }
      } else if (value !== undefined && value !== null) {
        cell.value = value;
        ctx.stats.totalCells++;
      }

      // Apply style
      if (opts.importStyles && styleIndex > 0 && ctx.cellXfs[styleIndex]) {
        const style = buildCellStyle(ctx.cellXfs[styleIndex], ctx);
        if (style && Object.keys(style).length > 0) {
          cell.style = style;
        }
      }
    }
  }

  // Parse merged cells
  if (opts.importMergedCells) {
    const mergeCells = parseElements(xml, 'mergeCell');
    for (const merge of mergeCells) {
      const ref = getAttr(merge, 'ref');
      if (!ref) continue;

      try {
        // Use A1-style string reference directly
        sheet.mergeCells(ref);
        ctx.stats.mergedRanges++;
      } catch (error) {
        ctx.warnings.push({
          code: 'INVALID_MERGE',
          message: `Invalid merge range: ${ref}`,
          location: ref,
        });
      }
    }
  }

  // Parse freeze panes
  if (opts.importFreezePanes) {
    const pane = parseElement(xml, 'pane');
    if (pane) {
      const xSplit = parseInt(getAttr(pane, 'xSplit') || '0', 10);
      const ySplit = parseInt(getAttr(pane, 'ySplit') || '0', 10);
      const state = getAttr(pane, 'state');

      if (state === 'frozen' && (xSplit > 0 || ySplit > 0)) {
        sheet.freeze(ySplit, xSplit);
      }
    }
  }

  // Parse autofilter
  const autoFilter = parseElement(xml, 'autoFilter');
  if (autoFilter) {
    const ref = getAttr(autoFilter, 'ref');
    if (ref) {
      // Use A1-style string reference directly
      sheet.setAutoFilter(ref);
    }
  }
}

/**
 * Parse comments XML and apply to sheet
 */
function parseComments(xml: string, sheet: Sheet): void {
  const authors: string[] = [];
  const authorElements = parseElements(xml, 'author');
  for (const authorEl of authorElements) {
    // Author element contains text directly
    authors.push(unescapeXml(authorEl.inner) || '');
  }

  const commentElements = parseElements(xml, 'comment');
  for (const commentEl of commentElements) {
    const ref = getAttr(commentEl, 'ref');
    const authorId = parseInt(getAttr(commentEl, 'authorId') || '0', 10);

    if (!ref) continue;

    const tElements = parseElements(commentEl.inner, 't');
    const textContent = unescapeXml(tElements.map(t => t.inner).join(''));
    if (!textContent) continue;

    const { row, col } = a1ToAddress(ref);
    const author = authors[authorId] || undefined;

    const cell = sheet.cell(row, col);
    cell.setComment(textContent, author);
  }
}

/**
 * Parse hyperlinks from worksheet XML and apply to sheet
 */
function parseHyperlinks(
  xml: string,
  sheet: Sheet,
  hyperlinkTargets: Map<string, string>
): void {
  const hyperlinkElements = parseElements(xml, 'hyperlink');

  for (const hyperlinkEl of hyperlinkElements) {
    const ref = getAttr(hyperlinkEl, 'ref');
    if (!ref) continue;

    const rId = getAttr(hyperlinkEl, 'r:id');
    const location = getAttr(hyperlinkEl, 'location');
    const tooltip = getAttr(hyperlinkEl, 'tooltip');
    const display = getAttr(hyperlinkEl, 'display');

    let target: string | undefined;

    if (rId) {
      target = hyperlinkTargets.get(rId);
    } else if (location) {
      target = `#${location}`;
    }

    if (!target) continue;
    const { row, col } = a1ToAddress(ref);
    const cell = sheet.cell(row, col);

    if (display) {
      cell.setHyperlink(target, tooltip);
    } else {
      cell.setHyperlink(target, tooltip);
    }
  }
}

/**
 * Parse cell value based on type attribute
 */
function parseCellValue(
  cellType: string | undefined,
  valueStr: string | undefined,
  cellInner: string,
  ctx: XlsxParseContext,
  styleIndex: number
): string | number | boolean | Date | null {
  if (valueStr === undefined && cellType !== 'inlineStr') {
    return null;
  }

  switch (cellType) {
    case 's': {
      // Shared string
      const index = parseInt(valueStr || '0', 10);
      return ctx.sharedStrings[index] ?? '';
    }

    case 'b': {
      // Boolean
      return valueStr === '1';
    }

    case 'e': {
      // Error - return as string
      return valueStr || '#VALUE!';
    }

    case 'str': {
      // Formula string result
      return unescapeXml(valueStr || '');
    }

    case 'inlineStr': {
      // Inline string
      const tContent = getTextContent(cellInner, 't');
      if (tContent !== undefined) return tContent;

      // Rich text inline
      const is = parseElement(cellInner, 'is');
      if (is) {
        const rElements = parseElements(is.inner, 'r');
        if (rElements.length > 0) {
          let text = '';
          for (const r of rElements) {
            const rText = getTextContent(r.inner, 't');
            if (rText) text += rText;
          }
          return text;
        }
        return getTextContent(is.inner, 't') || '';
      }
      return '';
    }

    default: {
      // Number (or date)
      const num = parseFloat(valueStr || '0');

      // Check if this is a date format
      if (ctx.options.importStyles && styleIndex < ctx.cellXfs.length) {
        const xf = ctx.cellXfs[styleIndex];
        if (xf && isDateFormat(xf.numFmtId, ctx.numberFormats.get(xf.numFmtId))) {
          return excelSerialToDate(num);
        }
      }

      return num;
    }
  }
}

/**
 * Check if a number format represents a date/time
 */
function isDateFormat(numFmtId: number, formatCode?: string): boolean {
  // Check built-in date format IDs
  if (DATE_FORMAT_IDS.has(numFmtId)) {
    return true;
  }

  // Check custom format code for date/time patterns
  if (formatCode) {
    // Look for date/time tokens (but not in text sections)
    const withoutTextSections = formatCode.replace(/"[^"]*"/g, '');
    // Date tokens: d, m, y, h, s (but not [m] which is minutes, and not standalone s/h in number formats)
    if (/[dmyhs]/i.test(withoutTextSections) && !/[0#?]/.test(withoutTextSections)) {
      return true;
    }
  }

  return false;
}

/**
 * Build CellStyle from parsed style info
 */
function buildCellStyle(xf: CellXfInfo, ctx: XlsxParseContext): CellStyle | undefined {
  const style: CellStyle = {};

  // Font
  if (xf.fontId > 0 && xf.fontId < ctx.fonts.length) {
    const font = ctx.fonts[xf.fontId];
    if (font) {
      style.font = {};
      if (font.bold) style.font.bold = true;
      if (font.italic) style.font.italic = true;
      if (font.underline) style.font.underline = 'single';
      if (font.strike) style.font.strikethrough = true;
      if (font.size) style.font.size = font.size;
      if (font.color) style.font.color = font.color;
      if (font.name) style.font.name = font.name;

      // Remove empty font object
      if (Object.keys(style.font).length === 0) {
        delete style.font;
      }
    }
  }

  // Fill
  if (xf.fillId > 0 && xf.fillId < ctx.fills.length) {
    const fill = ctx.fills[xf.fillId];
    if (fill && fill.patternType === 'solid' && fill.fgColor) {
      style.fill = {
        type: 'pattern',
        pattern: 'solid',
        foregroundColor: fill.fgColor,
      };
    }
  }

  // Border
  if (xf.borderId > 0 && xf.borderId < ctx.borders.length) {
    const border = ctx.borders[xf.borderId];
    if (border) {
      const borders: CellBorders = {};
      let hasBorder = false;

      for (const side of ['left', 'right', 'top', 'bottom'] as const) {
        const sideInfo = border[side];
        if (sideInfo?.style && sideInfo.style !== 'none') {
          const bs: Border = {
            style: mapBorderStyle(sideInfo.style),
            color: sideInfo.color || '#000000',
          };
          borders[side] = bs;
          hasBorder = true;
        }
      }

      if (hasBorder) {
        style.borders = borders;
      }
    }
  }

  // Alignment
  if (xf.alignment) {
    style.alignment = {};
    if (xf.alignment.horizontal) {
      style.alignment.horizontal = xf.alignment.horizontal as 'left' | 'center' | 'right';
    }
    if (xf.alignment.vertical) {
      style.alignment.vertical = xf.alignment.vertical as 'top' | 'middle' | 'bottom';
    }
    if (xf.alignment.wrapText) {
      style.alignment.wrapText = true;
    }
    if (xf.alignment.textRotation) {
      style.alignment.textRotation = xf.alignment.textRotation;
    }

    // Remove empty alignment object
    if (Object.keys(style.alignment).length === 0) {
      delete style.alignment;
    }
  }

  // Number format
  if (xf.numFmtId > 0) {
    const formatCode = ctx.numberFormats.get(xf.numFmtId);
    if (formatCode) {
      style.numberFormat = { formatCode };
    }
  }

  return Object.keys(style).length > 0 ? style : undefined;
}

/**
 * Map Excel border style to Cellify style
 */
function mapBorderStyle(excelStyle: string): BorderStyleType {
  const styleMap: Record<string, BorderStyleType> = {
    thin: 'thin',
    medium: 'medium',
    thick: 'thick',
    dashed: 'dashed',
    dotted: 'dotted',
    double: 'double',
    hair: 'hair',
    dashDot: 'dashDot',
    dashDotDot: 'dashDotDot',
    mediumDashed: 'mediumDashed',
    mediumDashDot: 'mediumDashDot',
    mediumDashDotDot: 'mediumDashDotDot',
    slantDashDot: 'slantDashDot',
  };

  return styleMap[excelStyle] || 'thin';
}

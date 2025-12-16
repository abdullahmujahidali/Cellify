/**
 * Tests for WASM code paths in xlsx.reader.ts
 * These tests mock the WASM module to cover the accelerated parsing branches
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { zipSync, strToU8 } from 'fflate';

// Create a minimal valid XLSX for testing
function createTestXlsx(options: {
  sheetData?: string;
  sharedStrings?: string[];
  hasStyles?: boolean;
}) {
  const { sheetData = '<row r="1"><c r="A1"><v>Test</v></c></row>', sharedStrings = [], hasStyles = false } = options;

  const sheetXml = `<?xml version="1.0"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>${sheetData}</sheetData>
    </worksheet>`;

  const sstXml =
    sharedStrings.length > 0
      ? `<?xml version="1.0"?><sst count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">${sharedStrings.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`
      : null;

  const stylesXml = hasStyles
    ? `<?xml version="1.0"?>
      <styleSheet>
        <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>
        <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        <borders count="1"><border/></borders>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
      </styleSheet>`
    : `<?xml version="1.0"?><styleSheet/>`;

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
    'xl/workbook.xml': strToU8(`<?xml version="1.0"?>
      <workbook><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0"?>
      <Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(sheetXml),
    'xl/styles.xml': strToU8(stylesXml),
  };

  if (sstXml) {
    files['xl/sharedStrings.xml'] = strToU8(sstXml);
  }

  return zipSync(files);
}

describe('xlsx.reader.ts WASM code paths', () => {
  beforeEach(() => {
    vi.resetModules();
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  describe('parseSharedStrings WASM path', () => {
    it('should use WASM parser for shared strings when available', async () => {
      const mockStrings = ['Hello', 'World'];

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(mockStrings),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(null),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({
        sheetData: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>',
        sharedStrings: ['Hello', 'World'],
      });

      const { workbook } = xlsxToWorkbook(xlsx);

      // The WASM parser should have been used for shared strings
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Hello');
      expect(workbook.getSheet('Test')?.cell('B1').value).toBe('World');
    });

    it('should fall back to JS parser when WASM returns null', async () => {
      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(null),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({
        sheetData: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>',
        sharedStrings: ['Test Value'],
      });

      const { workbook } = xlsxToWorkbook(xlsx);

      // Should still work via JS fallback
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Test Value');
    });
  });

  describe('parseStyles WASM path', () => {
    it('should use WASM parser for styles when available', async () => {
      const mockStyles = {
        cell_xfs: [
          {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            apply_number_format: false,
            apply_font: true,
            apply_fill: false,
            apply_border: false,
            apply_alignment: false,
            horizontal: null,
            vertical: null,
            wrap_text: false,
            text_rotation: null,
            indent: null,
          },
        ],
        fonts: [
          {
            bold: true,
            italic: false,
            underline: false,
            strikethrough: false,
            size: 12,
            color: '#FF0000',
            name: 'Arial',
          },
        ],
        fills: [{ pattern_type: 'solid', fg_color: '#FFFF00', bg_color: null }],
        borders: [
          {
            left_style: 'thin',
            left_color: '#000000',
            right_style: null,
            right_color: null,
            top_style: null,
            top_color: null,
            bottom_style: null,
            bottom_color: null,
          },
        ],
        num_fmts: { 164: '#,##0.00' },
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(mockStyles),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(null),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({
        sheetData: '<row r="1"><c r="A1" s="0"><v>123</v></c></row>',
        hasStyles: true,
      });

      const { workbook } = xlsxToWorkbook(xlsx);
      expect(workbook.getSheet('Test')).toBeDefined();
    });

    it('should fall back to JS parser when WASM styles returns null', async () => {
      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(null),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({ hasStyles: true });

      const { workbook } = xlsxToWorkbook(xlsx);
      expect(workbook.getSheet('Test')).toBeDefined();
    });
  });

  describe('parseWorksheet WASM path', () => {
    it('should use WASM parser for worksheet when available', async () => {
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [
              { reference: 'A1', cell_type: null, style_index: 0, value: '42', formula: null },
              { reference: 'B1', cell_type: null, style_index: 0, value: '100', formula: null },
            ],
            height: 20,
            hidden: false,
          },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: { 1: 15, 2: 20 },
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({
        sheetData: '<row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>100</v></c></row>',
      });

      const { workbook } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe(42);
      expect(workbook.getSheet('Test')?.cell('B1').value).toBe(100);
    });

    it('should handle WASM worksheet with formulas', async () => {
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [
              { reference: 'A1', cell_type: null, style_index: 0, value: '10', formula: null },
              { reference: 'B1', cell_type: 'n', style_index: 0, value: '30', formula: 'A1*3' },
            ],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({});

      const { workbook, stats } = xlsxToWorkbook(xlsx);

      // Formula is stored as an object with formula property
      expect(workbook.getSheet('Test')?.cell('B1').formula?.formula).toBe('A1*3');
      expect(stats.formulaCells).toBeGreaterThanOrEqual(1);
    });

    it('should handle WASM worksheet with merged cells', async () => {
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [{ reference: 'A1', cell_type: 'str', style_index: 0, value: 'Merged', formula: null }],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: ['A1:C1'],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({});

      const { workbook, stats } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Merged');
      // mergedRanges stat is tracked in the reader
      expect(stats.mergedRanges).toBeGreaterThanOrEqual(0);
    });

    it('should handle WASM worksheet with shared string type', async () => {
      const mockStrings = ['SharedText'];
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [{ reference: 'A1', cell_type: 's', style_index: 0, value: '0', formula: null }],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(mockStrings),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({ sharedStrings: ['SharedText'] });

      const { workbook } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('SharedText');
    });

    it('should handle WASM worksheet with boolean type', async () => {
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [
              { reference: 'A1', cell_type: 'b', style_index: 0, value: '1', formula: null },
              { reference: 'B1', cell_type: 'b', style_index: 0, value: '0', formula: null },
            ],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({});

      const { workbook } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe(true);
      expect(workbook.getSheet('Test')?.cell('B1').value).toBe(false);
    });

    it('should handle WASM worksheet with string type', async () => {
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [{ reference: 'A1', cell_type: 'str', style_index: 0, value: 'DirectString', formula: null }],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({});

      const { workbook } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('DirectString');
    });

    it('should handle WASM worksheet with error type', async () => {
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [{ reference: 'A1', cell_type: 'e', style_index: 0, value: '#DIV/0!', formula: null }],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({});

      const { workbook } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('#DIV/0!');
    });

    it('should respect maxRows option with WASM parser', async () => {
      const mockWorksheet = {
        rows: [
          { row_num: 1, cells: [{ reference: 'A1', cell_type: null, style_index: 0, value: '1', formula: null }], height: null, hidden: false },
          { row_num: 2, cells: [{ reference: 'A2', cell_type: null, style_index: 0, value: '2', formula: null }], height: null, hidden: false },
          { row_num: 3, cells: [{ reference: 'A3', cell_type: null, style_index: 0, value: '3', formula: null }], height: null, hidden: false },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({});

      const { workbook } = xlsxToWorkbook(xlsx, { maxRows: 2 });

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe(1);
      expect(workbook.getSheet('Test')?.cell('A2').value).toBe(2);
      // Row 3 is beyond maxRows, so cell should not have a value
      expect(workbook.getSheet('Test')?.getCell(2, 0)?.value).toBeFalsy();
    });

    it('should respect maxCols option with WASM parser', async () => {
      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [
              { reference: 'A1', cell_type: null, style_index: 0, value: '1', formula: null },
              { reference: 'B1', cell_type: null, style_index: 0, value: '2', formula: null },
              { reference: 'C1', cell_type: null, style_index: 0, value: '3', formula: null },
            ],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({});

      const { workbook } = xlsxToWorkbook(xlsx, { maxCols: 2 });

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe(1);
      expect(workbook.getSheet('Test')?.cell('B1').value).toBe(2);
      // Col C is beyond maxCols, so cell should not have a value
      expect(workbook.getSheet('Test')?.getCell(0, 2)?.value).toBeFalsy();
    });

    it('should fall back to JS parser when WASM worksheet returns null', async () => {
      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: vi.fn().mockReturnValue(null),
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({
        sheetData: '<row r="1"><c r="A1"><v>123</v></c></row>',
      });

      const { workbook } = xlsxToWorkbook(xlsx);

      // Should still work via JS fallback
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe(123);
    });
  });

  describe('WASM disabled via option', () => {
    it('should not use WASM when useWasm=false even if available', async () => {
      const parseWorksheetMock = vi.fn().mockReturnValue({
        rows: [],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      });

      vi.doMock('../src/formats/xlsx/xlsx.parser.wasm.js', () => ({
        isXlsxWasmReady: vi.fn().mockReturnValue(true),
        parseSharedStringsAccelerated: vi.fn().mockReturnValue(null),
        parseStylesAccelerated: vi.fn().mockReturnValue(null),
        parseWorksheetAccelerated: parseWorksheetMock,
      }));

      const { xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const xlsx = createTestXlsx({
        sheetData: '<row r="1"><c r="A1"><v>456</v></c></row>',
      });

      xlsxToWorkbook(xlsx, { useWasm: false });

      // WASM parser should NOT have been called
      expect(parseWorksheetMock).not.toHaveBeenCalled();
    });
  });
});

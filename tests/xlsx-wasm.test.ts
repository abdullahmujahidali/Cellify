/**
 * Tests for WASM parser wrapper
 * Tests fallback behavior and WASM integration
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mock the WASM module before importing the wasm wrapper
vi.mock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
  default: vi.fn().mockResolvedValue(undefined),
  init: vi.fn(),
  parse_worksheet: vi.fn(),
  parse_shared_strings: vi.fn(),
  parse_styles: vi.fn(),
  parse_workbook: vi.fn(),
  parse_relationships: vi.fn(),
}));

describe('xlsx.wasm.ts', () => {
  beforeEach(() => {
    vi.resetModules();
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  describe('initWasm', () => {
    it('should initialize WASM module successfully', async () => {
      const { initWasm, isWasmAvailable } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result = await initWasm();

      expect(result).toBe(true);
      expect(isWasmAvailable()).toBe(true);
    });

    it('should return same result on multiple calls', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
      }));

      const { initWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result1 = await initWasm();
      const result2 = await initWasm();

      expect(result1).toBe(result2);
      expect(result1).toBe(true);
    });

    it('should handle WASM load failure gracefully', async () => {
      vi.resetModules();

      // Mock to throw an error
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('WASM not supported')),
      }));

      const { initWasm, isWasmAvailable } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result = await initWasm();

      expect(result).toBe(false);
      expect(isWasmAvailable()).toBe(false);
    });
  });

  describe('parseWorksheetWasm', () => {
    it('should return null when WASM not initialized', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseWorksheetWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result = parseWorksheetWasm('<worksheet/>');
      expect(result).toBeNull();
    });

    it('should parse worksheet when WASM is available', async () => {
      vi.resetModules();

      const mockWorksheet = {
        rows: [{ row_num: 1, cells: [], height: null, hidden: false }],
        merge_cells: [],
        hyperlinks: [],
        col_widths: {},
      };

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_worksheet: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { initWasm, parseWorksheetWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseWorksheetWasm('<worksheet/>');

      expect(result).toEqual(mockWorksheet);
    });

    it('should return null on parse error', async () => {
      vi.resetModules();

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_worksheet: vi.fn().mockImplementation(() => {
          throw new Error('Parse error');
        }),
      }));

      const { initWasm, parseWorksheetWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseWorksheetWasm('<invalid>');

      expect(result).toBeNull();
    });
  });

  describe('parseSharedStringsWasm', () => {
    it('should return null when WASM not initialized', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseSharedStringsWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result = parseSharedStringsWasm('<sst/>');
      expect(result).toBeNull();
    });

    it('should parse shared strings when WASM is available', async () => {
      vi.resetModules();

      const mockStrings = ['Hello', 'World'];

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_shared_strings: vi.fn().mockReturnValue(mockStrings),
      }));

      const { initWasm, parseSharedStringsWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseSharedStringsWasm('<sst><si><t>Hello</t></si></sst>');

      expect(result).toEqual(mockStrings);
    });

    it('should return null on parse error', async () => {
      vi.resetModules();

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_shared_strings: vi.fn().mockImplementation(() => {
          throw new Error('Parse error');
        }),
      }));

      const { initWasm, parseSharedStringsWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseSharedStringsWasm('<invalid>');

      expect(result).toBeNull();
    });
  });

  describe('parseStylesWasm', () => {
    it('should return null when WASM not initialized', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseStylesWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result = parseStylesWasm('<styleSheet/>');
      expect(result).toBeNull();
    });

    it('should parse styles when WASM is available', async () => {
      vi.resetModules();

      const mockStyles = {
        cell_xfs: [],
        fonts: [],
        fills: [],
        borders: [],
        num_fmts: {},
      };

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_styles: vi.fn().mockReturnValue(mockStyles),
      }));

      const { initWasm, parseStylesWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseStylesWasm('<styleSheet/>');

      expect(result).toEqual(mockStyles);
    });

    it('should return null on parse error', async () => {
      vi.resetModules();

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_styles: vi.fn().mockImplementation(() => {
          throw new Error('Parse error');
        }),
      }));

      const { initWasm, parseStylesWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseStylesWasm('<invalid>');

      expect(result).toBeNull();
    });
  });

  describe('parseWorkbookWasm', () => {
    it('should return null when WASM not initialized', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseWorkbookWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result = parseWorkbookWasm('<workbook/>');
      expect(result).toBeNull();
    });

    it('should parse workbook when WASM is available', async () => {
      vi.resetModules();

      const mockSheets = [{ name: 'Sheet1', sheet_id: 1, rid: 'rId1', state: null }];

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_workbook: vi.fn().mockReturnValue(mockSheets),
      }));

      const { initWasm, parseWorkbookWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseWorkbookWasm('<workbook/>');

      expect(result).toEqual(mockSheets);
    });

    it('should return null on parse error', async () => {
      vi.resetModules();

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_workbook: vi.fn().mockImplementation(() => {
          throw new Error('Parse error');
        }),
      }));

      const { initWasm, parseWorkbookWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseWorkbookWasm('<invalid>');

      expect(result).toBeNull();
    });
  });

  describe('parseRelationshipsWasm', () => {
    it('should return null when WASM not initialized', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseRelationshipsWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      const result = parseRelationshipsWasm('<Relationships/>');
      expect(result).toBeNull();
    });

    it('should parse relationships when WASM is available', async () => {
      vi.resetModules();

      const mockRels = [{ id: 'rId1', rel_type: 'worksheet', target: 'sheet1.xml', target_mode: null }];

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_relationships: vi.fn().mockReturnValue(mockRels),
      }));

      const { initWasm, parseRelationshipsWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseRelationshipsWasm('<Relationships/>');

      expect(result).toEqual(mockRels);
    });

    it('should return null on parse error', async () => {
      vi.resetModules();

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_relationships: vi.fn().mockImplementation(() => {
          throw new Error('Parse error');
        }),
      }));

      const { initWasm, parseRelationshipsWasm } = await import('../src/formats/xlsx/xlsx.wasm.js');

      await initWasm();
      const result = parseRelationshipsWasm('<invalid>');

      expect(result).toBeNull();
    });
  });
});

describe('xlsx.parser.wasm.ts', () => {
  beforeEach(() => {
    vi.resetModules();
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  describe('initXlsxWasm', () => {
    it('should return WASM availability status', async () => {
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
      }));

      const { initXlsxWasm } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      const result = await initXlsxWasm();
      expect(typeof result).toBe('boolean');
    });

    it('should return cached result on subsequent calls', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
      }));

      const { initXlsxWasm, isXlsxWasmReady } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      await initXlsxWasm();
      const result = await initXlsxWasm();

      expect(result).toBe(isXlsxWasmReady());
    });
  });

  describe('isXlsxWasmReady', () => {
    it('should return false before initialization', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { isXlsxWasmReady } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      expect(isXlsxWasmReady()).toBe(false);
    });
  });

  describe('parseSharedStringsAccelerated', () => {
    it('should return null when WASM not ready', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseSharedStringsAccelerated } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      const result = parseSharedStringsAccelerated('<sst/>');
      expect(result).toBeNull();
    });
  });

  describe('parseWorksheetAccelerated', () => {
    it('should return null when WASM not ready', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseWorksheetAccelerated } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      const result = parseWorksheetAccelerated('<worksheet/>');
      expect(result).toBeNull();
    });
  });

  describe('parseStylesAccelerated', () => {
    it('should return null when WASM not ready', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseStylesAccelerated } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      const result = parseStylesAccelerated('<styleSheet/>');
      expect(result).toBeNull();
    });
  });

  describe('parseWorkbookAccelerated', () => {
    it('should return null when WASM not ready', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseWorkbookAccelerated } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      const result = parseWorkbookAccelerated('<workbook/>');
      expect(result).toBeNull();
    });
  });

  describe('parseRelationshipsAccelerated', () => {
    it('should return null when WASM not ready', async () => {
      vi.resetModules();
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockRejectedValue(new Error('Not available')),
      }));

      const { parseRelationshipsAccelerated } = await import('../src/formats/xlsx/xlsx.parser.wasm.js');

      const result = parseRelationshipsAccelerated('<Relationships/>');
      expect(result).toBeNull();
    });
  });
});

describe('xlsx.reader.ts WASM Integration', () => {
  describe('useWasm option', () => {
    it('should use JS parser when useWasm=false', async () => {
      const { Workbook } = await import('../src/core/Workbook.js');
      const { workbookToXlsx, xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const wb = new Workbook();
      wb.addSheet('Test').cell('A1').value = 'Hello';
      const xlsx = workbookToXlsx(wb);

      // Import with WASM disabled
      const { workbook } = xlsxToWorkbook(xlsx, { useWasm: false });

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Hello');
    });

    it('should work with default useWasm=true (falls back to JS)', async () => {
      const { Workbook } = await import('../src/core/Workbook.js');
      const { workbookToXlsx, xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const wb = new Workbook();
      wb.addSheet('Test').cell('A1').value = 'World';
      const xlsx = workbookToXlsx(wb);

      // Import with default options (useWasm=true, but WASM unavailable so falls back)
      const { workbook } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('World');
    });

    it('should handle complex workbook with styles when WASM unavailable', async () => {
      const { Workbook } = await import('../src/core/Workbook.js');
      const { workbookToXlsx, xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const wb = new Workbook();
      const sheet = wb.addSheet('Styled');
      sheet.cell('A1').value = 'Bold';
      sheet.cell('A1').style = { font: { bold: true } };
      sheet.cell('B1').value = 42;
      sheet.cell('B1').style = { numberFormat: { formatCode: '#,##0' } };
      sheet.mergeCells('C1:D1');
      sheet.cell('C1').value = 'Merged';

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      expect(workbook.getSheet('Styled')?.cell('A1').value).toBe('Bold');
      expect(workbook.getSheet('Styled')?.cell('B1').value).toBe(42);
    });

    it('should import multiple sheets correctly', async () => {
      const { Workbook } = await import('../src/core/Workbook.js');
      const { workbookToXlsx, xlsxToWorkbook } = await import('../src/formats/xlsx/index.js');

      const wb = new Workbook();
      wb.addSheet('Sheet1').cell('A1').value = 'First';
      wb.addSheet('Sheet2').cell('A1').value = 'Second';
      wb.addSheet('Sheet3').cell('A1').value = 'Third';

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx);

      expect(stats.sheetCount).toBe(3);
      expect(workbook.getSheet('Sheet1')?.cell('A1').value).toBe('First');
      expect(workbook.getSheet('Sheet2')?.cell('A1').value).toBe('Second');
      expect(workbook.getSheet('Sheet3')?.cell('A1').value).toBe('Third');
    });
  });
});

describe('xlsx.parser.wasm.ts accelerated functions', () => {
  describe('when WASM is initialized', () => {
    it('parseSharedStringsAccelerated returns data when WASM ready', async () => {
      vi.resetModules();

      const mockStrings = ['String1', 'String2'];
      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_shared_strings: vi.fn().mockReturnValue(mockStrings),
      }));

      const { initXlsxWasm, parseSharedStringsAccelerated } = await import(
        '../src/formats/xlsx/xlsx.parser.wasm.js'
      );

      await initXlsxWasm();
      const result = parseSharedStringsAccelerated('<sst/>');

      expect(result).toEqual(mockStrings);
    });

    it('parseWorksheetAccelerated returns data when WASM ready', async () => {
      vi.resetModules();

      const mockWorksheet = {
        rows: [
          {
            row_num: 1,
            cells: [{ reference: 'A1', cell_type: 's', style_index: 0, value: '0', formula: null }],
            height: null,
            hidden: false,
          },
        ],
        merge_cells: ['A1:B1'],
        hyperlinks: [],
        col_widths: { 1: 15 },
      };

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_worksheet: vi.fn().mockReturnValue(mockWorksheet),
      }));

      const { initXlsxWasm, parseWorksheetAccelerated } = await import(
        '../src/formats/xlsx/xlsx.parser.wasm.js'
      );

      await initXlsxWasm();
      const result = parseWorksheetAccelerated('<worksheet/>');

      expect(result).toEqual(mockWorksheet);
    });

    it('parseStylesAccelerated returns data when WASM ready', async () => {
      vi.resetModules();

      const mockStyles = {
        cell_xfs: [
          {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            xf_id: null,
            apply_number_format: false,
            apply_font: false,
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
        fonts: [{ bold: false, italic: false, underline: false, strikethrough: false, size: 11, color: null, name: 'Calibri' }],
        fills: [{ pattern_type: 'none', fg_color: null, bg_color: null }],
        borders: [{ left_style: null, left_color: null, right_style: null, right_color: null, top_style: null, top_color: null, bottom_style: null, bottom_color: null }],
        num_fmts: {},
      };

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_styles: vi.fn().mockReturnValue(mockStyles),
      }));

      const { initXlsxWasm, parseStylesAccelerated } = await import(
        '../src/formats/xlsx/xlsx.parser.wasm.js'
      );

      await initXlsxWasm();
      const result = parseStylesAccelerated('<styleSheet/>');

      expect(result).toEqual(mockStyles);
    });

    it('parseWorkbookAccelerated returns data when WASM ready', async () => {
      vi.resetModules();

      const mockSheets = [
        { name: 'Sheet1', sheet_id: 1, rid: 'rId1', state: null },
        { name: 'Sheet2', sheet_id: 2, rid: 'rId2', state: null },
      ];

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_workbook: vi.fn().mockReturnValue(mockSheets),
      }));

      const { initXlsxWasm, parseWorkbookAccelerated } = await import(
        '../src/formats/xlsx/xlsx.parser.wasm.js'
      );

      await initXlsxWasm();
      const result = parseWorkbookAccelerated('<workbook/>');

      expect(result).toEqual(mockSheets);
    });

    it('parseRelationshipsAccelerated returns data when WASM ready', async () => {
      vi.resetModules();

      const mockRels = [
        { id: 'rId1', rel_type: 'worksheet', target: 'worksheets/sheet1.xml', target_mode: null },
        { id: 'rId2', rel_type: 'styles', target: 'styles.xml', target_mode: null },
      ];

      vi.doMock('../src/formats/xlsx/wasm/cellify_wasm.js', () => ({
        default: vi.fn().mockResolvedValue(undefined),
        init: vi.fn(),
        parse_relationships: vi.fn().mockReturnValue(mockRels),
      }));

      const { initXlsxWasm, parseRelationshipsAccelerated } = await import(
        '../src/formats/xlsx/xlsx.parser.wasm.js'
      );

      await initXlsxWasm();
      const result = parseRelationshipsAccelerated('<Relationships/>');

      expect(result).toEqual(mockRels);
    });
  });
});

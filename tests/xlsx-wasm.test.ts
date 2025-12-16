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
  });
});

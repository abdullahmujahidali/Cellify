import { describe, it, expect } from 'vitest';
import { unzipSync } from 'fflate';
import type { Unzipped } from 'fflate';
import { Workbook } from '../src/core/Workbook.js';
import { workbookToXlsx, xlsxToWorkbook } from '../src/formats/xlsx/index.js';

/**
 * Cache for unzipped files to avoid repeated decompression
 */
const unzipCache = new WeakMap<Uint8Array, Unzipped>();

function getUnzippedFiles(xlsx: Uint8Array): Unzipped {
  let files = unzipCache.get(xlsx);
  if (!files) {
    files = unzipSync(xlsx);
    unzipCache.set(xlsx, files);
  }
  return files;
}

/**
 * Helper to extract and decode a file from the XLSX ZIP
 */
function getXmlFile(xlsx: Uint8Array, path: string): string {
  const files = getUnzippedFiles(xlsx);
  const data = files[path];
  if (!data) {
    throw new Error(`File not found in XLSX: ${path}`);
  }
  return new TextDecoder().decode(data);
}

/**
 * Helper to check if a file exists in the XLSX ZIP
 */
function hasFile(xlsx: Uint8Array, path: string): boolean {
  const files = getUnzippedFiles(xlsx);
  return path in files;
}

describe('XLSX Export', () => {
  describe('Basic Export', () => {
    it('should export empty workbook', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1');

      const xlsx = workbookToXlsx(wb);

      expect(xlsx).toBeInstanceOf(Uint8Array);
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it('should create valid ZIP structure', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1');

      const xlsx = workbookToXlsx(wb);

      // Check required files exist
      expect(hasFile(xlsx, '[Content_Types].xml')).toBe(true);
      expect(hasFile(xlsx, '_rels/.rels')).toBe(true);
      expect(hasFile(xlsx, 'xl/workbook.xml')).toBe(true);
      expect(hasFile(xlsx, 'xl/styles.xml')).toBe(true);
      expect(hasFile(xlsx, 'xl/worksheets/sheet1.xml')).toBe(true);
      expect(hasFile(xlsx, 'xl/_rels/workbook.xml.rels')).toBe(true);
    });

    it('should export string values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Hello';
      sheet.cell('B1').value = 'World';

      const xlsx = workbookToXlsx(wb);

      // Check shared strings
      const sst = getXmlFile(xlsx, 'xl/sharedStrings.xml');
      expect(sst).toContain('<t>Hello</t>');
      expect(sst).toContain('<t>World</t>');

      // Check sheet references shared strings
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');
      expect(sheetXml).toContain('t="s"'); // String type
    });

    it('should export number values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 42;
      sheet.cell('B1').value = 3.14159;
      sheet.cell('C1').value = -100;

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('<v>42</v>');
      expect(sheetXml).toContain('<v>3.14159</v>');
      expect(sheetXml).toContain('<v>-100</v>');
    });

    it('should export boolean values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = true;
      sheet.cell('B1').value = false;

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('t="b"'); // Boolean type
      expect(sheetXml).toContain('<v>1</v>');
      expect(sheetXml).toContain('<v>0</v>');
    });

    it('should export date values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      // January 1, 2024 at midnight UTC should be serial number 45292
      sheet.cell('A1').value = new Date('2024-01-01T00:00:00.000Z');

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      // Date at midnight UTC should give exact integer
      expect(sheetXml).toContain('<v>45292</v>');
    });
  });

  describe('Multiple Sheets', () => {
    it('should export multiple sheets', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1').cell('A1').value = 'First';
      wb.addSheet('Sheet2').cell('A1').value = 'Second';
      wb.addSheet('Sheet3').cell('A1').value = 'Third';

      const xlsx = workbookToXlsx(wb);

      expect(hasFile(xlsx, 'xl/worksheets/sheet1.xml')).toBe(true);
      expect(hasFile(xlsx, 'xl/worksheets/sheet2.xml')).toBe(true);
      expect(hasFile(xlsx, 'xl/worksheets/sheet3.xml')).toBe(true);

      const workbookXml = getXmlFile(xlsx, 'xl/workbook.xml');
      expect(workbookXml).toContain('name="Sheet1"');
      expect(workbookXml).toContain('name="Sheet2"');
      expect(workbookXml).toContain('name="Sheet3"');
    });
  });

  describe('Shared Strings', () => {
    it('should deduplicate repeated strings', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Hello';
      sheet.cell('A2').value = 'Hello'; // Duplicate
      sheet.cell('A3').value = 'World';

      const xlsx = workbookToXlsx(wb);
      const sst = getXmlFile(xlsx, 'xl/sharedStrings.xml');

      // Should have count="3" (total references) and uniqueCount="2"
      expect(sst).toContain('count="3"');
      expect(sst).toContain('uniqueCount="2"');
    });

    it('should preserve whitespace in strings', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = '  leading spaces';
      sheet.cell('A2').value = 'trailing spaces  ';

      const xlsx = workbookToXlsx(wb);
      const sst = getXmlFile(xlsx, 'xl/sharedStrings.xml');

      expect(sst).toContain('xml:space="preserve"');
    });
  });

  describe('Styles', () => {
    it('should export bold font', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Bold Text';
      sheet.cell('A1').style = { font: { bold: true } };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('<b/>');
    });

    it('should export italic font', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Italic Text';
      sheet.cell('A1').style = { font: { italic: true } };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('<i/>');
    });

    it('should export font color', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Red Text';
      sheet.cell('A1').style = { font: { color: '#FF0000' } };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('FFFF0000'); // ARGB format
    });

    it('should export fill color', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Yellow Background';
      sheet.cell('A1').style = {
        fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FFFF00' },
      };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('FFFFFF00'); // Yellow in ARGB
    });

    it('should export borders', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Bordered';
      sheet.cell('A1').style = {
        borders: {
          top: { style: 'thin', color: '#000000' },
          bottom: { style: 'thin', color: '#000000' },
          left: { style: 'thin', color: '#000000' },
          right: { style: 'thin', color: '#000000' },
        },
      };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('style="thin"');
    });

    it('should export number format', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 1234.56;
      sheet.cell('A1').style = {
        numberFormat: { formatCode: '#,##0.00' },
      };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      // Built-in format #,##0.00 has ID 4
      expect(styles).toContain('numFmtId="4"');
    });

    it('should export alignment', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Centered';
      sheet.cell('A1').style = {
        alignment: { horizontal: 'center', vertical: 'middle' },
      };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('horizontal="center"');
      expect(styles).toContain('vertical="middle"');
    });

    it('should export text wrap', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Wrapped\nText';
      sheet.cell('A1').style = {
        alignment: { wrapText: true },
      };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('wrapText="1"');
    });
  });

  describe('Formulas', () => {
    it('should export formulas', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 10;
      sheet.cell('A2').value = 20;
      sheet.cell('A3').setFormula('SUM(A1:A2)');

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('<f>SUM(A1:A2)</f>');
    });
  });

  describe('Merged Cells', () => {
    it('should export merge ranges', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Merged';
      sheet.mergeCells('A1:C1');

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('<mergeCells');
      expect(sheetXml).toContain('ref="A1:C1"');
    });

    it('should handle multiple merges', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Merge 1';
      sheet.mergeCells('A1:B1');
      sheet.cell('A2').value = 'Merge 2';
      sheet.mergeCells('A2:C2');

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('count="2"');
      expect(sheetXml).toContain('ref="A1:B1"');
      expect(sheetXml).toContain('ref="A2:C2"');
    });
  });

  describe('Column and Row Dimensions', () => {
    it('should export column widths', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.setColumnWidth(0, 20);
      sheet.cell('A1').value = 'Wide column';

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('<cols>');
      expect(sheetXml).toContain('customWidth="1"');
    });

    it('should export row heights', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.setRowHeight(0, 30);
      sheet.cell('A1').value = 'Tall row';

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('ht="30"');
      expect(sheetXml).toContain('customHeight="1"');
    });

    it('should export hidden columns', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.hideColumn(1); // Hide column B (0-based index)
      sheet.cell('A1').value = 'A';
      sheet.cell('B1').value = 'Hidden';

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('hidden="1"');
    });
  });

  describe('Freeze Panes', () => {
    it('should export frozen rows', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.freeze(1, 0);
      sheet.cell('A1').value = 'Header';

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('state="frozen"');
      expect(sheetXml).toContain('ySplit="1"');
    });

    it('should export frozen columns', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.freeze(0, 1);
      sheet.cell('A1').value = 'Fixed';

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('state="frozen"');
      expect(sheetXml).toContain('xSplit="1"');
    });

    it('should export both frozen rows and columns', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.freeze(2, 2);
      sheet.cell('A1').value = 'Corner';

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('xSplit="2"');
      expect(sheetXml).toContain('ySplit="2"');
      expect(sheetXml).toContain('activePane="bottomRight"');
    });
  });

  describe('Content Types', () => {
    it('should declare all required content types', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1');

      const xlsx = workbookToXlsx(wb);
      const contentTypes = getXmlFile(xlsx, '[Content_Types].xml');

      expect(contentTypes).toContain('spreadsheetml.sheet.main');
      expect(contentTypes).toContain('spreadsheetml.worksheet');
      expect(contentTypes).toContain('spreadsheetml.styles');
    });
  });

  describe('Relationships', () => {
    it('should create valid workbook relationships', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1');
      wb.addSheet('Sheet2');

      const xlsx = workbookToXlsx(wb);
      const rels = getXmlFile(xlsx, 'xl/_rels/workbook.xml.rels');

      expect(rels).toContain('Target="worksheets/sheet1.xml"');
      expect(rels).toContain('Target="worksheets/sheet2.xml"');
      expect(rels).toContain('Target="styles.xml"');
    });
  });

  describe('XML Escaping', () => {
    it('should escape special characters in strings', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = '<script>alert("XSS")</script>';
      sheet.cell('A2').value = 'Tom & Jerry';
      sheet.cell('A3').value = "It's a test";

      const xlsx = workbookToXlsx(wb);
      const sst = getXmlFile(xlsx, 'xl/sharedStrings.xml');

      expect(sst).toContain('&lt;script&gt;');
      expect(sst).toContain('&amp;');
      expect(sst).toContain('&apos;');
      expect(sst).not.toContain('<script>');
    });

    it('should escape special characters in sheet names', () => {
      const wb = new Workbook();
      wb.addSheet('Data & Info');

      const xlsx = workbookToXlsx(wb);
      const workbookXml = getXmlFile(xlsx, 'xl/workbook.xml');

      expect(workbookXml).toContain('Data &amp; Info');
    });
  });

  describe('Document Properties', () => {
    it('should include document properties when enabled', () => {
      const wb = new Workbook();
      wb.properties.title = 'Test Workbook';
      wb.properties.author = 'Test Author';
      wb.addSheet('Sheet1');

      const xlsx = workbookToXlsx(wb, { includeProperties: true });

      expect(hasFile(xlsx, 'docProps/core.xml')).toBe(true);
      expect(hasFile(xlsx, 'docProps/app.xml')).toBe(true);

      const core = getXmlFile(xlsx, 'docProps/core.xml');
      expect(core).toContain('Test Workbook');
      expect(core).toContain('Test Author');
    });

    it('should exclude document properties when disabled', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1');

      const xlsx = workbookToXlsx(wb, { includeProperties: false });

      expect(hasFile(xlsx, 'docProps/core.xml')).toBe(false);
      expect(hasFile(xlsx, 'docProps/app.xml')).toBe(false);
    });
  });

  describe('Edge Cases', () => {
    it('should handle empty cells with styles', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').style = { fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FFFF00' } };

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      // Cell should have style but no value
      expect(sheetXml).toContain('<c r="A1"');
      expect(sheetXml).toContain('s="');
    });

    it('should handle NaN and Infinity as errors', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = NaN;
      sheet.cell('B1').value = Infinity;

      const xlsx = workbookToXlsx(wb);
      const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

      expect(sheetXml).toContain('t="e"'); // Error type
      expect(sheetXml).toContain('#NUM!');
    });

    it('should handle very long strings', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      const longString = 'A'.repeat(10000);
      sheet.cell('A1').value = longString;

      const xlsx = workbookToXlsx(wb);
      const sst = getXmlFile(xlsx, 'xl/sharedStrings.xml');

      expect(sst).toContain('A'.repeat(100)); // Contains part of the string
    });
  });
});

describe('XLSX Date Conversion', () => {
  it('should correctly convert dates after 1900 leap year bug cutoff', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    // March 1, 1900 at midnight UTC should be serial 61 (accounting for leap year bug)
    // Days from Dec 31, 1899: Jan=31 + Feb=28 + 1 = 60, then +1 for bug = 61
    sheet.cell('A1').value = new Date('1900-03-01T00:00:00.000Z');

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('<v>61</v>');
  });

  it('should correctly convert modern dates', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    // December 15, 2025 at midnight UTC = serial 46006
    sheet.cell('A1').value = new Date('2025-12-15T00:00:00.000Z');

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    // Exact integer for midnight UTC date
    expect(sheetXml).toContain('<v>46006</v>');
  });

  it('should handle dates before leap year bug cutoff', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    // February 28, 1900 should be serial 59 (before the fake Feb 29)
    sheet.cell('A1').value = new Date('1900-02-28T00:00:00.000Z');

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('<v>59</v>');
  });
});

describe('XLSX XML Utilities', () => {
  it('should escape XML special characters in sheet names', () => {
    const wb = new Workbook();
    // Sheet name with special chars will be escaped in workbook.xml
    wb.addSheet('Test & <Sheet>');

    const xlsx = workbookToXlsx(wb);
    const workbookXml = getXmlFile(xlsx, 'xl/workbook.xml');

    expect(workbookXml).toContain('&amp;');
    expect(workbookXml).toContain('&lt;');
    expect(workbookXml).toContain('&gt;');
  });

  it('should escape special characters in shared strings', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    // String with special chars goes to shared strings
    sheet.cell('A1').value = 'Test & <value>';

    const xlsx = workbookToXlsx(wb);
    const stringsXml = getXmlFile(xlsx, 'xl/sharedStrings.xml');

    // Should be escaped in shared strings table
    expect(stringsXml).toContain('&amp;');
    expect(stringsXml).toContain('&lt;');
    expect(stringsXml).toContain('&gt;');
  });

  it('should sanitize control characters from shared strings', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    // String with control characters (NULL, BEL)
    sheet.cell('A1').value = 'Clean\x00text\x07here';

    const xlsx = workbookToXlsx(wb);
    const stringsXml = getXmlFile(xlsx, 'xl/sharedStrings.xml');

    // Control chars should be removed in shared strings
    expect(stringsXml).not.toContain('\x00');
    expect(stringsXml).not.toContain('\x07');
    expect(stringsXml).toContain('Cleantexthere');
  });

  it('should handle Excel error values', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = '#DIV/0!';
    sheet.cell('A2').value = '#VALUE!';
    sheet.cell('A3').value = '#REF!';
    sheet.cell('A4').value = '#NAME?';
    sheet.cell('A5').value = '#NUM!';
    sheet.cell('A6').value = '#N/A';

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    // Errors should be marked with t="e" (error type)
    expect(sheetXml).toContain('t="e"');
    expect(sheetXml).toContain('#DIV/0!');
  });
});

describe('XLSX Style Coverage', () => {
  it('should handle shorthand colors (#RGB)', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Red';
    sheet.cell('A1').applyStyle({
      font: { color: '#F00' }, // Shorthand red
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    // Should expand to FFFF0000
    expect(stylesXml).toContain('FF0000');
  });

  it('should handle italic font style', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Italic';
    sheet.cell('A1').applyStyle({
      font: { italic: true },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('<i/>');
  });

  it('should handle underline font style', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Underlined';
    sheet.cell('A1').applyStyle({
      font: { underline: 'single' },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('<u');
  });

  it('should handle strikethrough font style', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Strikethrough';
    sheet.cell('A1').applyStyle({
      font: { strikethrough: true },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('<strike/>');
  });

  it('should handle text wrap alignment', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Long text that wraps';
    sheet.cell('A1').applyStyle({
      alignment: { wrapText: true },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('wrapText="1"');
  });

  it('should handle text rotation', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Rotated';
    sheet.cell('A1').applyStyle({
      alignment: { textRotation: 45 },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('textRotation="45"');
  });

  it('should handle indent level', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Indented';
    sheet.cell('A1').applyStyle({
      alignment: { indent: 2 },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('indent="2"');
  });

  it('should handle all border sides', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Bordered';
    sheet.cell('A1').applyStyle({
      borders: {
        top: { style: 'thin', color: '#000000' },
        bottom: { style: 'thin', color: '#000000' },
        left: { style: 'thin', color: '#000000' },
        right: { style: 'thin', color: '#000000' },
      },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('<left');
    expect(stylesXml).toContain('<right');
    expect(stylesXml).toContain('<top');
    expect(stylesXml).toContain('<bottom');
  });

  it('should handle pattern fills with background color', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Pattern fill';
    sheet.cell('A1').applyStyle({
      fill: {
        type: 'pattern',
        pattern: 'solid',
        foregroundColor: '#FFFF00',
        backgroundColor: '#000000',
      },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('<patternFill');
    expect(stylesXml).toContain('FFFFFF00'); // Yellow with alpha
  });

  it('should handle custom number formats', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 1234.567;
    sheet.cell('A1').applyStyle({
      // Use truly custom format not in builtin list
      numberFormat: { formatCode: '$#,##0.000_);[Red]($#,##0.000)' },
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    expect(stylesXml).toContain('<numFmt');
    expect(stylesXml).toContain('$#,##0.000');
  });

  it('should handle builtin number format codes', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 0.5;
    sheet.cell('A1').applyStyle({
      numberFormat: { formatCode: '0%' }, // Built-in format code
    });

    const xlsx = workbookToXlsx(wb);
    const stylesXml = getXmlFile(xlsx, 'xl/styles.xml');

    // Built-in format (id 9) should use numFmtId directly
    expect(stylesXml).toContain('numFmtId="9"');
  });
});

describe('XLSX Parts Coverage', () => {
  it('should handle hidden rows', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.hideRow(0); // Hide row 1
    sheet.cell('A1').value = 'Hidden row';
    sheet.cell('A2').value = 'Visible row';

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('hidden="1"');
  });

  it('should handle sheet views with grid lines disabled', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'No grid';
    sheet.view.showGridLines = false;

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('showGridLines="0"');
  });

  it('should handle sheet views with zoom scale', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 'Zoomed';
    sheet.view.zoomScale = 150;

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('zoomScale="150"');
  });

  it('should handle frozen columns only', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.freeze(0, 2); // Freeze 2 columns, no rows
    sheet.cell('A1').value = 'Frozen columns';

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('xSplit="2"');
    expect(sheetXml).toContain('ySplit="0"');
    expect(sheetXml).toContain('activePane="topRight"');
  });

  it('should export document properties when enabled', () => {
    const wb = new Workbook();
    wb.properties.title = 'My Workbook';
    wb.properties.author = 'Test Author';
    wb.properties.lastModifiedBy = 'Editor';
    wb.properties.created = new Date('2024-01-01T00:00:00.000Z');
    wb.properties.modified = new Date('2024-06-01T00:00:00.000Z');
    wb.addSheet('Test').cell('A1').value = 'Test';

    const xlsx = workbookToXlsx(wb, { includeProperties: true });
    const coreXml = getXmlFile(xlsx, 'docProps/core.xml');
    const appXml = getXmlFile(xlsx, 'docProps/app.xml');

    expect(coreXml).toContain('<dc:title>My Workbook</dc:title>');
    expect(coreXml).toContain('<dc:creator>Test Author</dc:creator>');
    expect(coreXml).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(coreXml).toContain('<dcterms:created');
    expect(coreXml).toContain('<dcterms:modified');
    expect(appXml).toContain('<Application>');
  });

  it('should handle formula cells', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = 10;
    sheet.cell('A2').value = 20;
    sheet.cell('A3').setFormula('SUM(A1:A2)');

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('<f>SUM(A1:A2)</f>');
  });

  it('should handle empty cells with borders', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    // Empty cell with style
    sheet.cell('A1').applyStyle({
      borders: {
        left: { style: 'thin', color: '#000000' },
      },
    });

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    // Empty cell should still be exported with style
    expect(sheetXml).toContain('<c r="A1"');
    expect(sheetXml).toContain('s="');
  });
});

describe('XLSX Writer Edge Cases', () => {
  it('should handle workbook with empty sheet names', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet 1'); // Name with space
    wb.addSheet('Sheet&2'); // Name with ampersand

    const xlsx = workbookToXlsx(wb);
    const workbookXml = getXmlFile(xlsx, 'xl/workbook.xml');

    expect(workbookXml).toContain('Sheet 1');
    expect(workbookXml).toContain('Sheet&amp;2');
  });

  it('should handle compression level option', () => {
    const wb = new Workbook();
    wb.addSheet('Test').cell('A1').value = 'Test';

    // Different compression levels should produce valid output
    const xlsx0 = workbookToXlsx(wb, { compressionLevel: 0 });
    const xlsx9 = workbookToXlsx(wb, { compressionLevel: 9 });

    // Both should be valid ZIPs
    expect(xlsx0.length).toBeGreaterThan(0);
    expect(xlsx9.length).toBeGreaterThan(0);
    // Higher compression should be smaller (usually)
    expect(xlsx9.length).toBeLessThanOrEqual(xlsx0.length);
  });

  it('should handle Infinity values', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('Test');
    sheet.cell('A1').value = Infinity;
    sheet.cell('A2').value = -Infinity;

    const xlsx = workbookToXlsx(wb);
    const sheetXml = getXmlFile(xlsx, 'xl/worksheets/sheet1.xml');

    // Should be exported as #NUM! errors
    expect(sheetXml).toContain('t="e"');
    expect(sheetXml).toContain('#NUM!');
  });
});

describe('XLSX Import', () => {
  describe('Basic Import', () => {
    it('should import empty workbook', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1');

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx);

      expect(workbook.sheetCount).toBe(1);
      expect(workbook.sheets[0].name).toBe('Sheet1');
      expect(stats.sheetCount).toBe(1);
    });

    it('should import string values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Hello';
      sheet.cell('B1').value = 'World';

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').value).toBe('Hello');
      expect(imported?.cell('B1').value).toBe('World');
      expect(stats.totalCells).toBe(2);
    });

    it('should import number values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 42;
      sheet.cell('B1').value = 3.14159;
      sheet.cell('C1').value = -100;

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').value).toBe(42);
      expect(imported?.cell('B1').value).toBeCloseTo(3.14159);
      expect(imported?.cell('C1').value).toBe(-100);
    });

    it('should import boolean values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = true;
      sheet.cell('B1').value = false;

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').value).toBe(true);
      expect(imported?.cell('B1').value).toBe(false);
    });

    it('should import date values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      const testDate = new Date('2024-01-15T00:00:00.000Z');
      sheet.cell('A1').value = testDate;
      sheet.cell('A1').style = { numberFormat: { formatCode: 'yyyy-mm-dd' } };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      const value = imported?.cell('A1').value;
      expect(value).toBeInstanceOf(Date);
      if (value instanceof Date) {
        expect(value.toISOString().split('T')[0]).toBe('2024-01-15');
      }
    });
  });

  describe('Multiple Sheets', () => {
    it('should import multiple sheets', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1').cell('A1').value = 'First';
      wb.addSheet('Sheet2').cell('A1').value = 'Second';
      wb.addSheet('Sheet3').cell('A1').value = 'Third';

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx);

      expect(workbook.sheetCount).toBe(3);
      expect(stats.sheetCount).toBe(3);
      expect(workbook.getSheet('Sheet1')?.cell('A1').value).toBe('First');
      expect(workbook.getSheet('Sheet2')?.cell('A1').value).toBe('Second');
      expect(workbook.getSheet('Sheet3')?.cell('A1').value).toBe('Third');
    });

    it('should import specific sheets by name', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1').cell('A1').value = 'First';
      wb.addSheet('Sheet2').cell('A1').value = 'Second';
      wb.addSheet('Sheet3').cell('A1').value = 'Third';

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx, { sheets: ['Sheet2'] });

      expect(workbook.sheetCount).toBe(1);
      expect(stats.sheetCount).toBe(1);
      expect(workbook.getSheet('Sheet2')?.cell('A1').value).toBe('Second');
    });

    it('should import specific sheets by index', () => {
      const wb = new Workbook();
      wb.addSheet('Sheet1').cell('A1').value = 'First';
      wb.addSheet('Sheet2').cell('A1').value = 'Second';
      wb.addSheet('Sheet3').cell('A1').value = 'Third';

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx, { sheets: [0, 2] });

      expect(workbook.sheetCount).toBe(2);
      expect(stats.sheetCount).toBe(2);
      expect(workbook.sheets[0].cell('A1').value).toBe('First');
      expect(workbook.sheets[1].cell('A1').value).toBe('Third');
    });
  });

  describe('Formulas', () => {
    it('should import formulas', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 10;
      sheet.cell('A2').value = 20;
      sheet.cell('A3').setFormula('SUM(A1:A2)');

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A3').formula?.formula).toBe('SUM(A1:A2)');
      expect(stats.formulaCells).toBe(1);
    });

    it('should respect importFormulas option', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').setFormula('1+1');

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx, { importFormulas: false });

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').formula).toBeUndefined();
      expect(stats.formulaCells).toBe(0);
    });
  });

  describe('Styles', () => {
    it('should import bold font', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Bold';
      sheet.cell('A1').style = { font: { bold: true } };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').style?.font?.bold).toBe(true);
    });

    it('should import italic font', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Italic';
      sheet.cell('A1').style = { font: { italic: true } };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').style?.font?.italic).toBe(true);
    });

    it('should import background color', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Colored';
      sheet.cell('A1').style = {
        fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FF0000' },
      };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').style?.fill?.foregroundColor).toBe('#FF0000');
    });

    it('should respect importStyles option', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Styled';
      sheet.cell('A1').style = { font: { bold: true } };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx, { importStyles: false });

      const imported = workbook.getSheet('Test');
      // Value should be there, but style should be empty/undefined
      expect(imported?.cell('A1').value).toBe('Styled');
    });
  });

  describe('Merged Cells', () => {
    it('should import merged cells', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Merged';
      sheet.mergeCells('A1:C3');

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').value).toBe('Merged');
      expect(imported?.cell('A1').isMergeMaster).toBe(true);
      expect(imported?.cell('B2').isMergedSlave).toBe(true);
      expect(stats.mergedRanges).toBe(1);
    });

    it('should respect importMergedCells option', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Merged';
      sheet.mergeCells('A1:C3');

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx, { importMergedCells: false });

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').isMergeMaster).toBeFalsy();
      expect(stats.mergedRanges).toBe(0);
    });
  });

  describe('Dimensions', () => {
    it('should import column widths', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Wide column';
      sheet.setColumnWidth(0, 20);

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      const colConfig = imported?.getColumn(0);
      // Excel column width conversion adds some offset, allow tolerance
      expect(colConfig?.width).toBeGreaterThan(19);
      expect(colConfig?.width).toBeLessThan(22);
    });

    it('should import row heights', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Tall row';
      sheet.setRowHeight(0, 30);

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      const rowConfig = imported?.getRow(0);
      expect(rowConfig?.height).toBeCloseTo(30, 0);
    });
  });

  describe('Freeze Panes', () => {
    it('should import frozen rows', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Header';
      sheet.freeze(1);

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.view.frozenRows).toBe(1);
    });

    it('should import frozen columns', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Label';
      sheet.freeze(0, 1);

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.view.frozenCols).toBe(1);
    });

    it('should import frozen rows and columns', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Corner';
      sheet.freeze(2, 3);

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.view.frozenRows).toBe(2);
      expect(imported?.view.frozenCols).toBe(3);
    });

    it('should respect importFreezePanes option', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.freeze(1);

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx, { importFreezePanes: false });

      const imported = workbook.getSheet('Test');
      expect(imported?.view.frozenRows).toBeUndefined();
    });
  });

  describe('Import Statistics', () => {
    it('should report accurate statistics', () => {
      const wb = new Workbook();
      const sheet1 = wb.addSheet('Sheet1');
      sheet1.cell('A1').value = 'Text';
      sheet1.cell('A2').value = 42;
      sheet1.cell('A3').setFormula('A2*2');
      sheet1.mergeCells('B1:C2');

      const sheet2 = wb.addSheet('Sheet2');
      sheet2.cell('A1').value = 'More data';

      const xlsx = workbookToXlsx(wb);
      const { stats } = xlsxToWorkbook(xlsx);

      expect(stats.sheetCount).toBe(2);
      expect(stats.totalCells).toBe(3); // A1, A2 in Sheet1 (A3 formula has no cached value), A1 in Sheet2
      expect(stats.formulaCells).toBe(1);
      expect(stats.mergedRanges).toBe(1);
      expect(stats.durationMs).toBeGreaterThanOrEqual(0);
    });
  });

  describe('Edge Cases', () => {
    it('should handle empty strings', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = '';

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').value).toBe('');
    });

    it('should handle special characters', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = '<>&"\'';
      sheet.cell('A2').value = '日本語テスト';

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').value).toBe('<>&"\'');
      expect(imported?.cell('A2').value).toBe('日本語テスト');
    });

    it('should handle large numbers', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 1234567890123456;
      sheet.cell('A2').value = 0.000000001;

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.cell('A1').value).toBeCloseTo(1234567890123456);
      expect(imported?.cell('A2').value).toBeCloseTo(0.000000001);
    });

    it('should handle maxRows option', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      for (let i = 0; i < 10; i++) {
        sheet.cell(i, 0).value = `Row ${i}`;
      }

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx, { maxRows: 5 });

      const imported = workbook.getSheet('Test');
      expect(imported?.cell(4, 0).value).toBe('Row 4');
      expect(imported?.cell(5, 0).value).toBeNull();
      expect(stats.totalCells).toBe(5);
    });

    it('should handle maxCols option', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      for (let i = 0; i < 10; i++) {
        sheet.cell(0, i).value = `Col ${i}`;
      }

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx, { maxCols: 5 });

      const imported = workbook.getSheet('Test');
      expect(imported?.cell(0, 4).value).toBe('Col 4');
      expect(imported?.cell(0, 5).value).toBeNull();
      expect(stats.totalCells).toBe(5);
    });
  });

  describe('Round Trip', () => {
    it('should preserve data through export and import', () => {
      const wb = new Workbook();
      wb.properties.title = 'Test Workbook';
      wb.properties.author = 'Test Author';

      const sheet = wb.addSheet('Data');
      sheet.cell('A1').value = 'Name';
      sheet.cell('B1').value = 'Value';
      sheet.cell('A1').style = { font: { bold: true } };
      sheet.cell('B1').style = { font: { bold: true } };
      sheet.cell('A2').value = 'Item 1';
      sheet.cell('B2').value = 100;
      sheet.cell('A3').value = 'Item 2';
      sheet.cell('B3').value = 200;
      sheet.cell('B4').setFormula('SUM(B2:B3)');
      sheet.mergeCells('A1:B1');
      sheet.setColumnWidth(0, 15);
      sheet.setColumnWidth(1, 10);
      sheet.freeze(1);

      const xlsx = workbookToXlsx(wb);
      const { workbook, stats } = xlsxToWorkbook(xlsx);

      // Verify structure
      expect(workbook.sheetCount).toBe(1);
      expect(stats.sheetCount).toBe(1);

      const imported = workbook.getSheet('Data');
      expect(imported).toBeDefined();

      // Verify values
      expect(imported?.cell('A2').value).toBe('Item 1');
      expect(imported?.cell('B2').value).toBe(100);
      expect(imported?.cell('A3').value).toBe('Item 2');
      expect(imported?.cell('B3').value).toBe(200);

      // Verify formula
      expect(imported?.cell('B4').formula?.formula).toBe('SUM(B2:B3)');

      // Verify styles (bold font)
      expect(imported?.cell('A1').style?.font?.bold).toBe(true);

      // Verify merged cells
      expect(imported?.cell('A1').isMergeMaster).toBe(true);

      // Verify freeze pane
      expect(imported?.view.frozenRows).toBe(1);

      // Verify properties
      expect(workbook.properties.title).toBe('Test Workbook');
      expect(workbook.properties.author).toBe('Test Author');
    });
  });

  describe('Comments', () => {
    it('should export cell comments', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Sheet1');
      const cell1 = sheet.cell('A1');
      cell1.value = 'Value';
      cell1.setComment('This is a comment');
      const cell2 = sheet.cell('A2');
      cell2.value = 'Another';
      cell2.setComment('Another comment', 'Author Name');

      const xlsx = workbookToXlsx(wb);

      // Check that comments file exists
      expect(hasFile(xlsx, 'xl/comments1.xml')).toBe(true);

      // Check comments content
      const commentsXml = getXmlFile(xlsx, 'xl/comments1.xml');
      expect(commentsXml).toContain('<comments');
      expect(commentsXml).toContain('This is a comment');
      expect(commentsXml).toContain('Another comment');
      expect(commentsXml).toContain('Author Name');
      expect(commentsXml).toContain('ref="A1"');
      expect(commentsXml).toContain('ref="A2"');
    });

    it('should not create comments file when no comments exist', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Sheet1');
      sheet.cell('A1').value = 'No comment here';

      const xlsx = workbookToXlsx(wb);

      expect(hasFile(xlsx, 'xl/comments1.xml')).toBe(false);
    });

    it('should export comments with default author', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Sheet1');
      const cell = sheet.cell('A1');
      cell.value = 'Value';
      cell.setComment('Comment without author');

      const xlsx = workbookToXlsx(wb);
      const commentsXml = getXmlFile(xlsx, 'xl/comments1.xml');

      // Should have authors section with at least empty default author
      expect(commentsXml).toContain('<authors>');
      expect(commentsXml).toContain('</authors>');
      expect(commentsXml).toContain('authorId="0"');
    });

    it('should export comments with multiple authors', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Sheet1');
      const cellA = sheet.cell('A1');
      cellA.value = 'A';
      cellA.setComment('Comment 1', 'Alice');
      const cellB = sheet.cell('A2');
      cellB.value = 'B';
      cellB.setComment('Comment 2', 'Bob');
      const cellC = sheet.cell('A3');
      cellC.value = 'C';
      cellC.setComment('Comment 3', 'Alice'); // Same author

      const xlsx = workbookToXlsx(wb);
      const commentsXml = getXmlFile(xlsx, 'xl/comments1.xml');

      expect(commentsXml).toContain('<author>Alice</author>');
      expect(commentsXml).toContain('<author>Bob</author>');
    });

    it('should round-trip comments', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Comments');
      const cell1 = sheet.cell('A1');
      cell1.value = 'Value 1';
      cell1.setComment('First comment');
      const cell2 = sheet.cell('B2');
      cell2.value = 'Value 2';
      cell2.setComment('Second comment', 'Test Author');
      const cell3 = sheet.cell('A3');
      cell3.value = 'Value 3';
      cell3.setComment('Third comment with special chars: <>&"');

      const xlsx = workbookToXlsx(wb);
      const result = xlsxToWorkbook(xlsx);
      const imported = result.workbook.sheets[0];

      // Check comments were imported
      const importedCell1 = imported.getCell(0, 0);
      expect(importedCell1?.comment).toBeDefined();
      expect(importedCell1?.comment?.text).toBe('First comment');

      const importedCell2 = imported.getCell(1, 1);
      expect(importedCell2?.comment).toBeDefined();
      expect(importedCell2?.comment?.text).toBe('Second comment');
      expect(importedCell2?.comment?.author).toBe('Test Author');

      const importedCell3 = imported.getCell(2, 0);
      expect(importedCell3?.comment).toBeDefined();
      expect(importedCell3?.comment?.text).toBe('Third comment with special chars: <>&"');
    });

    it('should handle comments on multiple sheets', () => {
      const wb = new Workbook();
      const sheet1 = wb.addSheet('Sheet1');
      const sheet2 = wb.addSheet('Sheet2');

      const s1Cell = sheet1.cell('A1');
      s1Cell.value = 'S1';
      s1Cell.setComment('Sheet 1 comment');
      const s2Cell = sheet2.cell('A1');
      s2Cell.value = 'S2';
      s2Cell.setComment('Sheet 2 comment');

      const xlsx = workbookToXlsx(wb);

      expect(hasFile(xlsx, 'xl/comments1.xml')).toBe(true);
      expect(hasFile(xlsx, 'xl/comments2.xml')).toBe(true);

      const comments1 = getXmlFile(xlsx, 'xl/comments1.xml');
      const comments2 = getXmlFile(xlsx, 'xl/comments2.xml');

      expect(comments1).toContain('Sheet 1 comment');
      expect(comments2).toContain('Sheet 2 comment');
    });

    it('should import comments with importComments option disabled', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Sheet1');
      const cell = sheet.cell('A1');
      cell.value = 'Value';
      cell.setComment('This comment should not be imported');

      const xlsx = workbookToXlsx(wb);
      const result = xlsxToWorkbook(xlsx, { importComments: false });
      const imported = result.workbook.sheets[0];

      // Comment should not be imported
      const importedCell = imported.getCell(0, 0);
      expect(importedCell?.comment).toBeUndefined();
    });

    it('should create worksheet rels for comments', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Sheet1');
      const cell = sheet.cell('A1');
      cell.value = 'Value';
      cell.setComment('Comment');

      const xlsx = workbookToXlsx(wb);

      // Check worksheet rels file exists
      expect(hasFile(xlsx, 'xl/worksheets/_rels/sheet1.xml.rels')).toBe(true);

      const relsXml = getXmlFile(xlsx, 'xl/worksheets/_rels/sheet1.xml.rels');
      expect(relsXml).toContain('comments');
    });

    it('should include comments content type', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Sheet1');
      const cell = sheet.cell('A1');
      cell.value = 'Value';
      cell.setComment('Comment');

      const xlsx = workbookToXlsx(wb);
      const contentTypes = getXmlFile(xlsx, '[Content_Types].xml');

      expect(contentTypes).toContain('comments');
    });
  });
});

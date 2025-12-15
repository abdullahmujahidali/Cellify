import { describe, it, expect } from 'vitest';
import { unzipSync } from 'fflate';
import { Workbook } from '../src/core/Workbook.js';
import { workbookToXlsx } from '../src/formats/xlsx/index.js';

/**
 * Helper to extract and decode a file from the XLSX ZIP
 */
function getXmlFile(xlsx: Uint8Array, path: string): string {
  const files = unzipSync(xlsx);
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
  const files = unzipSync(xlsx);
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
        alignment: { horizontal: 'center', vertical: 'center' },
      };

      const xlsx = workbookToXlsx(wb);
      const styles = getXmlFile(xlsx, 'xl/styles.xml');

      expect(styles).toContain('horizontal="center"');
      expect(styles).toContain('vertical="center"');
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
});

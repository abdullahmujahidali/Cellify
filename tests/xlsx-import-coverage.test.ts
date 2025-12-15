/**
 * Additional tests for XLSX import to improve code coverage
 * Tests edge cases, error handling, and less common code paths
 */

import { describe, it, expect, vi } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Workbook } from '../src/core/Workbook.js';
import { workbookToXlsx, xlsxToWorkbook, xlsxBlobToWorkbook } from '../src/formats/xlsx/index.js';
import {
  unescapeXml,
  parseCellRef,
  parseRangeRef,
  resolveRelPath,
  parseElement,
  parseElements,
  getAttr,
  getTextContent,
} from '../src/formats/xlsx/xlsx.parser.js';
import { unescapeXml as unescapeXmlFromXml } from '../src/formats/xlsx/xlsx.xml.js';

describe('xlsx.xml.ts Coverage', () => {
  describe('unescapeXml', () => {
    it('should unescape all XML entities', () => {
      expect(unescapeXmlFromXml('&lt;')).toBe('<');
      expect(unescapeXmlFromXml('&gt;')).toBe('>');
      expect(unescapeXmlFromXml('&quot;')).toBe('"');
      expect(unescapeXmlFromXml('&apos;')).toBe("'");
      expect(unescapeXmlFromXml('&amp;')).toBe('&');
    });

    it('should handle mixed entities', () => {
      expect(unescapeXmlFromXml('&lt;tag&gt;')).toBe('<tag>');
      expect(unescapeXmlFromXml('&quot;hello&quot;')).toBe('"hello"');
    });

    it('should handle ampersand last to avoid double-unescaping', () => {
      // &amp;lt; should become &lt; not <
      expect(unescapeXmlFromXml('&amp;lt;')).toBe('&lt;');
    });
  });
});

describe('xlsx.parser.ts Coverage', () => {
  describe('parseCellRef', () => {
    it('should parse simple cell references', () => {
      expect(parseCellRef('A1')).toEqual({ row: 0, col: 0 });
      expect(parseCellRef('B2')).toEqual({ row: 1, col: 1 });
      expect(parseCellRef('Z26')).toEqual({ row: 25, col: 25 });
    });

    it('should parse multi-letter column references', () => {
      expect(parseCellRef('AA1')).toEqual({ row: 0, col: 26 });
      expect(parseCellRef('AB1')).toEqual({ row: 0, col: 27 });
      expect(parseCellRef('AZ1')).toEqual({ row: 0, col: 51 });
    });

    it('should throw on invalid cell reference', () => {
      expect(() => parseCellRef('invalid')).toThrow('Invalid cell reference');
      expect(() => parseCellRef('123')).toThrow('Invalid cell reference');
      expect(() => parseCellRef('')).toThrow('Invalid cell reference');
    });
  });

  describe('parseRangeRef', () => {
    it('should parse range references', () => {
      expect(parseRangeRef('A1:C3')).toEqual({
        startRow: 0,
        startCol: 0,
        endRow: 2,
        endCol: 2,
      });
    });

    it('should parse single cell as range', () => {
      expect(parseRangeRef('B2')).toEqual({
        startRow: 1,
        startCol: 1,
        endRow: 1,
        endCol: 1,
      });
    });
  });

  describe('resolveRelPath', () => {
    it('should handle relative paths', () => {
      expect(resolveRelPath('xl/', 'worksheets/sheet1.xml')).toBe('xl/worksheets/sheet1.xml');
    });

    it('should handle absolute paths', () => {
      expect(resolveRelPath('xl/', '/docProps/core.xml')).toBe('docProps/core.xml');
    });

    it('should handle parent directory navigation', () => {
      expect(resolveRelPath('xl/worksheets/', '../sharedStrings.xml')).toBe('xl/sharedStrings.xml');
    });

    it('should handle current directory', () => {
      expect(resolveRelPath('xl/', './styles.xml')).toBe('xl/styles.xml');
    });
  });

  describe('parseElement', () => {
    it('should parse self-closing elements', () => {
      const el = parseElement('<item attr="val"/>', 'item');
      expect(el).toBeDefined();
      expect(el?.attrs.attr).toBe('val');
      expect(el?.inner).toBe('');
    });

    it('should parse namespaced elements', () => {
      const el = parseElement('<x:sheet name="Test"/>', 'sheet');
      expect(el).toBeDefined();
      expect(el?.attrs.name).toBe('Test');
    });

    it('should return undefined for missing element', () => {
      expect(parseElement('<other/>', 'missing')).toBeUndefined();
    });
  });

  describe('parseElements', () => {
    it('should parse multiple elements', () => {
      const xml = '<root><item id="1"/><item id="2"/><item id="3"/></root>';
      const items = parseElements(xml, 'item');
      expect(items.length).toBe(3);
      expect(items[0].attrs.id).toBe('1');
      expect(items[2].attrs.id).toBe('3');
    });

    it('should handle nested same-name elements', () => {
      const xml = '<outer><div><div>inner</div></div></outer>';
      const divs = parseElements(xml, 'div');
      expect(divs.length).toBe(2);
    });
  });

  describe('getAttr', () => {
    it('should get attribute from ParsedElement', () => {
      const el = parseElement('<item name="test" value="123"/>', 'item');
      expect(getAttr(el!, 'name')).toBe('test');
      expect(getAttr(el!, 'value')).toBe('123');
      expect(getAttr(el!, 'missing')).toBeUndefined();
    });

    it('should get attribute from string', () => {
      expect(getAttr(' name="test" value="123"', 'name')).toBe('test');
    });
  });

  describe('getTextContent', () => {
    it('should get text content', () => {
      expect(getTextContent('<root><name>Hello</name></root>', 'name')).toBe('Hello');
    });

    it('should return undefined for missing element', () => {
      expect(getTextContent('<root></root>', 'missing')).toBeUndefined();
    });

    it('should unescape XML entities in content', () => {
      expect(getTextContent('<root><val>&lt;test&gt;</val></root>', 'val')).toBe('<test>');
    });
  });

  describe('unescapeXml (parser)', () => {
    it('should unescape all entities', () => {
      expect(unescapeXml('&lt;&gt;&quot;&apos;&amp;')).toBe('<>"\'&');
    });
  });
});

describe('xlsx.reader.ts Coverage', () => {
  describe('Error Handling', () => {
    it('should throw on invalid ZIP data', () => {
      const invalidData = new Uint8Array([1, 2, 3, 4, 5]);
      expect(() => xlsxToWorkbook(invalidData)).toThrow('Failed to unzip XLSX file');
    });

    it('should throw on missing workbook.xml', () => {
      // Create a valid ZIP but missing workbook.xml
      const files = {
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
      };
      const xlsx = zipSync(files);
      expect(() => xlsxToWorkbook(xlsx)).toThrow('Invalid XLSX: missing xl/workbook.xml');
    });

    it('should add warning for missing sheet file', () => {
      // Create minimal valid XLSX with sheet reference but no sheet file
      const files = {
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
        'xl/workbook.xml': strToU8(`<?xml version="1.0"?>
          <workbook><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0"?>
          <Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`),
      };
      const xlsx = zipSync(files);
      const { warnings } = xlsxToWorkbook(xlsx);
      expect(warnings.some((w) => w.code === 'MISSING_SHEET')).toBe(true);
    });
  });

  describe('Progress Callbacks', () => {
    it('should call progress callback during import', () => {
      const wb = new Workbook();
      wb.addSheet('Test').cell('A1').value = 'Hello';
      const xlsx = workbookToXlsx(wb);

      const progressCalls: Array<{ phase: string; current: number; total: number }> = [];
      xlsxToWorkbook(xlsx, {
        onProgress: (phase, current, total) => {
          progressCalls.push({ phase, current, total });
        },
      });

      expect(progressCalls.length).toBeGreaterThan(0);
      expect(progressCalls.some((p) => p.phase === 'unzip')).toBe(true);
      expect(progressCalls.some((p) => p.phase === 'sheets')).toBe(true);
    });
  });

  describe('Blob Import', () => {
    it('should import from Blob', async () => {
      const wb = new Workbook();
      wb.addSheet('Test').cell('A1').value = 'Blob Test';
      const xlsx = workbookToXlsx(wb);
      const blob = new Blob([xlsx], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      const { workbook } = await xlsxBlobToWorkbook(blob);
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Blob Test');
    });
  });

  describe('Inline Strings', () => {
    it('should handle inline string cells', () => {
      // Create XLSX with inline string
      const sheetXml = `<?xml version="1.0"?>
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1" t="inlineStr"><is><t>Inline Text</t></is></c>
            </row>
          </sheetData>
        </worksheet>`;

      const files = {
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
        'xl/workbook.xml': strToU8(`<?xml version="1.0"?>
          <workbook><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0"?>
          <Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`),
        'xl/worksheets/sheet1.xml': strToU8(sheetXml),
        'xl/styles.xml': strToU8('<?xml version="1.0"?><styleSheet/>'),
      };
      const xlsx = zipSync(files);
      const { workbook } = xlsxToWorkbook(xlsx);
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Inline Text');
    });

    it('should handle inline rich text', () => {
      const sheetXml = `<?xml version="1.0"?>
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1" t="inlineStr"><is><r><t>RichText</t></r></is></c>
            </row>
          </sheetData>
        </worksheet>`;

      const files = {
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
        'xl/workbook.xml': strToU8(`<?xml version="1.0"?>
          <workbook><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0"?>
          <Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`),
        'xl/worksheets/sheet1.xml': strToU8(sheetXml),
        'xl/styles.xml': strToU8('<?xml version="1.0"?><styleSheet/>'),
      };
      const xlsx = zipSync(files);
      const { workbook } = xlsxToWorkbook(xlsx);
      // Rich text is flattened - covers the rich text code path
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('RichText');
    });
  });

  describe('Error Cell Values', () => {
    it('should import error values', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      // Errors are exported as t="e" with value like #DIV/0!
      sheet.cell('A1').value = '#DIV/0!';

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      // Note: Errors are exported as strings by default
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('#DIV/0!');
    });
  });

  describe('AutoFilter', () => {
    it('should import autofilter', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Header';
      sheet.cell('A2').value = 'Data';
      sheet.setAutoFilter('A1:A2');

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      expect(imported?.autoFilter).toBeDefined();
    });
  });

  describe('Border Styles', () => {
    it('should import various border styles', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Border';
      sheet.cell('A1').style = {
        borders: {
          top: { style: 'thin', color: '#000000' },
          bottom: { style: 'medium', color: '#FF0000' },
          left: { style: 'dashed', color: '#00FF00' },
          right: { style: 'dotted', color: '#0000FF' },
        },
      };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      const borders = imported?.cell('A1').style?.borders;
      expect(borders?.top?.style).toBe('thin');
      expect(borders?.bottom?.style).toBe('medium');
    });
  });

  describe('Alignment', () => {
    it('should import text alignment', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Aligned';
      sheet.cell('A1').style = {
        alignment: {
          horizontal: 'center',
          vertical: 'middle',
          wrapText: true,
          textRotation: 45,
        },
      };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      const alignment = imported?.cell('A1').style?.alignment;
      expect(alignment?.horizontal).toBe('center');
      expect(alignment?.wrapText).toBe(true);
    });
  });

  describe('Custom Date Formats', () => {
    it('should detect custom date format patterns', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      const testDate = new Date('2024-06-15T00:00:00.000Z');
      sheet.cell('A1').value = testDate;
      sheet.cell('A1').style = { numberFormat: { formatCode: 'dd/mm/yyyy' } };

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx);

      const imported = workbook.getSheet('Test');
      const value = imported?.cell('A1').value;
      expect(value).toBeInstanceOf(Date);
    });
  });

  describe('Rich Text in Shared Strings', () => {
    it('should flatten rich text to plain text', () => {
      // Create XLSX with rich text shared string (single run)
      const sstXml = `<?xml version="1.0"?>
        <sst count="1" uniqueCount="1">
          <si>
            <r><rPr><b/></rPr><t>BoldText</t></r>
          </si>
        </sst>`;

      const sheetXml = `<?xml version="1.0"?>
        <worksheet>
          <sheetData>
            <row r="1"><c r="A1" t="s"><v>0</v></c></row>
          </sheetData>
        </worksheet>`;

      const files = {
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
        'xl/workbook.xml': strToU8(`<?xml version="1.0"?>
          <workbook><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0"?>
          <Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`),
        'xl/worksheets/sheet1.xml': strToU8(sheetXml),
        'xl/sharedStrings.xml': strToU8(sstXml),
        'xl/styles.xml': strToU8('<?xml version="1.0"?><styleSheet/>'),
      };
      const xlsx = zipSync(files);
      const { workbook } = xlsxToWorkbook(xlsx);
      // Rich text is flattened - covers the code path
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('BoldText');
    });
  });

  describe('Empty Shared String', () => {
    it('should handle empty shared string entries', () => {
      const sstXml = `<?xml version="1.0"?>
        <sst count="2" uniqueCount="2">
          <si><t>Text</t></si>
          <si></si>
        </sst>`;

      const sheetXml = `<?xml version="1.0"?>
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1" t="s"><v>1</v></c>
            </row>
          </sheetData>
        </worksheet>`;

      const files = {
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
        'xl/workbook.xml': strToU8(`<?xml version="1.0"?>
          <workbook><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0"?>
          <Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`),
        'xl/worksheets/sheet1.xml': strToU8(sheetXml),
        'xl/sharedStrings.xml': strToU8(sstXml),
        'xl/styles.xml': strToU8('<?xml version="1.0"?><styleSheet/>'),
      };
      const xlsx = zipSync(files);
      const { workbook } = xlsxToWorkbook(xlsx);
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Text');
      expect(workbook.getSheet('Test')?.cell('B1').value).toBe('');
    });
  });

  describe('Formula String Result', () => {
    it('should handle formula with string result type', () => {
      const sheetXml = `<?xml version="1.0"?>
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1" t="str"><f>CONCATENATE("Hello"," ","World")</f><v>Hello World</v></c>
            </row>
          </sheetData>
        </worksheet>`;

      const files = {
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
        'xl/workbook.xml': strToU8(`<?xml version="1.0"?>
          <workbook><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0"?>
          <Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`),
        'xl/worksheets/sheet1.xml': strToU8(sheetXml),
        'xl/styles.xml': strToU8('<?xml version="1.0"?><styleSheet/>'),
      };
      const xlsx = zipSync(files);
      const { workbook } = xlsxToWorkbook(xlsx);
      expect(workbook.getSheet('Test')?.cell('A1').value).toBe('Hello World');
    });
  });

  describe('Import Options', () => {
    it('should respect importDimensions=false', () => {
      const wb = new Workbook();
      const sheet = wb.addSheet('Test');
      sheet.cell('A1').value = 'Test';
      sheet.setColumnWidth(0, 25);
      sheet.setRowHeight(0, 30);

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx, { importDimensions: false });

      const imported = workbook.getSheet('Test');
      // Column width should not be set
      expect(imported?.getColumn(0)?.width).toBeUndefined();
    });

    it('should respect importProperties=false', () => {
      const wb = new Workbook();
      wb.properties.title = 'Test Title';
      wb.addSheet('Test');

      const xlsx = workbookToXlsx(wb);
      const { workbook } = xlsxToWorkbook(xlsx, { importProperties: false });

      expect(workbook.properties.title).toBeUndefined();
    });
  });
});

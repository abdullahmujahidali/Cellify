/**
 * Style Registry for XLSX export
 *
 * Excel stores styles in a centralized styles.xml file with separate collections:
 * - fonts: Font definitions
 * - fills: Fill patterns/colors
 * - borders: Border definitions
 * - numFmts: Number format codes
 * - cellXfs: Cell format records (combining the above)
 *
 * Cells reference styles via the "s" attribute, which is an index into cellXfs.
 * This registry deduplicates styles to minimize file size.
 */

import type { CellStyle, CellFont, CellFill, CellBorders, CellAlignment, NumberFormat, Border } from '../../types/style.types.js';
import type { StyleRegistry as IStyleRegistry, CellXf } from './xlsx.types.js';
import { XML_DECLARATION, escapeXml, toArgbColor } from './xlsx.xml.js';
import { NS, BUILTIN_NUM_FMT_IDS } from './xlsx.types.js';

/**
 * Style Registry implementation
 */
export class StyleRegistry implements IStyleRegistry {
  // Maps for deduplication (hash -> index)
  private fonts: Map<string, number> = new Map();
  private fills: Map<string, number> = new Map();
  private borders: Map<string, number> = new Map();
  private numFmts: Map<string, number> = new Map();
  private cellXfs: Map<string, number> = new Map();

  // Ordered lists for XML generation
  private fontList: CellFont[] = [];
  private fillList: CellFill[] = [];
  private borderList: CellBorders[] = [];
  private numFmtList: Array<{ id: number; formatCode: string }> = [];
  private cellXfList: CellXf[] = [];

  // Next custom number format ID (164+ are custom, 0-163 reserved)
  private nextNumFmtId = 164;

  constructor() {
    this.addDefaultStyles();
  }

  /**
   * Add required default styles
   * Excel requires at least one entry in each collection
   */
  private addDefaultStyles(): void {
    // Default font (Calibri 11pt - Excel's default)
    this.fontList.push({ name: 'Calibri', size: 11 });
    this.fonts.set(this.hashFont({ name: 'Calibri', size: 11 }), 0);

    // Excel requires exactly these two fills first:
    // Index 0: "none" pattern
    // Index 1: "gray125" pattern
    this.fillList.push({ type: 'pattern', pattern: 'none' });
    this.fills.set(this.hashFill({ type: 'pattern', pattern: 'none' }), 0);
    this.fillList.push({ type: 'pattern', pattern: 'gray125' });
    this.fills.set(this.hashFill({ type: 'pattern', pattern: 'gray125' }), 1);

    // Default border (empty)
    this.borderList.push({});
    this.borders.set(this.hashBorders({}), 0);

    // Default cell format (xf) - combines defaults
    this.cellXfList.push({
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId: 0,
    });
    this.cellXfs.set('0|0|0|0||', 0);
  }

  /**
   * Register a complete cell style and return its xf index
   *
   * @param style - CellStyle to register (or undefined for default)
   * @returns Index to use as "s" attribute on cell
   */
  registerStyle(style: CellStyle | undefined): number {
    if (!style) return 0; // Default style

    const fontId = style.font ? this.registerFont(style.font) : 0;
    const fillId = style.fill ? this.registerFill(style.fill) : 0;
    const borderId = style.borders ? this.registerBorders(style.borders) : 0;
    const numFmtId = style.numberFormat ? this.registerNumFmt(style.numberFormat) : 0;

    const alignKey = style.alignment ? JSON.stringify(style.alignment) : '';
    const protKey = style.protection ? JSON.stringify(style.protection) : '';

    const xfKey = `${fontId}|${fillId}|${borderId}|${numFmtId}|${alignKey}|${protKey}`;

    let xfIndex = this.cellXfs.get(xfKey);
    if (xfIndex === undefined) {
      xfIndex = this.cellXfList.length;
      this.cellXfList.push({
        fontId,
        fillId,
        borderId,
        numFmtId,
        alignment: style.alignment,
        protection: style.protection,
        applyFont: style.font !== undefined,
        applyFill: style.fill !== undefined,
        applyBorder: style.borders !== undefined,
        applyAlignment: style.alignment !== undefined,
        applyNumberFormat: style.numberFormat !== undefined,
        applyProtection: style.protection !== undefined,
      });
      this.cellXfs.set(xfKey, xfIndex);
    }

    return xfIndex;
  }

  /**
   * Register a font and return its index
   */
  private registerFont(font: CellFont): number {
    const key = this.hashFont(font);
    let idx = this.fonts.get(key);
    if (idx === undefined) {
      idx = this.fontList.length;
      this.fontList.push(font);
      this.fonts.set(key, idx);
    }
    return idx;
  }

  /**
   * Register a fill and return its index
   */
  private registerFill(fill: CellFill): number {
    const key = this.hashFill(fill);
    let idx = this.fills.get(key);
    if (idx === undefined) {
      idx = this.fillList.length;
      this.fillList.push(fill);
      this.fills.set(key, idx);
    }
    return idx;
  }

  /**
   * Register borders and return index
   */
  private registerBorders(borders: CellBorders): number {
    const key = this.hashBorders(borders);
    let idx = this.borders.get(key);
    if (idx === undefined) {
      idx = this.borderList.length;
      this.borderList.push(borders);
      this.borders.set(key, idx);
    }
    return idx;
  }

  /**
   * Register a number format and return its ID
   */
  private registerNumFmt(fmt: NumberFormat): number {
    // Check for built-in format
    const builtInId = BUILTIN_NUM_FMT_IDS[fmt.formatCode];
    if (builtInId !== undefined) {
      return builtInId;
    }

    // Check if we already have this custom format
    let id = this.numFmts.get(fmt.formatCode);
    if (id === undefined) {
      id = this.nextNumFmtId++;
      this.numFmtList.push({ id, formatCode: fmt.formatCode });
      this.numFmts.set(fmt.formatCode, id);
    }
    return id;
  }

  // Hash functions for deduplication
  private hashFont(f: CellFont): string {
    return JSON.stringify([
      f.name,
      f.size,
      f.color,
      f.bold,
      f.italic,
      f.underline,
      f.strikethrough,
      f.superscript,
      f.subscript,
    ]);
  }

  private hashFill(f: CellFill): string {
    return JSON.stringify(f);
  }

  private hashBorders(b: CellBorders): string {
    return JSON.stringify(b);
  }

  /**
   * Generate styles.xml content
   */
  generateStylesXml(): string {
    const parts: string[] = [XML_DECLARATION];

    parts.push(`<styleSheet xmlns="${NS.spreadsheetml}">`);

    // Number formats (custom only)
    if (this.numFmtList.length > 0) {
      parts.push(`<numFmts count="${this.numFmtList.length}">`);
      for (const fmt of this.numFmtList) {
        parts.push(`<numFmt numFmtId="${fmt.id}" formatCode="${escapeXml(fmt.formatCode)}"/>`);
      }
      parts.push('</numFmts>');
    }

    // Fonts
    parts.push(`<fonts count="${this.fontList.length}">`);
    for (const font of this.fontList) {
      parts.push(this.generateFontXml(font));
    }
    parts.push('</fonts>');

    // Fills
    parts.push(`<fills count="${this.fillList.length}">`);
    for (const fill of this.fillList) {
      parts.push(this.generateFillXml(fill));
    }
    parts.push('</fills>');

    // Borders
    parts.push(`<borders count="${this.borderList.length}">`);
    for (const border of this.borderList) {
      parts.push(this.generateBorderXml(border));
    }
    parts.push('</borders>');

    // Cell style formats (cellStyleXfs) - required, at least one entry
    parts.push('<cellStyleXfs count="1">');
    parts.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
    parts.push('</cellStyleXfs>');

    // Cell formats (cellXfs)
    parts.push(`<cellXfs count="${this.cellXfList.length}">`);
    for (const xf of this.cellXfList) {
      parts.push(this.generateCellXfXml(xf));
    }
    parts.push('</cellXfs>');

    // Cell styles - required, at least one entry
    parts.push('<cellStyles count="1">');
    parts.push('<cellStyle name="Normal" xfId="0" builtinId="0"/>');
    parts.push('</cellStyles>');

    parts.push('</styleSheet>');

    return parts.join('\n');
  }

  /**
   * Generate XML for a font
   */
  private generateFontXml(font: CellFont): string {
    const parts: string[] = ['<font>'];

    if (font.bold) parts.push('<b/>');
    if (font.italic) parts.push('<i/>');
    if (font.strikethrough) parts.push('<strike/>');
    if (font.underline) {
      if (font.underline === 'double') {
        parts.push('<u val="double"/>');
      } else {
        parts.push('<u/>');
      }
    }
    if (font.superscript) parts.push('<vertAlign val="superscript"/>');
    if (font.subscript) parts.push('<vertAlign val="subscript"/>');

    if (font.size) parts.push(`<sz val="${font.size}"/>`);
    if (font.color) parts.push(`<color rgb="${toArgbColor(font.color)}"/>`);
    if (font.name) parts.push(`<name val="${escapeXml(font.name)}"/>`);

    parts.push('</font>');
    return parts.join('');
  }

  /**
   * Generate XML for a fill
   */
  private generateFillXml(fill: CellFill): string {
    if (fill.type === 'pattern') {
      const pattern = this.mapPatternType(fill.pattern || 'none');

      if (pattern === 'none' || pattern === 'gray125') {
        return `<fill><patternFill patternType="${pattern}"/></fill>`;
      }

      const parts: string[] = [`<fill><patternFill patternType="${pattern}">`];

      if (fill.foregroundColor) {
        parts.push(`<fgColor rgb="${toArgbColor(fill.foregroundColor)}"/>`);
      }
      if (fill.backgroundColor) {
        parts.push(`<bgColor rgb="${toArgbColor(fill.backgroundColor)}"/>`);
      }

      parts.push('</patternFill></fill>');
      return parts.join('');
    }

    // Gradient fill
    if (fill.type === 'gradient') {
      const parts: string[] = ['<fill><gradientFill'];

      if (fill.degree !== undefined) {
        parts.push(` degree="${fill.degree}"`);
      }

      parts.push('>');

      if (fill.stops) {
        for (const stop of fill.stops) {
          parts.push(`<stop position="${stop.position}"><color rgb="${toArgbColor(stop.color)}"/></stop>`);
        }
      }

      parts.push('</gradientFill></fill>');
      return parts.join('');
    }

    return '<fill><patternFill patternType="none"/></fill>';
  }

  /**
   * Map Cellify pattern names to Excel pattern names
   */
  private mapPatternType(pattern: string): string {
    const mapping: Record<string, string> = {
      none: 'none',
      solid: 'solid',
      gray125: 'gray125',
      gray0625: 'gray0625',
      darkGray: 'darkGray',
      mediumGray: 'mediumGray',
      lightGray: 'lightGray',
      darkHorizontal: 'darkHorizontal',
      darkVertical: 'darkVertical',
      darkDown: 'darkDown',
      darkUp: 'darkUp',
      darkGrid: 'darkGrid',
      darkTrellis: 'darkTrellis',
      lightHorizontal: 'lightHorizontal',
      lightVertical: 'lightVertical',
      lightDown: 'lightDown',
      lightUp: 'lightUp',
      lightGrid: 'lightGrid',
      lightTrellis: 'lightTrellis',
    };
    return mapping[pattern] || 'solid';
  }

  /**
   * Generate XML for borders
   */
  private generateBorderXml(borders: CellBorders): string {
    const parts: string[] = ['<border>'];

    parts.push(this.generateBorderSideXml('left', borders.left));
    parts.push(this.generateBorderSideXml('right', borders.right));
    parts.push(this.generateBorderSideXml('top', borders.top));
    parts.push(this.generateBorderSideXml('bottom', borders.bottom));
    parts.push(this.generateBorderSideXml('diagonal', borders.diagonal));

    parts.push('</border>');
    return parts.join('');
  }

  /**
   * Generate XML for a single border side
   */
  private generateBorderSideXml(side: string, border: Border | undefined): string {
    if (!border || !border.style || border.style === 'none') {
      return `<${side}/>`;
    }

    const style = this.mapBorderStyle(border.style);

    if (border.color) {
      return `<${side} style="${style}"><color rgb="${toArgbColor(border.color)}"/></${side}>`;
    }

    return `<${side} style="${style}"/>`;
  }

  /**
   * Map Cellify border style to Excel border style
   */
  private mapBorderStyle(style: string): string {
    const mapping: Record<string, string> = {
      thin: 'thin',
      medium: 'medium',
      thick: 'thick',
      dashed: 'dashed',
      dotted: 'dotted',
      double: 'double',
      hair: 'hair',
      mediumDashed: 'mediumDashed',
      dashDot: 'dashDot',
      mediumDashDot: 'mediumDashDot',
      dashDotDot: 'dashDotDot',
      mediumDashDotDot: 'mediumDashDotDot',
      slantDashDot: 'slantDashDot',
    };
    return mapping[style] || 'thin';
  }

  /**
   * Generate XML for a cell format (xf)
   */
  private generateCellXfXml(xf: CellXf): string {
    const attrs: string[] = [
      `numFmtId="${xf.numFmtId}"`,
      `fontId="${xf.fontId}"`,
      `fillId="${xf.fillId}"`,
      `borderId="${xf.borderId}"`,
    ];

    if (xf.applyFont) attrs.push('applyFont="1"');
    if (xf.applyFill) attrs.push('applyFill="1"');
    if (xf.applyBorder) attrs.push('applyBorder="1"');
    if (xf.applyNumberFormat) attrs.push('applyNumberFormat="1"');
    if (xf.applyAlignment) attrs.push('applyAlignment="1"');
    if (xf.applyProtection) attrs.push('applyProtection="1"');

    // Check if we need alignment or protection children
    const hasChildren = xf.alignment || xf.protection;

    if (!hasChildren) {
      return `<xf ${attrs.join(' ')}/>`;
    }

    const parts: string[] = [`<xf ${attrs.join(' ')}>`];

    if (xf.alignment) {
      parts.push(this.generateAlignmentXml(xf.alignment));
    }

    if (xf.protection) {
      const protAttrs: string[] = [];
      if (xf.protection.locked !== undefined) {
        protAttrs.push(`locked="${xf.protection.locked ? 1 : 0}"`);
      }
      if (xf.protection.hidden !== undefined) {
        protAttrs.push(`hidden="${xf.protection.hidden ? 1 : 0}"`);
      }
      if (protAttrs.length > 0) {
        parts.push(`<protection ${protAttrs.join(' ')}/>`);
      }
    }

    parts.push('</xf>');
    return parts.join('');
  }

  /**
   * Generate XML for alignment
   */
  private generateAlignmentXml(align: CellAlignment): string {
    const attrs: string[] = [];

    if (align.horizontal) {
      const hMap: Record<string, string> = {
        left: 'left',
        center: 'center',
        right: 'right',
        fill: 'fill',
        justify: 'justify',
        centerContinuous: 'centerContinuous',
        distributed: 'distributed',
      };
      attrs.push(`horizontal="${hMap[align.horizontal] || align.horizontal}"`);
    }

    if (align.vertical) {
      const vMap: Record<string, string> = {
        top: 'top',
        center: 'center',
        bottom: 'bottom',
        justify: 'justify',
        distributed: 'distributed',
      };
      attrs.push(`vertical="${vMap[align.vertical] || align.vertical}"`);
    }

    if (align.wrapText) attrs.push('wrapText="1"');
    if (align.shrinkToFit) attrs.push('shrinkToFit="1"');
    if (align.textRotation !== undefined) attrs.push(`textRotation="${align.textRotation}"`);
    if (align.indent !== undefined) attrs.push(`indent="${align.indent}"`);

    return `<alignment ${attrs.join(' ')}/>`;
  }
}

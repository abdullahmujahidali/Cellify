/**
 * Border line styles supported by Excel
 */
export type BorderStyle =
  | 'none'
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

/**
 * Border definition for a single edge
 */
export interface Border {
  style: BorderStyle;
  color: string; // Hex color e.g., "#000000"
}

/**
 * Complete border configuration for a cell
 */
export interface CellBorders {
  top?: Border;
  right?: Border;
  bottom?: Border;
  left?: Border;
  diagonal?: Border;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

/**
 * Horizontal alignment options
 */
export type HorizontalAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

/**
 * Vertical alignment options
 */
export type VerticalAlignment = 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';

/**
 * Text alignment and wrapping configuration
 */
export interface CellAlignment {
  horizontal?: HorizontalAlignment;
  vertical?: VerticalAlignment;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number; // 0-255
  textRotation?: number; // -90 to 90, or 255 for vertical text
  readingOrder?: 'contextDependent' | 'leftToRight' | 'rightToLeft';
}

/**
 * Font underline styles
 */
export type UnderlineStyle = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/**
 * Font configuration
 */
export interface CellFont {
  name?: string; // Font family e.g., "Arial", "Calibri"
  size?: number; // Font size in points
  color?: string; // Hex color
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
}

/**
 * Fill pattern types supported by Excel
 */
export type FillPattern =
  | 'none'
  | 'solid'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'gray125'
  | 'gray0625'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis';

/**
 * Gradient fill types
 */
export type GradientType = 'linear' | 'path';

/**
 * Gradient stop definition
 */
export interface GradientStop {
  position: number; // 0 to 1
  color: string; // Hex color
}

/**
 * Fill configuration - either pattern or gradient
 */
export interface CellFill {
  type: 'pattern' | 'gradient';

  // Pattern fill properties
  pattern?: FillPattern;
  foregroundColor?: string; // Hex color
  backgroundColor?: string; // Hex color (for patterns)

  // Gradient fill properties
  gradientType?: GradientType;
  degree?: number; // Rotation angle for linear gradient (0-360)
  stops?: GradientStop[];
  // For path gradients
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
}

/**
 * Number format categories
 */
export type NumberFormatCategory =
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'date'
  | 'time'
  | 'percentage'
  | 'fraction'
  | 'scientific'
  | 'text'
  | 'custom';

/**
 * Number format configuration
 */
export interface NumberFormat {
  formatCode: string; // e.g., "0.00", "$#,##0.00", "yyyy-mm-dd"
  category?: NumberFormatCategory;
}

/**
 * Cell protection settings
 */
export interface CellProtection {
  locked?: boolean;
  hidden?: boolean; // Hide formula
}

/**
 * Complete cell style definition
 */
export interface CellStyle {
  font?: CellFont;
  fill?: CellFill;
  borders?: CellBorders;
  alignment?: CellAlignment;
  numberFormat?: NumberFormat;
  protection?: CellProtection;
}

/**
 * Named style that can be reused across cells
 */
export interface NamedStyle {
  name: string;
  style: CellStyle;
  builtIn?: boolean; // Whether this is a built-in Excel style
}

/**
 * Predefined number formats (Excel built-in format IDs)
 */
export const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
} as const;

/**
 * Default style values
 */
export const DEFAULT_STYLE: CellStyle = {
  font: {
    name: 'Calibri',
    size: 11,
    color: '#000000',
    bold: false,
    italic: false,
    underline: 'none',
    strikethrough: false,
  },
  fill: {
    type: 'pattern',
    pattern: 'none',
  },
  alignment: {
    horizontal: 'left',
    vertical: 'bottom',
    wrapText: false,
  },
  numberFormat: {
    formatCode: 'General',
    category: 'general',
  },
  protection: {
    locked: true,
    hidden: false,
  },
};

/**
 * Helper to create a solid fill
 */
export function createSolidFill(color: string): CellFill {
  return {
    type: 'pattern',
    pattern: 'solid',
    foregroundColor: color,
  };
}

/**
 * Helper to create a simple border
 */
export function createBorder(style: BorderStyle, color: string = '#000000'): Border {
  return { style, color };
}

/**
 * Helper to create uniform borders on all sides
 */
export function createUniformBorders(style: BorderStyle, color: string = '#000000'): CellBorders {
  const border = createBorder(style, color);
  return {
    top: border,
    right: border,
    bottom: border,
    left: border,
  };
}

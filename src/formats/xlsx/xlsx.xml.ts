/**
 * XML generation utilities for XLSX export
 */

/**
 * XML declaration for all OOXML parts
 */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Escape special XML characters
 * @param str - String to escape
 * @returns Escaped string safe for XML content
 */
export function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Unescape XML entities
 * @param str - String with XML entities
 * @returns Unescaped string
 */
export function unescapeXml(str: string): string {
  return str
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&'); // Must be last to avoid double-unescaping
}

/**
 * Sanitize string for XML 1.0
 * Removes invalid control characters and escapes special characters
 *
 * XML 1.0 valid chars: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
 */
export function sanitizeXmlString(str: string): string {
  // Remove invalid XML 1.0 control characters
  const sanitized = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  return escapeXml(sanitized);
}

/**
 * Create an XML attribute string
 * Returns empty string if value is undefined
 */
export function attr(name: string, value: string | number | boolean | undefined): string {
  if (value === undefined) return '';
  return ` ${name}="${escapeXml(String(value))}"`;
}

/**
 * Create XML attributes from an object
 * Filters out undefined values
 */
export function attrs(obj: Record<string, string | number | boolean | undefined>): string {
  return Object.entries(obj)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => ` ${k}="${escapeXml(String(v))}"`)
    .join('');
}

/**
 * Create a self-closing XML element
 */
export function emptyEl(tag: string, attributes: Record<string, string | number | boolean | undefined> = {}): string {
  return `<${tag}${attrs(attributes)}/>`;
}

/**
 * Create an XML element with content
 */
export function el(
  tag: string,
  attributes: Record<string, string | number | boolean | undefined>,
  content: string
): string {
  return `<${tag}${attrs(attributes)}>${content}</${tag}>`;
}

/**
 * Create an XML element with optional content (self-closing if no content)
 */
export function optEl(
  tag: string,
  attributes: Record<string, string | number | boolean | undefined> = {},
  content?: string
): string {
  if (content === undefined || content === '') {
    return emptyEl(tag, attributes);
  }
  return el(tag, attributes, content);
}

/**
 * Convert color string to ARGB format for Excel
 * Excel expects 8-character ARGB (e.g., "FF000000" for black)
 *
 * @param color - Color in various formats (#RGB, #RRGGBB, RRGGBB, AARRGGBB)
 * @returns 8-character ARGB string
 * @throws Error if color format is invalid
 */
export function toArgbColor(color: string): string {
  // Remove # prefix if present
  let hex = color.replace(/^#/, '');

  // Expand shorthand (#RGB → RRGGBB)
  if (hex.length === 3) {
    hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
  }

  // Add alpha channel if not present (FF = fully opaque)
  if (hex.length === 6) {
    hex = 'FF' + hex;
  } else if (hex.length !== 8) {
    throw new Error(`Invalid color format: "${color}". Expected #RGB, #RRGGBB, or AARRGGBB.`);
  }

  return hex.toUpperCase();
}

/**
 * Check if a value is a valid Excel error string
 */
export function isExcelError(value: string): boolean {
  const errors = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
  return errors.includes(value);
}

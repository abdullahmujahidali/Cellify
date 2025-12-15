/**
 * XML parsing utilities for XLSX import
 *
 * Uses regex-based parsing instead of DOM/SAX for:
 * - Zero dependencies (no xml2js, fast-xml-parser, etc.)
 * - Predictable performance
 * - Sufficient for well-formed OOXML
 */

/**
 * Parsed XML element
 */
export interface ParsedElement {
  /** Element tag name */
  tag: string;
  /** Element attributes */
  attrs: Record<string, string>;
  /** Inner content (everything between open and close tags) */
  inner: string;
  /** Full match including tags */
  full: string;
}

/**
 * Unescape XML entities
 *
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
 * Parse a single element by tag name
 * Returns the first match or undefined
 *
 * @param xml - XML string to search
 * @param tagName - Tag name to find (without namespace prefix by default)
 * @returns Parsed element or undefined
 */
export function parseElement(xml: string, tagName: string): ParsedElement | undefined {
  // Match both namespaced and non-namespaced versions
  // E.g., "sheet" matches both <sheet ...> and <x:sheet ...>
  const pattern = new RegExp(
    `<(?:[a-zA-Z0-9_]+:)?${escapeRegex(tagName)}(\\s[^>]*)?>([\\s\\S]*?)</(?:[a-zA-Z0-9_]+:)?${escapeRegex(tagName)}>`,
    'i'
  );

  const match = xml.match(pattern);
  if (!match) {
    // Try self-closing
    const selfClosing = new RegExp(
      `<(?:[a-zA-Z0-9_]+:)?${escapeRegex(tagName)}(\\s[^>]*)?\\/?>`,
      'i'
    );
    const selfMatch = xml.match(selfClosing);
    if (selfMatch) {
      return {
        tag: tagName,
        attrs: parseAttributes(selfMatch[1] || ''),
        inner: '',
        full: selfMatch[0],
      };
    }
    return undefined;
  }

  return {
    tag: tagName,
    attrs: parseAttributes(match[1] || ''),
    inner: match[2],
    full: match[0],
  };
}

/**
 * Parse all elements with given tag name
 *
 * @param xml - XML string to search
 * @param tagName - Tag name to find
 * @returns Array of parsed elements
 */
export function parseElements(xml: string, tagName: string): ParsedElement[] {
  const results: ParsedElement[] = [];

  // Match opening tags (both namespaced and non-namespaced)
  const pattern = new RegExp(
    `<(?:[a-zA-Z0-9_]+:)?${escapeRegex(tagName)}(\\s[^>]*)?>`,
    'gi'
  );

  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    const startIndex = match.index;
    const attrString = match[1] || '';
    const isSelfClosing = match[0].endsWith('/>');

    if (isSelfClosing) {
      results.push({
        tag: tagName,
        attrs: parseAttributes(attrString),
        inner: '',
        full: match[0],
      });
      continue;
    }

    // Find matching close tag
    const inner = findInnerContent(xml, startIndex + match[0].length, tagName);
    if (inner !== null) {
      const closeTag = `</${tagName}>`;
      const nsClosePattern = new RegExp(`</[a-zA-Z0-9_]+:${escapeRegex(tagName)}>`);
      const nsCloseMatch = xml.slice(startIndex + match[0].length + inner.length).match(nsClosePattern);
      const actualCloseTag = nsCloseMatch ? nsCloseMatch[0] : closeTag;

      results.push({
        tag: tagName,
        attrs: parseAttributes(attrString),
        inner,
        full: match[0] + inner + actualCloseTag,
      });
    }
  }

  return results;
}

/**
 * Find inner content handling nested same-name tags
 */
function findInnerContent(xml: string, startPos: number, tagName: string): string | null {
  const openPattern = new RegExp(`<(?:[a-zA-Z0-9_]+:)?${escapeRegex(tagName)}(?:\\s[^>]*)?>`, 'gi');
  const closePattern = new RegExp(`</(?:[a-zA-Z0-9_]+:)?${escapeRegex(tagName)}>`, 'gi');

  let depth = 1;
  const remaining = xml.slice(startPos);

  // Find all opens and closes, track depth
  const opens: number[] = [];
  const closes: number[] = [];

  let m: RegExpExecArray | null;

  openPattern.lastIndex = 0;
  while ((m = openPattern.exec(remaining)) !== null) {
    if (!m[0].endsWith('/>')) {
      opens.push(m.index);
    }
  }

  closePattern.lastIndex = 0;
  while ((m = closePattern.exec(remaining)) !== null) {
    closes.push(m.index);
  }

  // Merge and sort positions
  const events: Array<{ pos: number; type: 'open' | 'close' }> = [
    ...opens.map((p) => ({ pos: p, type: 'open' as const })),
    ...closes.map((p) => ({ pos: p, type: 'close' as const })),
  ].sort((a, b) => a.pos - b.pos);

  for (const event of events) {
    if (event.type === 'open') {
      depth++;
    } else {
      depth--;
      if (depth === 0) {
        return remaining.slice(0, event.pos);
      }
    }
  }

  return null;
}

/**
 * Get a single attribute value from element
 *
 * @param el - Parsed element or attribute string
 * @param attrName - Attribute name to get
 * @returns Attribute value or undefined
 */
export function getAttr(el: ParsedElement | string, attrName: string): string | undefined {
  const attrs = typeof el === 'string' ? parseAttributes(el) : el.attrs;
  return attrs[attrName];
}

/**
 * Get text content of a child element
 *
 * @param xml - XML string
 * @param tagName - Tag name of child element
 * @returns Text content or undefined
 */
export function getTextContent(xml: string, tagName: string): string | undefined {
  const el = parseElement(xml, tagName);
  if (!el) return undefined;

  // Get text content (strip any nested tags)
  const text = el.inner.replace(/<[^>]+>/g, '');
  return unescapeXml(text);
}

/**
 * Parse attribute string into key-value object
 */
function parseAttributes(attrString: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!attrString) return attrs;

  // Match attribute="value" or attribute='value'
  const pattern = /([a-zA-Z0-9_:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(attrString)) !== null) {
    const name = match[1];
    const value = match[2] ?? match[3] ?? '';
    attrs[name] = unescapeXml(value);
  }

  return attrs;
}

/**
 * Escape special regex characters
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Parse cell reference (A1 notation) to row/col indices
 *
 * @param ref - Cell reference (e.g., "A1", "AA100")
 * @returns Object with 0-based row and col indices
 */
export function parseCellRef(ref: string): { row: number; col: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }

  const colStr = match[1].toUpperCase();
  const rowStr = match[2];

  // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, etc.)
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col -= 1; // Convert to 0-based

  const row = parseInt(rowStr, 10) - 1; // Convert to 0-based

  return { row, col };
}

/**
 * Parse range reference to start/end coordinates
 *
 * @param range - Range reference (e.g., "A1:C10")
 * @returns Start and end coordinates
 */
export function parseRangeRef(range: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} {
  const [start, end] = range.split(':');
  const startCoords = parseCellRef(start);

  if (!end) {
    return {
      startRow: startCoords.row,
      startCol: startCoords.col,
      endRow: startCoords.row,
      endCol: startCoords.col,
    };
  }

  const endCoords = parseCellRef(end);
  return {
    startRow: startCoords.row,
    startCol: startCoords.col,
    endRow: endCoords.row,
    endCol: endCoords.col,
  };
}

/**
 * Extract ZIP file path from relationship target
 * Handles relative paths with ../ and absolute paths
 *
 * @param basePath - Base path (e.g., "xl/")
 * @param target - Relationship target (e.g., "worksheets/sheet1.xml" or "../docProps/core.xml")
 * @returns Normalized path (e.g., "xl/worksheets/sheet1.xml")
 */
export function resolveRelPath(basePath: string, target: string): string {
  // Absolute path
  if (target.startsWith('/')) {
    return target.slice(1);
  }

  // Handle relative navigation
  const baseParts = basePath.split('/').filter(Boolean);
  const targetParts = target.split('/');

  for (const part of targetParts) {
    if (part === '..') {
      baseParts.pop();
    } else if (part !== '.') {
      baseParts.push(part);
    }
  }

  return baseParts.join('/');
}

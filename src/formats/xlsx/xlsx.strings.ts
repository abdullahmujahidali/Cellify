/**
 * Shared Strings Table for XLSX export
 *
 * Excel can store text values in a central shared strings table (xl/sharedStrings.xml)
 * and reference them by index. This reduces file size when the same string appears
 * multiple times across the workbook.
 */

import { XML_DECLARATION, sanitizeXmlString } from './xlsx.xml.js';
import { NS } from './xlsx.types.js';
import type { SharedStringsTable as ISharedStringsTable } from './xlsx.types.js';

/**
 * Shared Strings Table implementation
 *
 * Manages string deduplication and generates the sharedStrings.xml content.
 */
export class SharedStringsTable implements ISharedStringsTable {
  /** Map from string value to its index */
  private strings: Map<string, number> = new Map();

  /** Ordered list of unique strings */
  private stringList: string[] = [];

  /** Total count of string references (including duplicates) */
  private totalCount = 0;

  /**
   * Add a string to the table and return its index
   *
   * If the string already exists, returns the existing index.
   * Otherwise, adds it and returns the new index.
   *
   * @param value - String value to add
   * @returns Index of the string in the shared strings table
   */
  add(value: string): number {
    this.totalCount++;

    let idx = this.strings.get(value);
    if (idx === undefined) {
      idx = this.stringList.length;
      this.stringList.push(value);
      this.strings.set(value, idx);
    }

    return idx;
  }

  /**
   * Get total count of string references (with duplicates)
   * This is used for the "count" attribute in sharedStrings.xml
   */
  get count(): number {
    return this.totalCount;
  }

  /**
   * Get count of unique strings
   * This is used for the "uniqueCount" attribute in sharedStrings.xml
   */
  get uniqueCount(): number {
    return this.stringList.length;
  }

  /**
   * Check if the table is empty
   */
  get isEmpty(): boolean {
    return this.stringList.length === 0;
  }

  /**
   * Generate the sharedStrings.xml content
   *
   * @returns XML string for xl/sharedStrings.xml
   */
  generateXml(): string {
    if (this.isEmpty) {
      return '';
    }

    const parts: string[] = [XML_DECLARATION];

    parts.push(
      `<sst xmlns="${NS.spreadsheetml}" count="${this.totalCount}" uniqueCount="${this.uniqueCount}">`
    );

    for (const str of this.stringList) {
      // Check if string has leading/trailing whitespace that needs preserving
      const needsPreserve = str !== str.trim() || str.includes('\n') || str.includes('\r');

      if (needsPreserve) {
        parts.push(`<si><t xml:space="preserve">${sanitizeXmlString(str)}</t></si>`);
      } else {
        parts.push(`<si><t>${sanitizeXmlString(str)}</t></si>`);
      }
    }

    parts.push('</sst>');

    return parts.join('\n');
  }
}

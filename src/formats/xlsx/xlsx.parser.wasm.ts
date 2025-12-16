/**
 * WASM-accelerated XML parser
 *
 * Provides fast parsing using WebAssembly with automatic fallback to JS.
 * This module wraps the WASM parser and provides a unified API.
 */

import {
  initWasm,
  isWasmAvailable,
  parseWorksheetWasm,
  parseSharedStringsWasm,
  parseStylesWasm,
  parseWorkbookWasm,
  parseRelationshipsWasm,
  type ParsedWorksheet,
  type ParsedStyles,
  type ParsedSheetInfo,
  type ParsedRelationship,
} from './xlsx.wasm.js';

// Re-export types for convenience
export type {
  ParsedWorksheet,
  ParsedStyles,
  ParsedSheetInfo,
  ParsedRelationship,
  ParsedCell,
  ParsedRow,
  ParsedHyperlink,
  ParsedStyle,
  ParsedFont,
  ParsedFill,
  ParsedBorder,
} from './xlsx.wasm.js';

// Module state
let initialized = false;
let initPromise: Promise<boolean> | null = null;

/**
 * Initialize the WASM parser (optional - will auto-init on first use)
 *
 * Call this at application startup for best performance.
 * Returns true if WASM is available, false if falling back to JS.
 *
 * @example
 * ```typescript
 * import { initXlsxWasm } from 'cellify';
 *
 * // At app startup
 * const wasmEnabled = await initXlsxWasm();
 * console.log('WASM parser:', wasmEnabled ? 'enabled' : 'disabled (using JS fallback)');
 * ```
 */
export async function initXlsxWasm(): Promise<boolean> {
  if (initialized) {
    return isWasmAvailable();
  }

  if (initPromise) {
    return initPromise;
  }

  initPromise = initWasm().then((result) => {
    initialized = true;
    return result;
  });

  return initPromise;
}

/**
 * Check if WASM parser is available and initialized
 */
export function isXlsxWasmReady(): boolean {
  return initialized && isWasmAvailable();
}

/**
 * Parse shared strings with WASM acceleration
 * Returns null if WASM is not available (use JS fallback)
 */
export function parseSharedStringsAccelerated(xml: string): string[] | null {
  if (!isXlsxWasmReady()) return null;
  return parseSharedStringsWasm(xml);
}

/**
 * Parse worksheet with WASM acceleration
 * Returns null if WASM is not available (use JS fallback)
 */
export function parseWorksheetAccelerated(xml: string): ParsedWorksheet | null {
  if (!isXlsxWasmReady()) return null;
  return parseWorksheetWasm(xml);
}

/**
 * Parse styles with WASM acceleration
 * Returns null if WASM is not available (use JS fallback)
 */
export function parseStylesAccelerated(xml: string): ParsedStyles | null {
  if (!isXlsxWasmReady()) return null;
  return parseStylesWasm(xml);
}

/**
 * Parse workbook with WASM acceleration
 * Returns null if WASM is not available (use JS fallback)
 */
export function parseWorkbookAccelerated(xml: string): ParsedSheetInfo[] | null {
  if (!isXlsxWasmReady()) return null;
  return parseWorkbookWasm(xml);
}

/**
 * Parse relationships with WASM acceleration
 * Returns null if WASM is not available (use JS fallback)
 */
export function parseRelationshipsAccelerated(xml: string): ParsedRelationship[] | null {
  if (!isXlsxWasmReady()) return null;
  return parseRelationshipsWasm(xml);
}

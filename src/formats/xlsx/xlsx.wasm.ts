/**
 * WASM Parser Wrapper
 *
 * Provides a high-performance XML parsing layer using WebAssembly.
 * Falls back to JavaScript parser when WASM is not available.
 */

// Type definitions for WASM module exports
export interface ParsedCell {
  reference: string;
  cell_type: string | null;
  style_index: number | null;
  value: string | null;
  formula: string | null;
}

export interface ParsedRow {
  row_num: number;
  cells: ParsedCell[];
  height: number | null;
  hidden: boolean;
}

export interface ParsedWorksheet {
  rows: ParsedRow[];
  merge_cells: string[];
  hyperlinks: ParsedHyperlink[];
  col_widths: Record<number, number>;
}

export interface ParsedHyperlink {
  reference: string;
  rid: string | null;
  location: string | null;
  display: string | null;
  tooltip: string | null;
}

export interface ParsedStyle {
  num_fmt_id: number | null;
  font_id: number | null;
  fill_id: number | null;
  border_id: number | null;
  xf_id: number | null;
  apply_number_format: boolean;
  apply_font: boolean;
  apply_fill: boolean;
  apply_border: boolean;
  apply_alignment: boolean;
  horizontal: string | null;
  vertical: string | null;
  wrap_text: boolean;
  text_rotation: number | null;
  indent: number | null;
}

export interface ParsedFont {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  size: number | null;
  color: string | null;
  name: string | null;
}

export interface ParsedFill {
  pattern_type: string | null;
  fg_color: string | null;
  bg_color: string | null;
}

export interface ParsedBorder {
  left_style: string | null;
  left_color: string | null;
  right_style: string | null;
  right_color: string | null;
  top_style: string | null;
  top_color: string | null;
  bottom_style: string | null;
  bottom_color: string | null;
}

export interface ParsedStyles {
  cell_xfs: ParsedStyle[];
  fonts: ParsedFont[];
  fills: ParsedFill[];
  borders: ParsedBorder[];
  num_fmts: Record<number, string>;
}

export interface ParsedSheetInfo {
  name: string;
  sheet_id: number;
  rid: string;
  state: string | null;
}

export interface ParsedRelationship {
  id: string;
  rel_type: string;
  target: string;
  target_mode: string | null;
}

// WASM module interface
interface WasmModule {
  init(): void;
  parse_worksheet(xml: string): ParsedWorksheet;
  parse_shared_strings(xml: string): string[];
  parse_styles(xml: string): ParsedStyles;
  parse_workbook(xml: string): ParsedSheetInfo[];
  parse_relationships(xml: string): ParsedRelationship[];
}

// Module state
let wasmModule: WasmModule | null = null;
let wasmLoadPromise: Promise<boolean> | null = null;
let wasmAvailable = false;

/**
 * Initialize the WASM module
 * Call this once at application startup for best performance
 */
export async function initWasm(): Promise<boolean> {
  // Return cached promise if already loading/loaded
  if (wasmLoadPromise) {
    return wasmLoadPromise;
  }

  wasmLoadPromise = loadWasmModule();
  return wasmLoadPromise;
}

async function loadWasmModule(): Promise<boolean> {
  try {
    // Dynamic import of the WASM module
    // Using explicit path for bundler compatibility
    const module = await import('./wasm/cellify_wasm.js');

    // Initialize the WASM module - this loads the .wasm file
    await module.default();

    if (typeof module.init === 'function') {
      module.init();
    }

    wasmModule = module as unknown as WasmModule;
    wasmAvailable = true;

    console.log('[Cellify] WASM parser loaded successfully');
    return true;
  } catch (error) {
    console.warn('[Cellify] WASM parser not available, using JS fallback:', error);
    wasmAvailable = false;
    return false;
  }
}

/**
 * Check if WASM parser is available
 */
export function isWasmAvailable(): boolean {
  return wasmAvailable;
}

/**
 * Parse worksheet XML using WASM (if available)
 */
export function parseWorksheetWasm(xml: string): ParsedWorksheet | null {
  if (!wasmModule) return null;
  try {
    return wasmModule.parse_worksheet(xml);
  } catch {
    return null;
  }
}

/**
 * Parse shared strings XML using WASM (if available)
 */
export function parseSharedStringsWasm(xml: string): string[] | null {
  if (!wasmModule) return null;
  try {
    return wasmModule.parse_shared_strings(xml);
  } catch {
    return null;
  }
}

/**
 * Parse styles XML using WASM (if available)
 */
export function parseStylesWasm(xml: string): ParsedStyles | null {
  if (!wasmModule) return null;
  try {
    return wasmModule.parse_styles(xml);
  } catch {
    return null;
  }
}

/**
 * Parse workbook XML using WASM (if available)
 */
export function parseWorkbookWasm(xml: string): ParsedSheetInfo[] | null {
  if (!wasmModule) return null;
  try {
    return wasmModule.parse_workbook(xml);
  } catch {
    return null;
  }
}

/**
 * Parse relationships XML using WASM (if available)
 */
export function parseRelationshipsWasm(xml: string): ParsedRelationship[] | null {
  if (!wasmModule) return null;
  try {
    return wasmModule.parse_relationships(xml);
  } catch {
    return null;
  }
}

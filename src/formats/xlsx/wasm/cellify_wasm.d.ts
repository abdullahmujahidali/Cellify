/* tslint:disable */
/* eslint-disable */

/**
 * Initialize the WASM module (call once at startup)
 */
export function init(): void;

/**
 * Parse relationships file (.rels)
 */
export function parse_relationships(xml: string): any;

/**
 * Parse shared strings XML
 */
export function parse_shared_strings(xml: string): any;

/**
 * Parse styles.xml
 */
export function parse_styles(xml: string): any;

/**
 * Parse workbook.xml to get sheet list
 */
export function parse_workbook(xml: string): any;

/**
 * Parse worksheet XML and return structured data
 */
export function parse_worksheet(xml: string): any;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
  readonly memory: WebAssembly.Memory;
  readonly init: () => void;
  readonly parse_relationships: (a: number, b: number) => any;
  readonly parse_shared_strings: (a: number, b: number) => any;
  readonly parse_styles: (a: number, b: number) => any;
  readonly parse_workbook: (a: number, b: number) => any;
  readonly parse_worksheet: (a: number, b: number) => any;
  readonly __wbindgen_free: (a: number, b: number, c: number) => void;
  readonly __wbindgen_malloc: (a: number, b: number) => number;
  readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
  readonly __wbindgen_externrefs: WebAssembly.Table;
  readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
* Instantiates the given `module`, which can either be bytes or
* a precompiled `WebAssembly.Module`.
*
* @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
*
* @returns {InitOutput}
*/
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
* If `module_or_path` is {RequestInfo} or {URL}, makes a request and
* for everything else, calls `WebAssembly.instantiate` directly.
*
* @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
*
* @returns {Promise<InitOutput>}
*/
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;

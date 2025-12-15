import { Cell } from './Cell.js';
import type {
  CellValue,
  MergeRange,
} from '../types/cell.types.js';
import { cellKey, parseKey, a1ToAddress } from '../types/cell.types.js';
import type { CellStyle } from '../types/style.types.js';
import type {
  RangeDefinition,
  ConditionalFormatRule,
  AutoFilter,
} from '../types/range.types.js';
import { parseRangeReference, iterateRange, rangesOverlap } from '../types/range.types.js';

/**
 * Row configuration
 */
export interface RowConfig {
  height?: number; // Height in points
  hidden?: boolean;
  outlineLevel?: number; // For grouping
  style?: CellStyle; // Default style for cells in this row
}

/**
 * Column configuration
 */
export interface ColumnConfig {
  width?: number; // Width in characters
  hidden?: boolean;
  outlineLevel?: number; // For grouping
  style?: CellStyle; // Default style for cells in this column
}

/**
 * Sheet view configuration
 */
export interface SheetView {
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showZeros?: boolean;
  tabSelected?: boolean;
  zoomScale?: number; // 10-400
  // Freeze panes
  frozenRows?: number;
  frozenCols?: number;
  // Split panes
  splitRow?: number;
  splitCol?: number;
}

/**
 * Page setup for printing
 */
export interface PageSetup {
  paperSize?: number;
  orientation?: 'portrait' | 'landscape';
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  margins?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
    header?: number;
    footer?: number;
  };
}

/**
 * Sheet protection options
 */
export interface SheetProtection {
  password?: string;
  sheet?: boolean;
  objects?: boolean;
  scenarios?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  selectLockedCells?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
  selectUnlockedCells?: boolean;
}

/**
 * Represents a worksheet in a workbook.
 *
 * Sheets contain cells organized in rows and columns. They support:
 * - Cell access by address (A1 notation) or row/col indices
 * - Merged cells
 * - Row and column configuration (height, width, hidden)
 * - Conditional formatting
 * - Auto filters
 * - Freeze panes
 * - Sheet protection
 */
export class Sheet {
  private _name: string;
  private _cells: Map<string, Cell> = new Map();
  private _merges: MergeRange[] = [];
  private _rows: Map<number, RowConfig> = new Map();
  private _cols: Map<number, ColumnConfig> = new Map();
  private _conditionalFormats: ConditionalFormatRule[] = [];
  private _autoFilter: AutoFilter | undefined;
  private _view: SheetView = {};
  private _pageSetup: PageSetup = {};
  private _protection: SheetProtection | undefined;

  // Track dimensions for efficient iteration
  private _minRow = Infinity;
  private _maxRow = -Infinity;
  private _minCol = Infinity;
  private _maxCol = -Infinity;

  constructor(name: string) {
    this._name = name;
  }

  /**
   * Get the sheet name
   */
  get name(): string {
    return this._name;
  }

  /**
   * Set the sheet name
   */
  set name(value: string) {
    this._name = value;
  }

  /**
   * Get cell by A1 notation (e.g., "A1", "B2")
   * Creates the cell if it doesn't exist
   */
  cell(address: string): Cell;

  /**
   * Get cell by row and column indices (0-based)
   * Creates the cell if it doesn't exist
   */
  cell(row: number, col: number): Cell;

  cell(addressOrRow: string | number, col?: number): Cell {
    let row: number;
    let column: number;

    if (typeof addressOrRow === 'string') {
      const addr = a1ToAddress(addressOrRow);
      row = addr.row;
      column = addr.col;
    } else {
      row = addressOrRow;
      column = col!;
    }

    const key = cellKey(row, column);
    let cell = this._cells.get(key);

    if (!cell) {
      cell = new Cell(row, column);
      this._cells.set(key, cell);
      this.updateDimensions(row, column);
    }

    return cell;
  }

  /**
   * Get cell if it exists, without creating it
   */
  getCell(address: string): Cell | undefined;
  getCell(row: number, col: number): Cell | undefined;
  getCell(addressOrRow: string | number, col?: number): Cell | undefined {
    let row: number;
    let column: number;

    if (typeof addressOrRow === 'string') {
      const addr = a1ToAddress(addressOrRow);
      row = addr.row;
      column = addr.col;
    } else {
      row = addressOrRow;
      column = col!;
    }

    return this._cells.get(cellKey(row, column));
  }

  /**
   * Check if a cell exists
   */
  hasCell(address: string): boolean;
  hasCell(row: number, col: number): boolean;
  hasCell(addressOrRow: string | number, col?: number): boolean {
    let row: number;
    let column: number;

    if (typeof addressOrRow === 'string') {
      const addr = a1ToAddress(addressOrRow);
      row = addr.row;
      column = addr.col;
    } else {
      row = addressOrRow;
      column = col!;
    }

    return this._cells.has(cellKey(row, column));
  }

  /**
   * Delete a cell
   */
  deleteCell(address: string): boolean;
  deleteCell(row: number, col: number): boolean;
  deleteCell(addressOrRow: string | number, col?: number): boolean {
    let row: number;
    let column: number;

    if (typeof addressOrRow === 'string') {
      const addr = a1ToAddress(addressOrRow);
      row = addr.row;
      column = addr.col;
    } else {
      row = addressOrRow;
      column = col!;
    }

    const key = cellKey(row, column);
    const deleted = this._cells.delete(key);

    if (deleted) {
      this.recalculateDimensions();
    }

    return deleted;
  }

  /**
   * Update dimension tracking
   */
  private updateDimensions(row: number, col: number): void {
    this._minRow = Math.min(this._minRow, row);
    this._maxRow = Math.max(this._maxRow, row);
    this._minCol = Math.min(this._minCol, col);
    this._maxCol = Math.max(this._maxCol, col);
  }

  /**
   * Recalculate dimensions after cell deletion
   */
  private recalculateDimensions(): void {
    this._minRow = Infinity;
    this._maxRow = -Infinity;
    this._minCol = Infinity;
    this._maxCol = -Infinity;

    for (const key of this._cells.keys()) {
      const { row, col } = parseKey(key);
      this.updateDimensions(row, col);
    }
  }

  /**
   * Get the used range of the sheet
   */
  get dimensions(): RangeDefinition | null {
    if (this._cells.size === 0) {
      return null;
    }

    return {
      startRow: this._minRow,
      startCol: this._minCol,
      endRow: this._maxRow,
      endCol: this._maxCol,
    };
  }

  /**
   * Get the number of rows with data
   */
  get rowCount(): number {
    if (this._cells.size === 0) return 0;
    return this._maxRow - this._minRow + 1;
  }

  /**
   * Get the number of columns with data
   */
  get columnCount(): number {
    if (this._cells.size === 0) return 0;
    return this._maxCol - this._minCol + 1;
  }

  /**
   * Get the total number of cells with data
   */
  get cellCount(): number {
    return this._cells.size;
  }

  /**
   * Iterate over all cells
   */
  *cells(): Generator<Cell> {
    for (const cell of this._cells.values()) {
      yield cell;
    }
  }

  /**
   * Iterate over cells in a range
   */
  *cellsInRange(range: string | RangeDefinition): Generator<Cell> {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;

    for (const { row, col } of iterateRange(rangeDef)) {
      const cell = this.getCell(row, col);
      if (cell) {
        yield cell;
      }
    }
  }

  /**
   * Set values for a range from a 2D array
   */
  setValues(startAddress: string, values: CellValue[][]): this;
  setValues(startRow: number, startCol: number, values: CellValue[][]): this;
  setValues(
    startAddressOrRow: string | number,
    startColOrValues: number | CellValue[][],
    valuesArg?: CellValue[][]
  ): this {
    let startRow: number;
    let startCol: number;
    let values: CellValue[][];

    if (typeof startAddressOrRow === 'string') {
      const addr = a1ToAddress(startAddressOrRow);
      startRow = addr.row;
      startCol = addr.col;
      values = startColOrValues as CellValue[][];
    } else {
      startRow = startAddressOrRow;
      startCol = startColOrValues as number;
      values = valuesArg!;
    }

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      for (let c = 0; c < row.length; c++) {
        this.cell(startRow + r, startCol + c).value = row[c];
      }
    }

    return this;
  }

  /**
   * Get values from a range as a 2D array
   */
  getValues(range: string | RangeDefinition): CellValue[][] {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;
    const result: CellValue[][] = [];

    for (let row = rangeDef.startRow; row <= rangeDef.endRow; row++) {
      const rowData: CellValue[] = [];
      for (let col = rangeDef.startCol; col <= rangeDef.endCol; col++) {
        const cell = this.getCell(row, col);
        rowData.push(cell?.value ?? null);
      }
      result.push(rowData);
    }

    return result;
  }

  // ============ Merge Operations ============

  /**
   * Merge cells in a range
   */
  mergeCells(range: string | RangeDefinition): this {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;

    // Check for overlapping merges
    for (const existing of this._merges) {
      if (rangesOverlap(existing, rangeDef)) {
        throw new Error('Cannot merge cells: overlaps with existing merge');
      }
    }

    // Add merge
    this._merges.push(rangeDef);

    // Set merge info on master cell
    const masterCell = this.cell(rangeDef.startRow, rangeDef.startCol);
    masterCell._setMerge(rangeDef);

    // Mark slave cells
    for (const { row, col } of iterateRange(rangeDef)) {
      if (row === rangeDef.startRow && col === rangeDef.startCol) continue;

      const cell = this.cell(row, col);
      cell._setMergedInto({ row: rangeDef.startRow, col: rangeDef.startCol });
    }

    return this;
  }

  /**
   * Unmerge cells
   */
  unmergeCells(range: string | RangeDefinition): this {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;

    const index = this._merges.findIndex(
      (m) =>
        m.startRow === rangeDef.startRow &&
        m.startCol === rangeDef.startCol &&
        m.endRow === rangeDef.endRow &&
        m.endCol === rangeDef.endCol
    );

    if (index === -1) {
      throw new Error('No merge found at specified range');
    }

    // Remove merge
    this._merges.splice(index, 1);

    // Clear merge info on master cell
    const masterCell = this.getCell(rangeDef.startRow, rangeDef.startCol);
    if (masterCell) {
      masterCell._setMerge(undefined);
    }

    // Clear slave cell references
    for (const { row, col } of iterateRange(rangeDef)) {
      if (row === rangeDef.startRow && col === rangeDef.startCol) continue;

      const cell = this.getCell(row, col);
      if (cell) {
        cell._setMergedInto(undefined);
      }
    }

    return this;
  }

  /**
   * Get all merge ranges
   */
  get merges(): readonly MergeRange[] {
    return this._merges;
  }

  // ============ Row/Column Configuration ============

  /**
   * Get row configuration
   */
  getRow(index: number): RowConfig {
    return this._rows.get(index) ?? {};
  }

  /**
   * Set row configuration
   */
  setRow(index: number, config: RowConfig): this {
    const existing = this._rows.get(index) ?? {};
    this._rows.set(index, { ...existing, ...config });
    return this;
  }

  /**
   * Set row height
   */
  setRowHeight(index: number, height: number): this {
    return this.setRow(index, { height });
  }

  /**
   * Hide a row
   */
  hideRow(index: number): this {
    return this.setRow(index, { hidden: true });
  }

  /**
   * Show a hidden row
   */
  showRow(index: number): this {
    return this.setRow(index, { hidden: false });
  }

  /**
   * Get column configuration
   */
  getColumn(index: number): ColumnConfig {
    return this._cols.get(index) ?? {};
  }

  /**
   * Set column configuration
   */
  setColumn(index: number, config: ColumnConfig): this {
    const existing = this._cols.get(index) ?? {};
    this._cols.set(index, { ...existing, ...config });
    return this;
  }

  /**
   * Set column width
   */
  setColumnWidth(index: number, width: number): this {
    return this.setColumn(index, { width });
  }

  /**
   * Hide a column
   */
  hideColumn(index: number): this {
    return this.setColumn(index, { hidden: true });
  }

  /**
   * Show a hidden column
   */
  showColumn(index: number): this {
    return this.setColumn(index, { hidden: false });
  }

  /**
   * Get all row configurations
   */
  get rows(): ReadonlyMap<number, RowConfig> {
    return this._rows;
  }

  /**
   * Get all column configurations
   */
  get columns(): ReadonlyMap<number, ColumnConfig> {
    return this._cols;
  }

  // ============ View Configuration ============

  /**
   * Get sheet view configuration
   */
  get view(): SheetView {
    return this._view;
  }

  /**
   * Set sheet view configuration
   */
  setView(view: Partial<SheetView>): this {
    this._view = { ...this._view, ...view };
    return this;
  }

  /**
   * Freeze rows and columns
   */
  freeze(rows: number, cols: number = 0): this {
    this._view.frozenRows = rows;
    this._view.frozenCols = cols;
    return this;
  }

  /**
   * Remove freeze panes
   */
  unfreeze(): this {
    this._view.frozenRows = undefined;
    this._view.frozenCols = undefined;
    return this;
  }

  // ============ Auto Filter ============

  /**
   * Set auto filter on a range
   */
  setAutoFilter(range: string | RangeDefinition): this {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;
    this._autoFilter = { range: rangeDef };
    return this;
  }

  /**
   * Remove auto filter
   */
  removeAutoFilter(): this {
    this._autoFilter = undefined;
    return this;
  }

  /**
   * Get auto filter configuration
   */
  get autoFilter(): AutoFilter | undefined {
    return this._autoFilter;
  }

  // ============ Conditional Formatting ============

  /**
   * Add a conditional formatting rule
   */
  addConditionalFormat(rule: ConditionalFormatRule): this {
    this._conditionalFormats.push(rule);
    return this;
  }

  /**
   * Get all conditional formatting rules
   */
  get conditionalFormats(): readonly ConditionalFormatRule[] {
    return this._conditionalFormats;
  }

  /**
   * Remove all conditional formatting rules
   */
  clearConditionalFormats(): this {
    this._conditionalFormats = [];
    return this;
  }

  // ============ Protection ============

  /**
   * Protect the sheet
   */
  protect(options: SheetProtection = {}): this {
    this._protection = { sheet: true, ...options };
    return this;
  }

  /**
   * Unprotect the sheet
   */
  unprotect(): this {
    this._protection = undefined;
    return this;
  }

  /**
   * Get protection settings
   */
  get protection(): SheetProtection | undefined {
    return this._protection;
  }

  /**
   * Check if sheet is protected
   */
  get isProtected(): boolean {
    return this._protection?.sheet === true;
  }

  // ============ Page Setup ============

  /**
   * Get page setup configuration
   */
  get pageSetup(): PageSetup {
    return this._pageSetup;
  }

  /**
   * Set page setup configuration
   */
  setPageSetup(setup: Partial<PageSetup>): this {
    this._pageSetup = { ...this._pageSetup, ...setup };
    return this;
  }

  // ============ Utility Methods ============

  /**
   * Apply a style to a range of cells
   */
  applyStyle(range: string | RangeDefinition, style: CellStyle): this {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;

    for (const { row, col } of iterateRange(rangeDef)) {
      this.cell(row, col).applyStyle(style);
    }

    return this;
  }

  /**
   * Clear all cells in a range
   */
  clearRange(range: string | RangeDefinition): this {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;

    for (const { row, col } of iterateRange(rangeDef)) {
      const cell = this.getCell(row, col);
      if (cell) {
        cell.clear();
      }
    }

    return this;
  }

  /**
   * Convert sheet to JSON
   */
  toJSON(): Record<string, unknown> {
    const cells: Record<string, unknown>[] = [];
    for (const cell of this._cells.values()) {
      if (!cell.isEmpty) {
        cells.push(cell.toJSON());
      }
    }

    return {
      name: this._name,
      dimensions: this.dimensions,
      cells,
      merges: this._merges,
      rows: Object.fromEntries(this._rows),
      columns: Object.fromEntries(this._cols),
      view: this._view,
      pageSetup: this._pageSetup,
      protection: this._protection,
      autoFilter: this._autoFilter,
      conditionalFormats: this._conditionalFormats,
    };
  }
}

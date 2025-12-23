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
  FilterCriteria,
  SearchOptions,
  PasteOptions,
} from '../types/range.types.js';
import { parseRangeReference, iterateRange, rangesOverlap } from '../types/range.types.js';
import type {
  SheetEventMap,
  SheetEventHandler,
  ChangeRecord,
  CellChangeEvent,
  CellStyleChangeEvent,
  CellAddedEvent,
  CellDeletedEvent,
} from '../types/event.types.js';

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

  // Event system
  private _eventListeners: Map<string, Set<SheetEventHandler>> = new Map();
  private _changes: ChangeRecord[] = [];
  private _changeIdCounter = 0;
  private _eventsEnabled = true;

  // Undo/Redo system
  private _undoStack: ChangeRecord[] = [];
  private _redoStack: ChangeRecord[] = [];
  private _maxUndoHistory = 100;
  private _isUndoRedoOperation = false;

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
      // Set up change callback
      cell._onChange = this.handleCellChange.bind(this);
      // Emit cell added event
      this.emitCellAdded(cell);
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
    const cell = this._cells.get(key);

    if (cell) {
      this.emitCellDeleted(cell);
      this._cells.delete(key);
      this.recalculateDimensions();
      return true;
    }

    return false;
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

  // ============ Sorting ============

  /**
   * Sort rows by the values in a column
   *
   * @param column - Column index (0-based) or letter (e.g., 'A') to sort by
   * @param options - Sort options
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * // Sort by column A ascending
   * sheet.sort('A');
   *
   * // Sort by column B descending
   * sheet.sort('B', { descending: true });
   *
   * // Sort with header row (don't move first row)
   * sheet.sort('A', { hasHeader: true });
   *
   * // Sort specific range
   * sheet.sort('A', { range: 'A1:C10' });
   * ```
   */
  sort(
    column: number | string,
    options: {
      descending?: boolean;
      hasHeader?: boolean;
      range?: string | RangeDefinition;
      numeric?: boolean;
      caseSensitive?: boolean;
    } = {}
  ): this {
    const {
      descending = false,
      hasHeader = false,
      range,
      numeric = false,
      caseSensitive = false,
    } = options;

    // Convert column letter to index if needed
    const colIndex = typeof column === 'string'
      ? this.columnLetterToIndex(column)
      : column;

    // Determine range to sort
    let sortRange: RangeDefinition;
    if (range) {
      sortRange = typeof range === 'string' ? parseRangeReference(range) : range;
    } else {
      const dims = this.dimensions;
      if (!dims) return this;
      sortRange = dims;
    }

    // Adjust for header row
    const startRow = hasHeader ? sortRange.startRow + 1 : sortRange.startRow;
    const endRow = sortRange.endRow;

    if (startRow > endRow) return this;

    // Collect rows with their data
    const rows: { rowIndex: number; sortValue: CellValue; cells: Map<number, Cell> }[] = [];

    for (let r = startRow; r <= endRow; r++) {
      const sortCell = this.getCell(r, colIndex);
      const sortValue = sortCell?.value ?? null;

      const cells = new Map<number, Cell>();
      for (let c = sortRange.startCol; c <= sortRange.endCol; c++) {
        const cell = this.getCell(r, c);
        if (cell) {
          cells.set(c, cell.clone());
        }
      }

      rows.push({ rowIndex: r, sortValue, cells });
    }

    // Sort rows
    rows.sort((a, b) => {
      const aVal = a.sortValue;
      const bVal = b.sortValue;

      let result = this.compareValues(aVal, bVal, numeric, caseSensitive);
      return descending ? -result : result;
    });

    // Disable events during reordering
    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;

    try {
      // Clear existing cells in range
      for (let r = startRow; r <= endRow; r++) {
        for (let c = sortRange.startCol; c <= sortRange.endCol; c++) {
          const key = cellKey(r, c);
          this._cells.delete(key);
        }
      }

      // Place sorted rows
      for (let i = 0; i < rows.length; i++) {
        const targetRow = startRow + i;
        const { cells } = rows[i];

        for (const [col, cell] of cells) {
          const newCell = this.createCellFromSource(targetRow, col, cell);
          this._cells.set(cellKey(targetRow, col), newCell);
        }
      }

      this.recalculateDimensions();
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  /**
   * Sort rows by multiple columns
   *
   * @param columns - Array of column sort specifications
   * @param options - Sort options
   *
   * @example
   * ```typescript
   * // Sort by column A, then by column B descending
   * sheet.sortBy([
   *   { column: 'A' },
   *   { column: 'B', descending: true }
   * ]);
   * ```
   */
  sortBy(
    columns: Array<{
      column: number | string;
      descending?: boolean;
      numeric?: boolean;
    }>,
    options: {
      hasHeader?: boolean;
      range?: string | RangeDefinition;
      caseSensitive?: boolean;
    } = {}
  ): this {
    const { hasHeader = false, range, caseSensitive = false } = options;

    // Convert column letters to indices
    const sortColumns = columns.map((c) => ({
      colIndex: typeof c.column === 'string' ? this.columnLetterToIndex(c.column) : c.column,
      descending: c.descending ?? false,
      numeric: c.numeric ?? false,
    }));

    // Determine range to sort
    let sortRange: RangeDefinition;
    if (range) {
      sortRange = typeof range === 'string' ? parseRangeReference(range) : range;
    } else {
      const dims = this.dimensions;
      if (!dims) return this;
      sortRange = dims;
    }

    const startRow = hasHeader ? sortRange.startRow + 1 : sortRange.startRow;
    const endRow = sortRange.endRow;

    if (startRow > endRow) return this;

    // Collect rows
    const rows: { rowIndex: number; sortValues: CellValue[]; cells: Map<number, Cell> }[] = [];

    for (let r = startRow; r <= endRow; r++) {
      const sortValues = sortColumns.map((sc) => {
        const cell = this.getCell(r, sc.colIndex);
        return cell?.value ?? null;
      });

      const cells = new Map<number, Cell>();
      for (let c = sortRange.startCol; c <= sortRange.endCol; c++) {
        const cell = this.getCell(r, c);
        if (cell) {
          cells.set(c, cell.clone());
        }
      }

      rows.push({ rowIndex: r, sortValues, cells });
    }

    // Sort with multi-column comparison
    rows.sort((a, b) => {
      for (let i = 0; i < sortColumns.length; i++) {
        const aVal = a.sortValues[i];
        const bVal = b.sortValues[i];
        const { descending, numeric } = sortColumns[i];

        const result = this.compareValues(aVal, bVal, numeric, caseSensitive);
        if (result !== 0) {
          return descending ? -result : result;
        }
      }
      return 0;
    });

    // Disable events and reorder
    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;

    try {
      for (let r = startRow; r <= endRow; r++) {
        for (let c = sortRange.startCol; c <= sortRange.endCol; c++) {
          this._cells.delete(cellKey(r, c));
        }
      }

      for (let i = 0; i < rows.length; i++) {
        const targetRow = startRow + i;
        const { cells } = rows[i];

        for (const [col, cell] of cells) {
          const newCell = this.createCellFromSource(targetRow, col, cell);
          this._cells.set(cellKey(targetRow, col), newCell);
        }
      }

      this.recalculateDimensions();
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  /**
   * Compare two cell values for sorting
   */
  private compareValues(a: CellValue, b: CellValue, numeric: boolean, caseSensitive: boolean): number {
    // Handle nulls - always sort to end
    if (a === null && b === null) return 0;
    if (a === null) return 1;
    if (b === null) return -1;

    // Numeric comparison
    if (numeric || (typeof a === 'number' && typeof b === 'number')) {
      const numA = typeof a === 'number' ? a : parseFloat(String(a));
      const numB = typeof b === 'number' ? b : parseFloat(String(b));

      if (!isNaN(numA) && !isNaN(numB)) {
        return numA - numB;
      }
    }

    // Date comparison
    if (a instanceof Date && b instanceof Date) {
      return a.getTime() - b.getTime();
    }

    // String comparison
    const strA = String(a);
    const strB = String(b);

    if (caseSensitive) {
      return strA.localeCompare(strB);
    }
    return strA.toLowerCase().localeCompare(strB.toLowerCase());
  }

  /**
   * Convert column letter to index
   */
  private columnLetterToIndex(letter: string): number {
    let index = 0;
    const upper = letter.toUpperCase();
    for (let i = 0; i < upper.length; i++) {
      index = index * 26 + (upper.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  // ============ Filtering ============

  // Track filtered (hidden) rows
  private _filteredRows: Set<number> = new Set();
  private _activeFilters: Map<number, FilterCriteria> = new Map();

  /**
   * Filter rows based on column values
   *
   * @param column - Column index (0-based) or letter (e.g., 'A') to filter by
   * @param criteria - Filter criteria
   * @param options - Filter options
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * // Show only rows where column A equals 'Active'
   * sheet.filter('A', { equals: 'Active' });
   *
   * // Show rows where column B contains 'test'
   * sheet.filter('B', { contains: 'test' });
   *
   * // Show rows where column C is greater than 100
   * sheet.filter('C', { greaterThan: 100 });
   *
   * // Custom filter function
   * sheet.filter('D', { custom: (value) => value !== null && value > 0 });
   * ```
   */
  filter(
    column: number | string,
    criteria: FilterCriteria,
    options: {
      hasHeader?: boolean;
      range?: string | RangeDefinition;
    } = {}
  ): this {
    const { hasHeader = false, range } = options;

    const colIndex = typeof column === 'string'
      ? this.columnLetterToIndex(column)
      : column;

    // Store the active filter
    this._activeFilters.set(colIndex, criteria);

    // Apply all filters
    this.applyFilters(hasHeader, range);

    return this;
  }

  /**
   * Filter rows based on multiple column criteria
   *
   * @param filters - Array of column filter specifications
   * @param options - Filter options
   *
   * @example
   * ```typescript
   * sheet.filterBy([
   *   { column: 'A', criteria: { equals: 'Active' } },
   *   { column: 'B', criteria: { greaterThan: 100 } }
   * ]);
   * ```
   */
  filterBy(
    filters: Array<{
      column: number | string;
      criteria: FilterCriteria;
    }>,
    options: {
      hasHeader?: boolean;
      range?: string | RangeDefinition;
    } = {}
  ): this {
    const { hasHeader = false, range } = options;

    // Store all filters
    for (const filter of filters) {
      const colIndex = typeof filter.column === 'string'
        ? this.columnLetterToIndex(filter.column)
        : filter.column;
      this._activeFilters.set(colIndex, filter.criteria);
    }

    // Apply all filters
    this.applyFilters(hasHeader, range);

    return this;
  }

  /**
   * Clear all filters and show all rows
   */
  clearFilter(): this {
    this._activeFilters.clear();

    // Show all filtered rows
    for (const row of this._filteredRows) {
      this.showRow(row);
    }
    this._filteredRows.clear();

    return this;
  }

  /**
   * Clear filter on a specific column
   */
  clearColumnFilter(column: number | string): this {
    const colIndex = typeof column === 'string'
      ? this.columnLetterToIndex(column)
      : column;

    this._activeFilters.delete(colIndex);

    // Re-apply remaining filters
    if (this._activeFilters.size > 0) {
      // Show all rows first, then re-apply filters
      for (const row of this._filteredRows) {
        this.showRow(row);
      }
      this._filteredRows.clear();
      this.applyFilters(false);
    } else {
      // No more filters, show all rows
      this.clearFilter();
    }

    return this;
  }

  /**
   * Get active filters
   */
  get activeFilters(): ReadonlyMap<number, FilterCriteria> {
    return this._activeFilters;
  }

  /**
   * Check if a row is currently filtered out (hidden by filter)
   */
  isRowFiltered(row: number): boolean {
    return this._filteredRows.has(row);
  }

  /**
   * Get all filtered row indices
   */
  get filteredRows(): ReadonlySet<number> {
    return this._filteredRows;
  }

  /**
   * Internal: Apply all active filters to the sheet
   */
  private applyFilters(hasHeader: boolean, range?: string | RangeDefinition): void {
    // Determine range to filter
    let filterRange: RangeDefinition;
    if (range) {
      filterRange = typeof range === 'string' ? parseRangeReference(range) : range;
    } else {
      const dims = this.dimensions;
      if (!dims) return;
      filterRange = dims;
    }

    const startRow = hasHeader ? filterRange.startRow + 1 : filterRange.startRow;
    const endRow = filterRange.endRow;

    // Show all rows first
    for (const row of this._filteredRows) {
      this.showRow(row);
    }
    this._filteredRows.clear();

    // Apply each filter
    for (let r = startRow; r <= endRow; r++) {
      let matches = true;

      for (const [colIndex, criteria] of this._activeFilters) {
        const cell = this.getCell(r, colIndex);
        const value = cell?.value ?? null;

        if (!this.matchesCriteria(value, criteria)) {
          matches = false;
          break;
        }
      }

      if (!matches) {
        this.hideRow(r);
        this._filteredRows.add(r);
      }
    }
  }

  /**
   * Internal: Check if a value matches filter criteria
   */
  private matchesCriteria(value: CellValue, criteria: FilterCriteria): boolean {
    // Custom function takes precedence
    if (criteria.custom) {
      return criteria.custom(value);
    }

    // isEmpty
    if (criteria.isEmpty !== undefined) {
      const isEmpty = value === null || value === undefined || value === '';
      return criteria.isEmpty ? isEmpty : !isEmpty;
    }

    // isNotEmpty
    if (criteria.isNotEmpty !== undefined) {
      const isEmpty = value === null || value === undefined || value === '';
      return criteria.isNotEmpty ? !isEmpty : isEmpty;
    }

    // equals
    if (criteria.equals !== undefined) {
      if (typeof criteria.equals === 'string' && typeof value === 'string') {
        return value.toLowerCase() === criteria.equals.toLowerCase();
      }
      return value === criteria.equals;
    }

    // notEquals
    if (criteria.notEquals !== undefined) {
      if (typeof criteria.notEquals === 'string' && typeof value === 'string') {
        return value.toLowerCase() !== criteria.notEquals.toLowerCase();
      }
      return value !== criteria.notEquals;
    }

    // String operations
    if (typeof value === 'string') {
      const lowerValue = value.toLowerCase();

      if (criteria.contains !== undefined) {
        return lowerValue.includes(criteria.contains.toLowerCase());
      }

      if (criteria.notContains !== undefined) {
        return !lowerValue.includes(criteria.notContains.toLowerCase());
      }

      if (criteria.startsWith !== undefined) {
        return lowerValue.startsWith(criteria.startsWith.toLowerCase());
      }

      if (criteria.endsWith !== undefined) {
        return lowerValue.endsWith(criteria.endsWith.toLowerCase());
      }
    }

    // Numeric operations
    const numValue = typeof value === 'number' ? value : parseFloat(String(value));
    if (!isNaN(numValue)) {
      if (criteria.greaterThan !== undefined) {
        return numValue > criteria.greaterThan;
      }

      if (criteria.greaterThanOrEqual !== undefined) {
        return numValue >= criteria.greaterThanOrEqual;
      }

      if (criteria.lessThan !== undefined) {
        return numValue < criteria.lessThan;
      }

      if (criteria.lessThanOrEqual !== undefined) {
        return numValue <= criteria.lessThanOrEqual;
      }

      if (criteria.between !== undefined) {
        const [min, max] = criteria.between;
        return numValue >= min && numValue <= max;
      }

      if (criteria.notBetween !== undefined) {
        const [min, max] = criteria.notBetween;
        return numValue < min || numValue > max;
      }
    }

    // in / notIn (value list)
    if (criteria.in !== undefined) {
      return criteria.in.some((v: string | number | boolean | null) => {
        if (typeof v === 'string' && typeof value === 'string') {
          return v.toLowerCase() === value.toLowerCase();
        }
        return v === value;
      });
    }

    if (criteria.notIn !== undefined) {
      return !criteria.notIn.some((v: string | number | boolean | null) => {
        if (typeof v === 'string' && typeof value === 'string') {
          return v.toLowerCase() === value.toLowerCase();
        }
        return v === value;
      });
    }

    // Default: if no criteria matched, include the row
    return true;
  }

  // ============ Search ============

  /**
   * Find the first cell matching the search criteria
   *
   * @param query - String, number, RegExp, or search options
   * @param options - Search options
   * @returns The first matching cell, or undefined if not found
   *
   * @example
   * ```typescript
   * // Find by exact value
   * const cell = sheet.find('Hello');
   *
   * // Find by regex
   * const cell = sheet.find(/error/i);
   *
   * // Find with options
   * const cell = sheet.find('test', { matchCase: true, searchIn: 'values' });
   * ```
   */
  find(
    query: string | number | RegExp | SearchOptions,
    options: SearchOptions = {}
  ): Cell | undefined {
    const results = this.findAllInternal(query, options, 1);
    return results[0];
  }

  /**
   * Find all cells matching the search criteria
   *
   * @param query - String, number, RegExp, or search options
   * @param options - Search options
   * @returns Array of matching cells
   *
   * @example
   * ```typescript
   * // Find all cells containing 'error'
   * const cells = sheet.findAll('error');
   *
   * // Find all numbers greater than 100 using regex
   * const cells = sheet.findAll(/^\d{3,}$/);
   *
   * // Search in formulas
   * const cells = sheet.findAll('SUM', { searchIn: 'formulas' });
   * ```
   */
  findAll(
    query: string | number | RegExp | SearchOptions,
    options: SearchOptions = {}
  ): Cell[] {
    return this.findAllInternal(query, options);
  }

  /**
   * Replace the first occurrence of a value
   *
   * @param search - Value to search for
   * @param replacement - Value to replace with
   * @param options - Search options
   * @returns The replaced cell, or undefined if not found
   */
  replace(
    search: string | number | RegExp,
    replacement: string | number,
    options: SearchOptions = {}
  ): Cell | undefined {
    const cell = this.find(search, options);
    if (cell) {
      this.replaceInCell(cell, search, replacement);
    }
    return cell;
  }

  /**
   * Replace all occurrences of a value
   *
   * @param search - Value to search for
   * @param replacement - Value to replace with
   * @param options - Search options
   * @returns Array of replaced cells
   */
  replaceAll(
    search: string | number | RegExp,
    replacement: string | number,
    options: SearchOptions = {}
  ): Cell[] {
    const cells = this.findAll(search, options);
    for (const cell of cells) {
      this.replaceInCell(cell, search, replacement);
    }
    return cells;
  }

  /**
   * Internal: Find all matching cells with optional limit
   */
  private findAllInternal(
    query: string | number | RegExp | SearchOptions,
    options: SearchOptions = {},
    limit?: number
  ): Cell[] {
    // Normalize query to options
    let searchOptions: SearchOptions;
    if (typeof query === 'string' || typeof query === 'number') {
      searchOptions = { ...options, query };
    } else if (query instanceof RegExp) {
      searchOptions = { ...options, regex: query };
    } else {
      searchOptions = { ...query, ...options };
    }

    const {
      query: searchQuery,
      regex,
      matchCase = false,
      matchCell = false,
      searchIn = 'values',
      range,
    } = searchOptions;

    const results: Cell[] = [];

    // Determine search range
    let searchRange: RangeDefinition | null = null;
    if (range) {
      searchRange = typeof range === 'string' ? parseRangeReference(range) : range;
    }

    for (const cell of this._cells.values()) {
      // Check if cell is in range
      if (searchRange) {
        if (
          cell.row < searchRange.startRow ||
          cell.row > searchRange.endRow ||
          cell.col < searchRange.startCol ||
          cell.col > searchRange.endCol
        ) {
          continue;
        }
      }

      // Get value to search in
      let searchValue: string | null = null;

      if (searchIn === 'values' || searchIn === 'both') {
        const val = cell.value;
        if (val !== null && val !== undefined) {
          searchValue = String(val);
        }
      }

      if (searchIn === 'formulas' || searchIn === 'both') {
        if (cell.formula) {
          const formulaStr = cell.formula.formula;
          searchValue = searchValue ? `${searchValue} ${formulaStr}` : formulaStr;
        }
      }

      if (searchValue === null) continue;

      // Perform search
      let matches = false;

      if (regex) {
        matches = regex.test(searchValue);
      } else if (searchQuery !== undefined) {
        const queryStr = String(searchQuery);
        const targetStr = matchCase ? searchValue : searchValue.toLowerCase();
        const searchStr = matchCase ? queryStr : queryStr.toLowerCase();

        if (matchCell) {
          matches = targetStr === searchStr;
        } else {
          matches = targetStr.includes(searchStr);
        }
      }

      if (matches) {
        results.push(cell);
        if (limit && results.length >= limit) {
          break;
        }
      }
    }

    return results;
  }

  /**
   * Internal: Replace value in a cell
   */
  private replaceInCell(
    cell: Cell,
    search: string | number | RegExp,
    replacement: string | number
  ): void {
    const currentValue = cell.value;
    if (currentValue === null || currentValue === undefined) return;

    if (typeof currentValue === 'string') {
      if (search instanceof RegExp) {
        cell.value = currentValue.replace(search, String(replacement));
      } else {
        cell.value = currentValue.replace(String(search), String(replacement));
      }
    } else if (typeof currentValue === 'number' && typeof search === 'number') {
      if (currentValue === search) {
        cell.value = typeof replacement === 'number' ? replacement : parseFloat(String(replacement));
      }
    }
  }

  // ============ Copy/Paste ============

  // Internal clipboard for copy/paste operations
  private _clipboard: {
    cells: Map<string, { value: CellValue; style?: CellStyle; formula?: string }>;
    range: RangeDefinition;
  } | null = null;

  /**
   * Copy a range of cells to the internal clipboard
   *
   * @param range - Range to copy (e.g., 'A1:C3' or RangeDefinition)
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * sheet.copyRange('A1:C3');
   * sheet.pasteRange('E1'); // Paste at E1
   * ```
   */
  copyRange(range: string | RangeDefinition): this {
    const rangeDef = typeof range === 'string' ? parseRangeReference(range) : range;

    const cells = new Map<string, { value: CellValue; style?: CellStyle; formula?: string }>();

    for (let row = rangeDef.startRow; row <= rangeDef.endRow; row++) {
      for (let col = rangeDef.startCol; col <= rangeDef.endCol; col++) {
        const cell = this.getCell(row, col);
        if (cell) {
          // Store relative position within the range
          const relKey = `${row - rangeDef.startRow},${col - rangeDef.startCol}`;
          cells.set(relKey, {
            value: cell.value,
            style: cell.style ? { ...cell.style } : undefined,
            formula: cell.formula?.formula,
          });
        }
      }
    }

    this._clipboard = { cells, range: rangeDef };
    return this;
  }

  /**
   * Cut a range of cells (copy and then clear)
   *
   * @param range - Range to cut
   * @returns this for chaining
   */
  cutRange(range: string | RangeDefinition): this {
    this.copyRange(range);
    this.clearRange(range);
    return this;
  }

  /**
   * Paste the clipboard contents at the specified location
   *
   * @param target - Target cell address or position (top-left corner of paste area)
   * @param options - Paste options
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * sheet.copyRange('A1:C3');
   * sheet.pasteRange('E1'); // Paste values and styles
   * sheet.pasteRange('H1', { valuesOnly: true }); // Paste values only
   * ```
   */
  pasteRange(
    target: string | { row: number; col: number },
    options: PasteOptions = {}
  ): this {
    if (!this._clipboard) {
      return this;
    }

    const {
      valuesOnly = false,
      stylesOnly = false,
      transpose = false,
    } = options;

    let targetRow: number;
    let targetCol: number;

    if (typeof target === 'string') {
      const addr = a1ToAddress(target);
      targetRow = addr.row;
      targetCol = addr.col;
    } else {
      targetRow = target.row;
      targetCol = target.col;
    }

    for (const [relKey, cellData] of this._clipboard.cells) {
      const [relRowStr, relColStr] = relKey.split(',');
      let relRow = parseInt(relRowStr, 10);
      let relCol = parseInt(relColStr, 10);

      // Handle transpose
      if (transpose) {
        [relRow, relCol] = [relCol, relRow];
      }

      const destRow = targetRow + relRow;
      const destCol = targetCol + relCol;
      const destCell = this.cell(destRow, destCol);

      if (!stylesOnly) {
        if (cellData.formula) {
          // Adjust formula references (simplified - just copy as-is for now)
          destCell.setFormula('=' + cellData.formula);
        } else {
          destCell.value = cellData.value;
        }
      }

      if (!valuesOnly && cellData.style) {
        destCell.style = { ...cellData.style };
      }
    }

    return this;
  }

  /**
   * Check if there's content in the clipboard
   */
  get hasClipboard(): boolean {
    return this._clipboard !== null && this._clipboard.cells.size > 0;
  }

  /**
   * Clear the internal clipboard
   */
  clearClipboard(): this {
    this._clipboard = null;
    return this;
  }

  /**
   * Duplicate a range to another location (copy + paste in one operation)
   *
   * @param source - Source range
   * @param target - Target location (top-left corner)
   * @param options - Paste options
   */
  duplicateRange(
    source: string | RangeDefinition,
    target: string | { row: number; col: number },
    options: PasteOptions = {}
  ): this {
    this.copyRange(source);
    this.pasteRange(target, options);
    return this;
  }

  // ============ Row/Column Insert/Delete ============

  /**
   * Insert one or more rows at the specified index
   *
   * @param rowIndex - Index where to insert (0-based)
   * @param count - Number of rows to insert (default: 1)
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * // Insert a single row at index 2 (before row 3)
   * sheet.insertRow(2);
   *
   * // Insert 3 rows at index 0 (at the top)
   * sheet.insertRow(0, 3);
   * ```
   */
  insertRow(rowIndex: number, count: number = 1): this {
    if (count <= 0) return this;

    // Disable events during restructuring
    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;

    try {
      // Collect cells that need to be shifted
      const cellsToShift: { key: string; cell: Cell; newRow: number }[] = [];

      for (const [key, cell] of this._cells) {
        if (cell.row >= rowIndex) {
          cellsToShift.push({ key, cell: cell.clone(), newRow: cell.row + count });
        }
      }

      // First, delete all old cells
      for (const { key } of cellsToShift) {
        this._cells.delete(key);
      }

      // Then add cells at new positions
      for (const { cell, newRow } of cellsToShift) {
        const newCell = this.createCellFromSource(newRow, cell.col, cell);
        this._cells.set(cellKey(newRow, cell.col), newCell);
      }

      // Shift row configurations
      const rowConfigs = new Map<number, RowConfig>();
      for (const [idx, config] of this._rows) {
        if (idx >= rowIndex) {
          rowConfigs.set(idx + count, config);
        } else {
          rowConfigs.set(idx, config);
        }
      }
      this._rows = rowConfigs;

      this.recalculateDimensions();
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  /**
   * Insert one or more columns at the specified index
   *
   * @param colIndex - Column index where to insert (0-based)
   * @param count - Number of columns to insert (default: 1)
   * @returns this for chaining
   */
  insertColumn(colIndex: number, count: number = 1): this {
    if (count <= 0) return this;

    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;

    try {
      const cellsToShift: { key: string; cell: Cell; newCol: number }[] = [];

      for (const [key, cell] of this._cells) {
        if (cell.col >= colIndex) {
          cellsToShift.push({ key, cell: cell.clone(), newCol: cell.col + count });
        }
      }

      // First, delete all old cells
      for (const { key } of cellsToShift) {
        this._cells.delete(key);
      }

      // Then add cells at new positions
      for (const { cell, newCol } of cellsToShift) {
        const newCell = this.createCellFromSource(cell.row, newCol, cell);
        this._cells.set(cellKey(cell.row, newCol), newCell);
      }

      // Shift column configurations
      const colConfigs = new Map<number, ColumnConfig>();
      for (const [idx, config] of this._cols) {
        if (idx >= colIndex) {
          colConfigs.set(idx + count, config);
        } else {
          colConfigs.set(idx, config);
        }
      }
      this._cols = colConfigs;

      this.recalculateDimensions();
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  /**
   * Delete one or more rows at the specified index
   *
   * @param rowIndex - Index of first row to delete (0-based)
   * @param count - Number of rows to delete (default: 1)
   * @returns this for chaining
   */
  deleteRow(rowIndex: number, count: number = 1): this {
    if (count <= 0) return this;

    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;

    try {
      const cellsToDelete: string[] = [];
      const cellsToShift: { key: string; cell: Cell; newRow: number }[] = [];

      for (const [key, cell] of this._cells) {
        if (cell.row >= rowIndex && cell.row < rowIndex + count) {
          // Cell is in deleted range
          cellsToDelete.push(key);
        } else if (cell.row >= rowIndex + count) {
          // Cell needs to shift up
          cellsToShift.push({ key, cell: cell.clone(), newRow: cell.row - count });
        }
      }

      // Delete cells in range
      for (const key of cellsToDelete) {
        this._cells.delete(key);
      }

      // Delete old shifted cells first
      for (const { key } of cellsToShift) {
        this._cells.delete(key);
      }

      // Then add cells at new positions
      for (const { cell, newRow } of cellsToShift) {
        const newCell = this.createCellFromSource(newRow, cell.col, cell);
        this._cells.set(cellKey(newRow, cell.col), newCell);
      }

      // Shift row configurations
      const rowConfigs = new Map<number, RowConfig>();
      for (const [idx, config] of this._rows) {
        if (idx >= rowIndex && idx < rowIndex + count) {
          continue;
        } else if (idx >= rowIndex + count) {
          rowConfigs.set(idx - count, config);
        } else {
          rowConfigs.set(idx, config);
        }
      }
      this._rows = rowConfigs;

      this.recalculateDimensions();
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  /**
   * Delete one or more columns at the specified index
   *
   * @param colIndex - Index of first column to delete (0-based)
   * @param count - Number of columns to delete (default: 1)
   * @returns this for chaining
   */
  deleteColumn(colIndex: number, count: number = 1): this {
    if (count <= 0) return this;

    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;

    try {
      const cellsToDelete: string[] = [];
      const cellsToShift: { key: string; cell: Cell; newCol: number }[] = [];

      for (const [key, cell] of this._cells) {
        if (cell.col >= colIndex && cell.col < colIndex + count) {
          cellsToDelete.push(key);
        } else if (cell.col >= colIndex + count) {
          cellsToShift.push({ key, cell: cell.clone(), newCol: cell.col - count });
        }
      }

      // Delete cells in range
      for (const key of cellsToDelete) {
        this._cells.delete(key);
      }

      // Delete old shifted cells first
      for (const { key } of cellsToShift) {
        this._cells.delete(key);
      }

      // Then add cells at new positions
      for (const { cell, newCol } of cellsToShift) {
        const newCell = this.createCellFromSource(cell.row, newCol, cell);
        this._cells.set(cellKey(cell.row, newCol), newCell);
      }

      // Shift column configurations
      const colConfigs = new Map<number, ColumnConfig>();
      for (const [idx, config] of this._cols) {
        if (idx >= colIndex && idx < colIndex + count) {
          continue;
        } else if (idx >= colIndex + count) {
          colConfigs.set(idx - count, config);
        } else {
          colConfigs.set(idx, config);
        }
      }
      this._cols = colConfigs;

      this.recalculateDimensions();
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  /**
   * Move a row from one position to another
   *
   * @param fromIndex - Source row index
   * @param toIndex - Target row index
   */
  moveRow(fromIndex: number, toIndex: number): this {
    if (fromIndex === toIndex) return this;

    // Copy the row data
    const rowCells: Cell[] = [];
    for (const cell of this._cells.values()) {
      if (cell.row === fromIndex) {
        rowCells.push(cell.clone());
      }
    }

    // Delete the source row
    this.deleteRow(fromIndex);

    // Adjust target if needed
    const adjustedTo = fromIndex < toIndex ? toIndex - 1 : toIndex;

    // Insert at target
    this.insertRow(adjustedTo);

    // Place the cells
    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;
    try {
      for (const cell of rowCells) {
        const newCell = this.cell(adjustedTo, cell.col);
        newCell.value = cell.value;
        if (cell.style) newCell.style = cell.style;
        if (cell.formula) newCell.setFormula('=' + cell.formula.formula, cell.formula.result);
      }
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  /**
   * Move a column from one position to another
   *
   * @param fromIndex - Source column index
   * @param toIndex - Target column index
   */
  moveColumn(fromIndex: number, toIndex: number): this {
    if (fromIndex === toIndex) return this;

    const colCells: Cell[] = [];
    for (const cell of this._cells.values()) {
      if (cell.col === fromIndex) {
        colCells.push(cell.clone());
      }
    }

    this.deleteColumn(fromIndex);

    const adjustedTo = fromIndex < toIndex ? toIndex - 1 : toIndex;

    this.insertColumn(adjustedTo);

    const wasEventsEnabled = this._eventsEnabled;
    this._eventsEnabled = false;
    try {
      for (const cell of colCells) {
        const newCell = this.cell(cell.row, adjustedTo);
        newCell.value = cell.value;
        if (cell.style) newCell.style = cell.style;
        if (cell.formula) newCell.setFormula('=' + cell.formula.formula, cell.formula.result);
      }
    } finally {
      this._eventsEnabled = wasEventsEnabled;
    }

    return this;
  }

  // ============ Data Import/Export Helpers ============

  /**
   * Populate sheet from a 2D array
   *
   * @param data - 2D array of values
   * @param options - Import options
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * sheet.fromArray([
   *   ['Name', 'Age', 'City'],
   *   ['Alice', 25, 'NYC'],
   *   ['Bob', 30, 'LA'],
   * ]);
   * ```
   */
  fromArray(
    data: CellValue[][],
    options: {
      startRow?: number;
      startCol?: number;
      headers?: boolean;
      headerStyle?: CellStyle;
    } = {}
  ): this {
    const {
      startRow = 0,
      startCol = 0,
      headers = false,
      headerStyle,
    } = options;

    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      for (let c = 0; c < row.length; c++) {
        const cell = this.cell(startRow + r, startCol + c);
        cell.value = row[c];

        // Apply header style to first row if specified
        if (headers && r === 0 && headerStyle) {
          cell.style = headerStyle;
        }
      }
    }

    return this;
  }

  /**
   * Populate sheet from an array of objects
   *
   * @param data - Array of objects
   * @param options - Import options
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * sheet.fromObjects([
   *   { name: 'Alice', age: 25, city: 'NYC' },
   *   { name: 'Bob', age: 30, city: 'LA' },
   * ], { includeHeaders: true });
   * ```
   */
  fromObjects<T extends Record<string, CellValue>>(
    data: T[],
    options: {
      startRow?: number;
      startCol?: number;
      includeHeaders?: boolean;
      headerStyle?: CellStyle;
      columns?: (keyof T)[];
    } = {}
  ): this {
    if (data.length === 0) return this;

    const {
      startRow = 0,
      startCol = 0,
      includeHeaders = true,
      headerStyle,
      columns,
    } = options;

    // Determine columns to use
    const keys = columns ?? (Object.keys(data[0]) as (keyof T)[]);

    let currentRow = startRow;

    // Add headers
    if (includeHeaders) {
      for (let c = 0; c < keys.length; c++) {
        const cell = this.cell(currentRow, startCol + c);
        cell.value = String(keys[c]);
        if (headerStyle) {
          cell.style = headerStyle;
        }
      }
      currentRow++;
    }

    // Add data rows
    for (const obj of data) {
      for (let c = 0; c < keys.length; c++) {
        const cell = this.cell(currentRow, startCol + c);
        cell.value = obj[keys[c]] ?? null;
      }
      currentRow++;
    }

    return this;
  }

  /**
   * Export sheet data as a 2D array
   *
   * @param options - Export options
   * @returns 2D array of cell values
   *
   * @example
   * ```typescript
   * const data = sheet.toArray();
   * // [['Name', 'Age'], ['Alice', 25], ['Bob', 30]]
   * ```
   */
  toArray(options: {
    range?: string | RangeDefinition;
    includeEmpty?: boolean;
  } = {}): CellValue[][] {
    const { range, includeEmpty = true } = options;

    let rangeToExport: RangeDefinition;
    if (range) {
      rangeToExport = typeof range === 'string' ? parseRangeReference(range) : range;
    } else {
      const dims = this.dimensions;
      if (!dims) return [];
      rangeToExport = dims;
    }

    const result: CellValue[][] = [];

    for (let r = rangeToExport.startRow; r <= rangeToExport.endRow; r++) {
      const row: CellValue[] = [];
      for (let c = rangeToExport.startCol; c <= rangeToExport.endCol; c++) {
        const cell = this.getCell(r, c);
        row.push(cell?.value ?? null);
      }

      // Skip empty rows if includeEmpty is false
      if (!includeEmpty && row.every(v => v === null)) {
        continue;
      }

      result.push(row);
    }

    return result;
  }

  /**
   * Export sheet data as an array of objects
   *
   * @param options - Export options
   * @returns Array of objects with column headers as keys
   *
   * @example
   * ```typescript
   * const data = sheet.toObjects();
   * // [{ Name: 'Alice', Age: 25 }, { Name: 'Bob', Age: 30 }]
   * ```
   */
  toObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(options: {
    range?: string | RangeDefinition;
    headerRow?: number;
  } = {}): T[] {
    const { range, headerRow = 0 } = options;

    let rangeToExport: RangeDefinition;
    if (range) {
      rangeToExport = typeof range === 'string' ? parseRangeReference(range) : range;
    } else {
      const dims = this.dimensions;
      if (!dims) return [];
      rangeToExport = dims;
    }

    // Get headers from the first row
    const headers: string[] = [];
    for (let c = rangeToExport.startCol; c <= rangeToExport.endCol; c++) {
      const cell = this.getCell(headerRow, c);
      headers.push(cell?.value !== null ? String(cell?.value) : `Column${c}`);
    }

    // Build objects from remaining rows
    const result: T[] = [];
    for (let r = headerRow + 1; r <= rangeToExport.endRow; r++) {
      const obj: Record<string, CellValue> = {};
      for (let c = rangeToExport.startCol; c <= rangeToExport.endCol; c++) {
        const cell = this.getCell(r, c);
        obj[headers[c - rangeToExport.startCol]] = cell?.value ?? null;
      }
      result.push(obj as T);
    }

    return result;
  }

  /**
   * Append a row of data to the end of the sheet
   *
   * @param values - Array of values for the new row
   * @param startCol - Starting column (default: 0)
   * @returns The row index of the appended row
   */
  appendRow(values: CellValue[], startCol: number = 0): number {
    const dims = this.dimensions;
    const newRow = dims ? dims.endRow + 1 : 0;

    for (let c = 0; c < values.length; c++) {
      this.cell(newRow, startCol + c).value = values[c];
    }

    return newRow;
  }

  /**
   * Append multiple rows of data to the end of the sheet
   *
   * @param rows - 2D array of values
   * @param startCol - Starting column (default: 0)
   * @returns The starting row index of the appended rows
   */
  appendRows(rows: CellValue[][], startCol: number = 0): number {
    const dims = this.dimensions;
    const startRow = dims ? dims.endRow + 1 : 0;

    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      for (let c = 0; c < row.length; c++) {
        this.cell(startRow + r, startCol + c).value = row[c];
      }
    }

    return startRow;
  }

  // ============ Internal Helpers ============

  /**
   * Internal: Create a new cell at (row, col) by copying all properties from a source cell
   */
  private createCellFromSource(row: number, col: number, source: Cell): Cell {
    const newCell = new Cell(row, col, source.value);
    if (source.style) newCell.style = source.style;
    if (source.formula) newCell.setFormula('=' + source.formula.formula, source.formula.result);
    if (source.hyperlink) newCell.setHyperlink(source.hyperlink.target, source.hyperlink.tooltip);
    if (source.comment) newCell.setComment(source.comment.text as string, source.comment.author);
    newCell._onChange = this.handleCellChange.bind(this);
    return newCell;
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

  // ============ Event System ============

  /**
   * Subscribe to sheet events
   *
   * @param eventType - The event type to listen for ('cellChange', 'cellStyleChange', 'cellAdded', 'cellDeleted', '*')
   * @param handler - The callback function to invoke when the event occurs
   * @returns this for chaining
   *
   * @example
   * ```typescript
   * sheet.on('cellChange', (event) => {
   *   console.log(`Cell ${event.address} changed from ${event.oldValue} to ${event.newValue}`);
   * });
   *
   * // Listen to all events
   * sheet.on('*', (event) => {
   *   console.log(`Event: ${event.type}`);
   * });
   * ```
   */
  on<K extends keyof SheetEventMap>(eventType: K, handler: SheetEventHandler<SheetEventMap[K]>): this {
    if (!this._eventListeners.has(eventType)) {
      this._eventListeners.set(eventType, new Set());
    }
    this._eventListeners.get(eventType)!.add(handler as SheetEventHandler);
    return this;
  }

  /**
   * Unsubscribe from sheet events
   *
   * @param eventType - The event type to stop listening for
   * @param handler - The callback function to remove
   * @returns this for chaining
   */
  off<K extends keyof SheetEventMap>(eventType: K, handler: SheetEventHandler<SheetEventMap[K]>): this {
    const listeners = this._eventListeners.get(eventType);
    if (listeners) {
      listeners.delete(handler as SheetEventHandler);
    }
    return this;
  }

  /**
   * Enable or disable event emission
   *
   * @param enabled - Whether events should be emitted
   *
   * @example
   * ```typescript
   * // Disable events during bulk operations
   * sheet.setEventsEnabled(false);
   * for (let i = 0; i < 1000; i++) {
   *   sheet.cell(i, 0).value = i;
   * }
   * sheet.setEventsEnabled(true);
   * ```
   */
  setEventsEnabled(enabled: boolean): this {
    this._eventsEnabled = enabled;
    return this;
  }

  /**
   * Check if events are currently enabled
   */
  get eventsEnabled(): boolean {
    return this._eventsEnabled;
  }

  /**
   * Get all tracked changes since last commit
   *
   * @returns Array of change records
   *
   * @example
   * ```typescript
   * sheet.cell('A1').value = 'Hello';
   * sheet.cell('B1').value = 'World';
   *
   * const changes = sheet.getChanges();
   * console.log(changes.length); // 2
   *
   * // Sync changes to server
   * await syncChanges(changes);
   *
   * // Clear change buffer
   * sheet.commitChanges();
   * ```
   */
  getChanges(): readonly ChangeRecord[] {
    return this._changes;
  }

  /**
   * Clear the change buffer
   *
   * Call this after successfully syncing changes to indicate they've been persisted.
   */
  commitChanges(): this {
    this._changes = [];
    return this;
  }

  /**
   * Get the number of pending changes
   */
  get changeCount(): number {
    return this._changes.length;
  }

  // ============ Undo/Redo ============

  /**
   * Check if there are changes that can be undone
   */
  get canUndo(): boolean {
    return this._undoStack.length > 0;
  }

  /**
   * Check if there are changes that can be redone
   */
  get canRedo(): boolean {
    return this._redoStack.length > 0;
  }

  /**
   * Get the number of undo steps available
   */
  get undoCount(): number {
    return this._undoStack.length;
  }

  /**
   * Get the number of redo steps available
   */
  get redoCount(): number {
    return this._redoStack.length;
  }

  /**
   * Undo the last change
   *
   * @returns true if undo was successful, false if nothing to undo
   *
   * @example
   * ```typescript
   * sheet.cell('A1').value = 'Hello';
   * sheet.cell('A1').value = 'World';
   *
   * sheet.undo(); // A1 is now 'Hello'
   * sheet.undo(); // A1 is now null
   * ```
   */
  undo(): boolean {
    if (!this.canUndo) {
      return false;
    }

    const change = this._undoStack.pop()!;
    this._isUndoRedoOperation = true;

    try {
      // Check if this is a batch change
      const batchChanges = (change as ChangeRecord & { _batchChanges?: ChangeRecord[] })._batchChanges;

      if (batchChanges) {
        // Undo batch changes in reverse order
        for (let i = batchChanges.length - 1; i >= 0; i--) {
          const batchChange = batchChanges[i];
          this.applyUndoChange(batchChange);
        }
      } else {
        this.applyUndoChange(change);
      }

      // Push to redo stack
      this._redoStack.push(change);

      return true;
    } finally {
      this._isUndoRedoOperation = false;
    }
  }

  /**
   * Internal: Apply a single undo change
   */
  private applyUndoChange(change: ChangeRecord): void {
    if (change.type === 'value' || change.type === 'formula') {
      const cell = this.cell(change.row, change.col);
      cell.value = change.oldValue ?? null;
    } else if (change.type === 'style') {
      const cell = this.cell(change.row, change.col);
      cell.style = change.oldStyle;
    }
  }

  /**
   * Redo the last undone change
   *
   * @returns true if redo was successful, false if nothing to redo
   *
   * @example
   * ```typescript
   * sheet.cell('A1').value = 'Hello';
   * sheet.undo(); // A1 is now null
   * sheet.redo(); // A1 is now 'Hello'
   * ```
   */
  redo(): boolean {
    if (!this.canRedo) {
      return false;
    }

    const change = this._redoStack.pop()!;
    this._isUndoRedoOperation = true;

    try {
      // Check if this is a batch change
      const batchChanges = (change as ChangeRecord & { _batchChanges?: ChangeRecord[] })._batchChanges;

      if (batchChanges) {
        // Redo batch changes in order
        for (const batchChange of batchChanges) {
          this.applyRedoChange(batchChange);
        }
      } else {
        this.applyRedoChange(change);
      }

      // Push back to undo stack
      this._undoStack.push(change);

      return true;
    } finally {
      this._isUndoRedoOperation = false;
    }
  }

  /**
   * Internal: Apply a single redo change
   */
  private applyRedoChange(change: ChangeRecord): void {
    if (change.type === 'value' || change.type === 'formula') {
      const cell = this.cell(change.row, change.col);
      cell.value = change.newValue ?? null;
    } else if (change.type === 'style') {
      const cell = this.cell(change.row, change.col);
      cell.style = change.newStyle;
    }
  }

  /**
   * Clear the undo and redo history
   *
   * Useful after saving or when you want to prevent undoing past a certain point.
   */
  clearHistory(): this {
    this._undoStack = [];
    this._redoStack = [];
    return this;
  }

  /**
   * Set the maximum number of undo steps to keep
   *
   * @param max - Maximum history size (default: 100)
   */
  setMaxUndoHistory(max: number): this {
    this._maxUndoHistory = max;
    // Trim if necessary
    while (this._undoStack.length > max) {
      this._undoStack.shift();
    }
    return this;
  }

  /**
   * Execute a batch of operations as a single undo step
   *
   * @param fn - Function containing the batch operations
   *
   * @example
   * ```typescript
   * sheet.batch(() => {
   *   sheet.cell('A1').value = 'Hello';
   *   sheet.cell('B1').value = 'World';
   *   sheet.cell('C1').value = '!';
   * });
   *
   * // Single undo reverts all three changes
   * sheet.undo();
   * ```
   */
  batch(fn: () => void): this {
    const startIndex = this._undoStack.length;

    fn();

    // Collect all changes made during the batch
    const batchChanges = this._undoStack.splice(startIndex);

    if (batchChanges.length > 0) {
      // Create a composite change record
      const compositeChange: ChangeRecord = {
        id: `${this._name}-batch-${++this._changeIdCounter}`,
        type: 'value',
        address: batchChanges[0].address,
        row: batchChanges[0].row,
        col: batchChanges[0].col,
        oldValue: batchChanges[0].oldValue,
        newValue: batchChanges[batchChanges.length - 1].newValue,
        timestamp: Date.now(),
      };

      // Store batch changes for proper undo
      (compositeChange as ChangeRecord & { _batchChanges?: ChangeRecord[] })._batchChanges = batchChanges;

      this._undoStack.push(compositeChange);
    }

    return this;
  }

  /**
   * Internal: Handle cell change notifications
   */
  private handleCellChange(cell: Cell, changeType: 'value' | 'style' | 'formula', oldValue?: CellValue | CellStyle): void {
    if (!this._eventsEnabled) return;

    const timestamp = Date.now();

    if (changeType === 'value' || changeType === 'formula') {
      const changeRecord: ChangeRecord = {
        id: `${this._name}-${++this._changeIdCounter}`,
        type: changeType,
        address: cell.address,
        row: cell.row,
        col: cell.col,
        oldValue: oldValue as CellValue,
        newValue: cell.value,
        timestamp,
      };

      // Record the change
      this._changes.push(changeRecord);

      // Add to undo stack (unless this is an undo/redo operation)
      if (!this._isUndoRedoOperation) {
        this._undoStack.push(changeRecord);
        // Limit undo history
        if (this._undoStack.length > this._maxUndoHistory) {
          this._undoStack.shift();
        }
        // Clear redo stack on new change
        this._redoStack = [];
      }

      // Emit event
      const event: CellChangeEvent = {
        type: 'cellChange',
        sheetName: this._name,
        address: cell.address,
        row: cell.row,
        col: cell.col,
        oldValue: oldValue as CellValue,
        newValue: cell.value,
        timestamp,
      };
      this.emit('cellChange', event);
    } else if (changeType === 'style') {
      const changeRecord: ChangeRecord = {
        id: `${this._name}-${++this._changeIdCounter}`,
        type: 'style',
        address: cell.address,
        row: cell.row,
        col: cell.col,
        oldStyle: oldValue as CellStyle,
        newStyle: cell.style,
        timestamp,
      };

      // Record the change
      this._changes.push(changeRecord);

      // Add to undo stack (unless this is an undo/redo operation)
      if (!this._isUndoRedoOperation) {
        this._undoStack.push(changeRecord);
        if (this._undoStack.length > this._maxUndoHistory) {
          this._undoStack.shift();
        }
        this._redoStack = [];
      }

      // Emit event
      const event: CellStyleChangeEvent = {
        type: 'cellStyleChange',
        sheetName: this._name,
        address: cell.address,
        row: cell.row,
        col: cell.col,
        oldStyle: oldValue as CellStyle,
        newStyle: cell.style,
        timestamp,
      };
      this.emit('cellStyleChange', event);
    }
  }

  /**
   * Internal: Emit a cell added event
   */
  private emitCellAdded(cell: Cell): void {
    if (!this._eventsEnabled) return;

    const event: CellAddedEvent = {
      type: 'cellAdded',
      sheetName: this._name,
      address: cell.address,
      row: cell.row,
      col: cell.col,
      timestamp: Date.now(),
    };
    this.emit('cellAdded', event);
  }

  /**
   * Internal: Emit a cell deleted event
   */
  private emitCellDeleted(cell: Cell): void {
    if (!this._eventsEnabled) return;

    const timestamp = Date.now();

    const event: CellDeletedEvent = {
      type: 'cellDeleted',
      sheetName: this._name,
      address: cell.address,
      row: cell.row,
      col: cell.col,
      value: cell.value,
      timestamp,
    };

    // Record the change
    this._changes.push({
      id: `${this._name}-${++this._changeIdCounter}`,
      type: 'delete',
      address: cell.address,
      row: cell.row,
      col: cell.col,
      oldValue: cell.value,
      timestamp,
    });

    this.emit('cellDeleted', event);
  }

  /**
   * Internal: Emit an event to all listeners
   */
  private emit<K extends keyof SheetEventMap>(eventType: K, event: SheetEventMap[K]): void {
    // Emit to specific listeners
    const listeners = this._eventListeners.get(eventType);
    if (listeners) {
      for (const handler of listeners) {
        handler(event);
      }
    }

    // Emit to wildcard listeners
    const wildcardListeners = this._eventListeners.get('*');
    if (wildcardListeners) {
      for (const handler of wildcardListeners) {
        handler(event);
      }
    }
  }
}

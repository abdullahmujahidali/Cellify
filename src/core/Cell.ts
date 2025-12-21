import type {
  CellValue,
  CellValueType,
  CellFormula,
  CellHyperlink,
  CellComment,
  CellValidation,
  MergeRange,
  CellAddress,
  RichTextValue,
} from '../types/cell.types.js';
import type { CellStyle } from '../types/style.types.js';
import { getCellValueType, addressToA1 } from '../types/cell.types.js';

/**
 * Callback type for cell change notifications
 */
export type CellChangeCallback = (
  cell: Cell,
  changeType: 'value' | 'style' | 'formula',
  oldValue?: CellValue | CellStyle
) => void;

/**
 * Represents a single cell in a spreadsheet.
 *
 * Cells are the fundamental unit of data in Cellify. Each cell can hold:
 * - A value (string, number, boolean, date, error, or rich text)
 * - A formula that computes the value
 * - Styling (font, fill, borders, alignment, number format)
 * - A hyperlink
 * - A comment/note
 * - Data validation rules
 * - Merge information
 */
export class Cell {
  private _value: CellValue = null;
  private _formula: CellFormula | undefined;
  private _style: CellStyle | undefined;
  private _hyperlink: CellHyperlink | undefined;
  private _comment: CellComment | undefined;
  private _validation: CellValidation | undefined;
  private _merge: MergeRange | undefined;
  private _mergedInto: CellAddress | undefined;

  /**
   * Optional callback for change notifications (set by Sheet)
   * @internal
   */
  _onChange: CellChangeCallback | undefined;

  /**
   * The row index of this cell (0-based)
   */
  readonly row: number;

  /**
   * The column index of this cell (0-based)
   */
  readonly col: number;

  constructor(row: number, col: number, value?: CellValue) {
    this.row = row;
    this.col = col;
    if (value !== undefined) {
      this._value = value;
    }
  }

  /**
   * Get the cell's address in A1 notation (e.g., "A1", "B2")
   */
  get address(): string {
    return addressToA1(this.row, this.col);
  }

  /**
   * Get the cell's value (returns formula result if formula exists)
   */
  get value(): CellValue {
    // If formula has a cached result, return it
    if (this._formula?.result !== undefined) {
      return this._formula.result;
    }
    return this._value;
  }

  /**
   * Set the cell's value
   */
  set value(val: CellValue) {
    const oldValue = this._value;
    this._value = val;
    // Clear formula when value is set directly
    if (this._formula) {
      this._formula = undefined;
    }
    // Notify of change
    if (this._onChange && oldValue !== val) {
      this._onChange(this, 'value', oldValue);
    }
  }

  /**
   * Get the type of the cell's value
   */
  get type(): CellValueType {
    if (this._formula) {
      return 'formula';
    }
    return getCellValueType(this._value);
  }

  /**
   * Get the cell's formula
   */
  get formula(): CellFormula | undefined {
    return this._formula;
  }

  /**
   * Set a formula on the cell
   * @param formulaText - The formula text (with or without leading '=')
   * @param result - Optional cached result value from Excel
   */
  setFormula(formulaText: string, result?: CellValue): this {
    const oldValue = this._value;
    const text = formulaText.startsWith('=') ? formulaText.slice(1) : formulaText;
    this._formula = {
      formula: text,
      result: result,
    };
    if (this._onChange) {
      this._onChange(this, 'formula', oldValue);
    }
    return this;
  }

  /**
   * Clear the formula from the cell
   */
  clearFormula(): this {
    this._formula = undefined;
    return this;
  }

  /**
   * Get the cell's style
   */
  get style(): CellStyle | undefined {
    return this._style;
  }

  /**
   * Set the cell's style (replaces existing style)
   */
  set style(style: CellStyle | undefined) {
    const oldStyle = this._style;
    this._style = style;
    // Notify of change
    if (this._onChange && oldStyle !== style) {
      this._onChange(this, 'style', oldStyle);
    }
  }

  /**
   * Apply partial style updates (merges with existing style)
   */
  applyStyle(style: Partial<CellStyle>): this {
    const oldStyle = this._style;
    if (!this._style) {
      this._style = { ...style };
    } else {
      this._style = this.mergeStyles(this._style, style);
    }
    if (this._onChange) {
      this._onChange(this, 'style', oldStyle);
    }
    return this;
  }

  /**
   * Deep merge two style objects
   */
  private mergeStyles(base: CellStyle, override: Partial<CellStyle>): CellStyle {
    const result: CellStyle = { ...base };

    if (override.font) {
      result.font = { ...base.font, ...override.font };
    }
    if (override.fill) {
      result.fill = { ...base.fill, ...override.fill };
    }
    if (override.borders) {
      result.borders = { ...base.borders, ...override.borders };
    }
    if (override.alignment) {
      result.alignment = { ...base.alignment, ...override.alignment };
    }
    if (override.numberFormat) {
      result.numberFormat = { ...base.numberFormat, ...override.numberFormat };
    }
    if (override.protection) {
      result.protection = { ...base.protection, ...override.protection };
    }

    return result;
  }

  /**
   * Get the cell's hyperlink
   */
  get hyperlink(): CellHyperlink | undefined {
    return this._hyperlink;
  }

  /**
   * Set a hyperlink on the cell
   */
  setHyperlink(target: string, tooltip?: string): this {
    this._hyperlink = { target, tooltip };
    return this;
  }

  /**
   * Remove the hyperlink from the cell
   */
  clearHyperlink(): this {
    this._hyperlink = undefined;
    return this;
  }

  /**
   * Get the cell's comment
   */
  get comment(): CellComment | undefined {
    return this._comment;
  }

  /**
   * Set the cell's comment
   */
  set comment(value: CellComment | string | undefined | null) {
    if (value === undefined || value === null) {
      this._comment = undefined;
    } else if (typeof value === 'string') {
      this._comment = { text: value };
    } else {
      this._comment = value;
    }
  }

  /**
   * Set a comment on the cell
   */
  setComment(text: string | RichTextValue, author?: string): this {
    this._comment = { text, author };
    return this;
  }

  /**
   * Remove the comment from the cell
   */
  clearComment(): this {
    this._comment = undefined;
    return this;
  }

  /**
   * Get the cell's validation rules
   */
  get validation(): CellValidation | undefined {
    return this._validation;
  }

  /**
   * Set data validation on the cell
   */
  setValidation(validation: CellValidation): this {
    this._validation = validation;
    return this;
  }

  /**
   * Remove validation from the cell
   */
  clearValidation(): this {
    this._validation = undefined;
    return this;
  }

  /**
   * Get merge information (only set on master cell of a merge)
   */
  get merge(): MergeRange | undefined {
    return this._merge;
  }

  /**
   * Check if this cell is the master (top-left) cell of a merge
   */
  get isMergeMaster(): boolean {
    return this._merge !== undefined;
  }

  /**
   * Get the address of the master cell if this cell is part of a merge
   */
  get mergedInto(): CellAddress | undefined {
    return this._mergedInto;
  }

  /**
   * Check if this cell is part of a merge (but not the master)
   */
  get isMergedSlave(): boolean {
    return this._mergedInto !== undefined;
  }

  /**
   * Check if this cell is part of any merge
   */
  get isMerged(): boolean {
    return this.isMergeMaster || this.isMergedSlave;
  }

  /**
   * Internal: Set merge info (called by Sheet)
   */
  _setMerge(merge: MergeRange | undefined): void {
    this._merge = merge;
  }

  /**
   * Internal: Set merged-into info (called by Sheet)
   */
  _setMergedInto(address: CellAddress | undefined): void {
    this._mergedInto = address;
  }

  /**
   * Check if the cell has any content or styling
   */
  get isEmpty(): boolean {
    return (
      this._value === null &&
      this._formula === undefined &&
      this._style === undefined &&
      this._hyperlink === undefined &&
      this._comment === undefined &&
      this._validation === undefined
    );
  }

  /**
   * Clear all content and styling from the cell
   */
  clear(): this {
    this._value = null;
    this._formula = undefined;
    this._style = undefined;
    this._hyperlink = undefined;
    this._comment = undefined;
    this._validation = undefined;
    // Note: merge info is managed by Sheet, not cleared here
    return this;
  }

  /**
   * Create a deep clone of this cell
   */
  clone(): Cell {
    const cell = new Cell(this.row, this.col, this._value);
    if (this._formula) {
      cell._formula = { ...this._formula };
    }
    if (this._style) {
      cell._style = JSON.parse(JSON.stringify(this._style));
    }
    if (this._hyperlink) {
      cell._hyperlink = { ...this._hyperlink };
    }
    if (this._comment) {
      cell._comment = { ...this._comment };
    }
    if (this._validation) {
      cell._validation = { ...this._validation };
    }
    return cell;
  }

  /**
   * Convert cell to a plain object for serialization
   */
  toJSON(): Record<string, unknown> {
    const obj: Record<string, unknown> = {
      row: this.row,
      col: this.col,
      address: this.address,
    };

    if (this._value !== null) {
      obj.value = this._value;
      obj.type = this.type;
    }
    if (this._formula) {
      obj.formula = this._formula;
    }
    if (this._style) {
      obj.style = this._style;
    }
    if (this._hyperlink) {
      obj.hyperlink = this._hyperlink;
    }
    if (this._comment) {
      obj.comment = this._comment;
    }
    if (this._validation) {
      obj.validation = this._validation;
    }
    if (this._merge) {
      obj.merge = this._merge;
    }
    if (this._mergedInto) {
      obj.mergedInto = this._mergedInto;
    }

    return obj;
  }
}

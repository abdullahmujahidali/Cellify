import { Sheet } from './Sheet.js';
import type { CellStyle, NamedStyle } from '../types/style.types.js';

/**
 * Workbook properties/metadata
 */
export interface WorkbookProperties {
  title?: string;
  subject?: string;
  author?: string;
  company?: string;
  category?: string;
  keywords?: string[];
  comments?: string;
  manager?: string;
  created?: Date;
  modified?: Date;
  lastModifiedBy?: string;
  revision?: number;
}

/**
 * Defined name (named range or formula)
 */
export interface DefinedName {
  name: string;
  formula: string; // Can be a range reference like "Sheet1!$A$1:$B$10" or a formula
  scope?: string; // Sheet name for local scope, undefined for global
  comment?: string;
  hidden?: boolean;
}

/**
 * Workbook calculation mode
 */
export type CalculationMode = 'auto' | 'manual' | 'autoNoTable';

/**
 * Workbook-level view settings
 */
export interface WorkbookView {
  activeSheet?: number; // Index of active sheet
  firstSheet?: number; // First visible sheet tab
  showSheetTabs?: boolean;
  tabRatio?: number; // Ratio of sheet tab bar to horizontal scroll bar
}

/**
 * Represents an Excel workbook.
 *
 * The Workbook is the top-level container that holds:
 * - One or more worksheets
 * - Workbook properties (title, author, etc.)
 * - Named styles for reuse across cells
 * - Defined names (named ranges)
 * - Calculation settings
 */
export class Workbook {
  private _sheets: Sheet[] = [];
  private _properties: WorkbookProperties = {};
  private _namedStyles: Map<string, NamedStyle> = new Map();
  private _definedNames: Map<string, DefinedName> = new Map();
  private _calculationMode: CalculationMode = 'auto';
  private _view: WorkbookView = {};

  constructor() {
    // Initialize with creation date
    this._properties.created = new Date();
  }

  // ============ Sheet Management ============

  /**
   * Add a new sheet to the workbook
   */
  addSheet(name?: string): Sheet {
    const sheetName = name ?? this.generateSheetName();
    this.validateSheetName(sheetName);

    const sheet = new Sheet(sheetName);
    this._sheets.push(sheet);
    return sheet;
  }

  /**
   * Get a sheet by name
   */
  getSheet(name: string): Sheet | undefined {
    return this._sheets.find((s) => s.name === name);
  }

  /**
   * Get a sheet by index (0-based)
   */
  getSheetByIndex(index: number): Sheet | undefined {
    return this._sheets[index];
  }

  /**
   * Get the index of a sheet
   */
  getSheetIndex(sheet: Sheet | string): number {
    if (typeof sheet === 'string') {
      return this._sheets.findIndex((s) => s.name === sheet);
    }
    return this._sheets.indexOf(sheet);
  }

  /**
   * Remove a sheet by name or reference
   */
  removeSheet(sheet: Sheet | string): boolean {
    const index = this.getSheetIndex(sheet);
    if (index === -1) return false;

    this._sheets.splice(index, 1);
    return true;
  }

  /**
   * Rename a sheet
   */
  renameSheet(oldName: string, newName: string): boolean {
    const sheet = this.getSheet(oldName);
    if (!sheet) return false;

    this.validateSheetName(newName);
    sheet.name = newName;
    return true;
  }

  /**
   * Move a sheet to a new position
   */
  moveSheet(sheet: Sheet | string, newIndex: number): boolean {
    const currentIndex = this.getSheetIndex(sheet);
    if (currentIndex === -1) return false;

    const [removed] = this._sheets.splice(currentIndex, 1);
    this._sheets.splice(newIndex, 0, removed);
    return true;
  }

  /**
   * Duplicate a sheet
   */
  duplicateSheet(sheet: Sheet | string, newName?: string): Sheet | undefined {
    const source = typeof sheet === 'string' ? this.getSheet(sheet) : sheet;
    if (!source) return undefined;

    const name = newName ?? this.generateCopyName(source.name);
    this.validateSheetName(name);

    // Create new sheet from JSON (deep copy)
    const newSheet = new Sheet(name);
    // TODO: Implement full deep copy when we have fromJSON
    this._sheets.push(newSheet);
    return newSheet;
  }

  /**
   * Get all sheets
   */
  get sheets(): readonly Sheet[] {
    return this._sheets;
  }

  /**
   * Get the number of sheets
   */
  get sheetCount(): number {
    return this._sheets.length;
  }

  /**
   * Generate a unique sheet name
   */
  private generateSheetName(): string {
    let index = 1;
    let name = `Sheet${index}`;
    while (this.getSheet(name)) {
      index++;
      name = `Sheet${index}`;
    }
    return name;
  }

  /**
   * Generate a name for a copied sheet
   */
  private generateCopyName(originalName: string): string {
    let index = 2;
    let name = `${originalName} (${index})`;
    while (this.getSheet(name)) {
      index++;
      name = `${originalName} (${index})`;
    }
    return name;
  }

  /**
   * Validate a sheet name
   */
  private validateSheetName(name: string): void {
    if (!name || name.trim() === '') {
      throw new Error('Sheet name cannot be empty');
    }
    if (name.length > 31) {
      throw new Error('Sheet name cannot exceed 31 characters');
    }
    if (/[\\\/\*\?\[\]:']/.test(name)) {
      throw new Error('Sheet name contains invalid characters: \\ / * ? [ ] : \'');
    }
    if (this.getSheet(name)) {
      throw new Error(`Sheet with name "${name}" already exists`);
    }
  }

  // ============ Workbook Properties ============

  /**
   * Get workbook properties
   */
  get properties(): WorkbookProperties {
    return this._properties;
  }

  /**
   * Set workbook properties
   */
  setProperties(props: Partial<WorkbookProperties>): this {
    this._properties = { ...this._properties, ...props };
    return this;
  }

  /**
   * Set the title
   */
  set title(value: string) {
    this._properties.title = value;
  }

  get title(): string | undefined {
    return this._properties.title;
  }

  /**
   * Set the author
   */
  set author(value: string) {
    this._properties.author = value;
  }

  get author(): string | undefined {
    return this._properties.author;
  }

  // ============ Named Styles ============

  /**
   * Add a named style for reuse
   */
  addNamedStyle(name: string, style: CellStyle): this {
    this._namedStyles.set(name, { name, style });
    return this;
  }

  /**
   * Get a named style
   */
  getNamedStyle(name: string): NamedStyle | undefined {
    return this._namedStyles.get(name);
  }

  /**
   * Remove a named style
   */
  removeNamedStyle(name: string): boolean {
    return this._namedStyles.delete(name);
  }

  /**
   * Get all named styles
   */
  get namedStyles(): ReadonlyMap<string, NamedStyle> {
    return this._namedStyles;
  }

  // ============ Defined Names ============

  /**
   * Add a defined name (named range or formula)
   */
  addDefinedName(
    name: string,
    formula: string,
    options: { scope?: string; comment?: string; hidden?: boolean } = {}
  ): this {
    this._definedNames.set(name, {
      name,
      formula,
      ...options,
    });
    return this;
  }

  /**
   * Get a defined name
   */
  getDefinedName(name: string): DefinedName | undefined {
    return this._definedNames.get(name);
  }

  /**
   * Remove a defined name
   */
  removeDefinedName(name: string): boolean {
    return this._definedNames.delete(name);
  }

  /**
   * Get all defined names
   */
  get definedNames(): ReadonlyMap<string, DefinedName> {
    return this._definedNames;
  }

  // ============ Calculation Settings ============

  /**
   * Get calculation mode
   */
  get calculationMode(): CalculationMode {
    return this._calculationMode;
  }

  /**
   * Set calculation mode
   */
  set calculationMode(mode: CalculationMode) {
    this._calculationMode = mode;
  }

  // ============ Workbook View ============

  /**
   * Get workbook view settings
   */
  get view(): WorkbookView {
    return this._view;
  }

  /**
   * Set workbook view settings
   */
  setView(view: Partial<WorkbookView>): this {
    this._view = { ...this._view, ...view };
    return this;
  }

  /**
   * Get the active sheet
   */
  get activeSheet(): Sheet | undefined {
    const index = this._view.activeSheet ?? 0;
    return this._sheets[index];
  }

  /**
   * Set the active sheet
   */
  setActiveSheet(sheet: Sheet | string | number): this {
    let index: number;

    if (typeof sheet === 'number') {
      index = sheet;
    } else {
      index = this.getSheetIndex(sheet);
    }

    if (index >= 0 && index < this._sheets.length) {
      this._view.activeSheet = index;
    }

    return this;
  }

  // ============ Utility Methods ============

  /**
   * Update the modified timestamp
   */
  touch(): this {
    this._properties.modified = new Date();
    return this;
  }

  /**
   * Convert workbook to JSON
   */
  toJSON(): Record<string, unknown> {
    return {
      properties: this._properties,
      sheets: this._sheets.map((s) => s.toJSON()),
      namedStyles: Object.fromEntries(this._namedStyles),
      definedNames: Object.fromEntries(this._definedNames),
      calculationMode: this._calculationMode,
      view: this._view,
    };
  }

  /**
   * Create a workbook from JSON
   */
  static fromJSON(_json: Record<string, unknown>): Workbook {
    // TODO: Implement full deserialization
    const workbook = new Workbook();
    return workbook;
  }
}

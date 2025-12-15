/**
 * Accessibility types for Cellify
 *
 * These types help rendering layers create accessible spreadsheet UIs.
 * The headless core provides metadata; renderers use it to generate
 * proper ARIA attributes, roles, and descriptions.
 */

/**
 * ARIA roles for spreadsheet elements
 */
export type SpreadsheetRole =
  | 'grid' // The spreadsheet container
  | 'rowgroup' // Group of rows (header/body)
  | 'row' // A single row
  | 'columnheader' // Column header cell
  | 'rowheader' // Row header cell
  | 'gridcell' // Regular data cell
  | 'presentation'; // Decorative element

/**
 * Cell scope for header cells
 */
export type CellScope = 'col' | 'row' | 'colgroup' | 'rowgroup';

/**
 * Accessibility metadata for a cell
 */
export interface CellAccessibility {
  /** Human-readable description for screen readers */
  description?: string;

  /** Whether this cell is a header */
  isHeader?: boolean;

  /** Scope of header cell */
  scope?: CellScope;

  /** IDs of header cells that label this cell */
  headers?: string[];

  /** ARIA role override */
  role?: SpreadsheetRole;

  /** ARIA label override (if different from cell value) */
  ariaLabel?: string;

  /** Additional description for complex cells */
  ariaDescribedBy?: string;

  /** Whether cell is currently selected */
  ariaSelected?: boolean;

  /** Whether cell is read-only */
  ariaReadOnly?: boolean;

  /** Whether cell is required (for input) */
  ariaRequired?: boolean;

  /** Error message for invalid input */
  ariaInvalid?: boolean | 'grammar' | 'spelling';

  /** Value description (e.g., "25 percent" for 0.25 with % format) */
  ariaValueText?: string;

  /** For cells that act as controls */
  ariaHasPopup?: boolean | 'menu' | 'listbox' | 'tree' | 'grid' | 'dialog';

  /** Whether cell content is expanded (for expandable headers) */
  ariaExpanded?: boolean;

  /** Level in hierarchy (for grouped rows/columns) */
  ariaLevel?: number;

  /** Position in set (for navigation) */
  ariaPosInSet?: number;

  /** Total items in set */
  ariaSetSize?: number;

  /** Column index (1-based for ARIA) */
  ariaColIndex?: number;

  /** Row index (1-based for ARIA) */
  ariaRowIndex?: number;

  /** Column span */
  ariaColSpan?: number;

  /** Row span */
  ariaRowSpan?: number;
}

/**
 * Accessibility metadata for a sheet
 */
export interface SheetAccessibility {
  /** Human-readable name for the sheet */
  label?: string;

  /** Description of sheet purpose */
  description?: string;

  /** Whether the grid is multi-selectable */
  ariaMultiSelectable?: boolean;

  /** Total row count (for virtualized grids) */
  ariaRowCount?: number;

  /** Total column count */
  ariaColCount?: number;

  /** Index of first header row (0-based) */
  headerRowStart?: number;

  /** Index of last header row (0-based) */
  headerRowEnd?: number;

  /** Index of first header column (0-based) */
  headerColStart?: number;

  /** Index of last header column (0-based) */
  headerColEnd?: number;

  /** Caption for the table */
  caption?: string;

  /** Summary of table structure (for complex tables) */
  summary?: string;
}

/**
 * Keyboard navigation configuration
 */
export interface KeyboardNavigation {
  /** Allow arrow key navigation */
  arrowKeys?: boolean;

  /** Allow Tab key to move between cells */
  tabNavigation?: boolean;

  /** Allow Enter to edit cell */
  enterToEdit?: boolean;

  /** Allow Escape to cancel edit */
  escapeToCancel?: boolean;

  /** Allow Home/End for row navigation */
  homeEnd?: boolean;

  /** Allow Ctrl+Home/End for sheet navigation */
  ctrlHomeEnd?: boolean;

  /** Allow Page Up/Down for scrolling */
  pageUpDown?: boolean;

  /** Allow Ctrl+Arrow for jump navigation */
  ctrlArrow?: boolean;

  /** Allow F2 to enter edit mode */
  f2ToEdit?: boolean;

  /** Custom keyboard shortcuts */
  customShortcuts?: KeyboardShortcut[];
}

/**
 * Custom keyboard shortcut definition
 */
export interface KeyboardShortcut {
  /** Key code or key name */
  key: string;

  /** Require Ctrl/Cmd key */
  ctrl?: boolean;

  /** Require Shift key */
  shift?: boolean;

  /** Require Alt/Option key */
  alt?: boolean;

  /** Action to perform */
  action: string;

  /** Description for help text */
  description: string;
}

/**
 * Live region configuration for announcements
 */
export interface LiveRegionConfig {
  /** ARIA live value */
  ariaLive: 'polite' | 'assertive' | 'off';

  /** Whether updates are atomic */
  ariaAtomic?: boolean;

  /** What parts are relevant */
  ariaRelevant?: ('additions' | 'removals' | 'text' | 'all')[];
}

/**
 * Announce types for screen reader notifications
 */
export type AnnounceType =
  | 'cellChange' // Cell value changed
  | 'selection' // Selection changed
  | 'navigation' // Navigation occurred
  | 'error' // Error message
  | 'success' // Success message
  | 'hint'; // Helpful hint

/**
 * Screen reader announcement
 */
export interface Announcement {
  /** The message to announce */
  message: string;

  /** Type of announcement */
  type: AnnounceType;

  /** Priority (assertive = interrupt, polite = queue) */
  priority: 'polite' | 'assertive';

  /** Optional delay before announcement (ms) */
  delay?: number;
}

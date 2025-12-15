// Accessibility types
export type {
  SpreadsheetRole,
  CellScope,
  CellAccessibility,
  SheetAccessibility,
  KeyboardNavigation,
  KeyboardShortcut,
  LiveRegionConfig,
  AnnounceType,
  Announcement,
} from './types.js';

// Accessibility helpers
export {
  getCellAccessibility,
  getValueText,
  getSheetAccessibility,
  describeCellPosition,
  describeCellFull,
  createAnnouncement,
  announceNavigation,
  announceSelection,
  announceError,
  announceSuccess,
  getAriaAttributes,
} from './helpers.js';

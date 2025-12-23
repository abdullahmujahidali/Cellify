// Style types
export type {
  BorderStyle,
  Border,
  CellBorders,
  HorizontalAlignment,
  VerticalAlignment,
  CellAlignment,
  UnderlineStyle,
  CellFont,
  FillPattern,
  GradientType,
  GradientStop,
  CellFill,
  NumberFormatCategory,
  NumberFormat,
  CellProtection,
  CellStyle,
  NamedStyle,
} from './style.types.js';

export {
  BUILTIN_NUMBER_FORMATS,
  DEFAULT_STYLE,
  createSolidFill,
  createBorder,
  createUniformBorders,
} from './style.types.js';

// Cell types
export type {
  CellValueType,
  CellErrorType,
  RichTextRun,
  RichTextValue,
  CellHyperlink,
  CellComment,
  ValidationType,
  ValidationOperator,
  ValidationErrorStyle,
  CellValidation,
  PrimitiveCellValue,
  CellValue,
  CellFormula,
  CellAddress,
  CellReference,
  MergeRange,
  CellData,
  CellStorage,
} from './cell.types.js';

export {
  columnIndexToLetter,
  columnLetterToIndex,
  addressToA1,
  a1ToAddress,
  cellKey,
  parseKey,
  getCellValueType,
} from './cell.types.js';

// Range types
export type {
  RangeDefinition,
  RangeReference,
  ConditionalFormatType,
  ConditionalFormatRule,
  AutoFilter,
  AutoFilterColumn,
  PasteOptions,
  SearchOptions,
  FilterCriteria,
} from './range.types.js';

export {
  parseRangeReference,
  rangeToA1,
  rangesOverlap,
  isCellInRange,
  getRangeIntersection,
  getRangeUnion,
  iterateRange,
  getRangeDimensions,
} from './range.types.js';

export type {
  SheetEvent,
  CellChangeEvent,
  CellStyleChangeEvent,
  RangeChangeEvent,
  CellAddedEvent,
  CellDeletedEvent,
  SheetEventType,
  SheetEventHandler,
  SheetEventMap,
  ChangeRecord,
} from './event.types.js';

---
sidebar_position: 4
---

# Types Reference

Complete TypeScript type definitions for Cellify.

## Cell Types

### CellValue

All possible values a cell can hold.

```typescript
type CellValue = PrimitiveCellValue | RichTextValue | CellErrorType;

type PrimitiveCellValue = string | number | boolean | Date | null;
```

### CellValueType

Value type identifier.

```typescript
type CellValueType = 'string' | 'number' | 'boolean' | 'date' | 'error' | 'formula' | 'null';
```

### CellErrorType

Excel error values.

```typescript
type CellErrorType =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA';
```

### CellAddress

Cell position.

```typescript
interface CellAddress {
  row: number; // 0-based
  col: number; // 0-based
}
```

### CellFormula

Formula definition.

```typescript
interface CellFormula {
  formula: string;       // Formula text without leading '='
  result?: CellValue;    // Cached result
  sharedIndex?: number;  // For shared formulas
}
```

### CellHyperlink

Hyperlink definition.

```typescript
interface CellHyperlink {
  target: string;    // URL, file path, or internal reference
  tooltip?: string;
  display?: string;  // Display text (if different from cell value)
}
```

### CellComment

Cell comment/note.

```typescript
interface CellComment {
  text: string | RichTextValue;
  author?: string;
  visible?: boolean; // Whether comment is always visible
}
```

### RichTextValue

Formatted text with multiple runs.

```typescript
interface RichTextValue {
  richText: RichTextRun[];
}

interface RichTextRun {
  text: string;
  font?: {
    name?: string;
    size?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
  };
}
```

### MergeRange

Merged cell range.

```typescript
interface MergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
```

## Style Types

### CellStyle

Complete cell style definition.

```typescript
interface CellStyle {
  font?: CellFont;
  fill?: CellFill;
  borders?: CellBorders;
  alignment?: CellAlignment;
  numberFormat?: NumberFormat;
  protection?: CellProtection;
}
```

### CellFont

Font styling.

```typescript
interface CellFont {
  name?: string;              // Font family (e.g., "Arial", "Calibri")
  size?: number;              // Size in points
  color?: string;             // Hex color
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
}

type UnderlineStyle = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
```

### CellFill

Fill/background styling.

```typescript
interface CellFill {
  type: 'pattern' | 'gradient';

  // Pattern fill
  pattern?: FillPattern;
  foregroundColor?: string;
  backgroundColor?: string;

  // Gradient fill
  gradientType?: GradientType;
  degree?: number;           // 0-360 for linear
  stops?: GradientStop[];
  left?: number;             // For path gradients
  right?: number;
  top?: number;
  bottom?: number;
}

type FillPattern =
  | 'none' | 'solid'
  | 'darkGray' | 'mediumGray' | 'lightGray' | 'gray125' | 'gray0625'
  | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis'
  | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis';

type GradientType = 'linear' | 'path';

interface GradientStop {
  position: number; // 0 to 1
  color: string;    // Hex color
}
```

### CellBorders

Border styling.

```typescript
interface CellBorders {
  top?: Border;
  right?: Border;
  bottom?: Border;
  left?: Border;
  diagonal?: Border;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

interface Border {
  style: BorderStyle;
  color: string; // Hex color
}

type BorderStyle =
  | 'none' | 'thin' | 'medium' | 'thick'
  | 'dashed' | 'dotted' | 'double' | 'hair'
  | 'mediumDashed' | 'dashDot' | 'mediumDashDot'
  | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
```

### CellAlignment

Text alignment and wrapping.

```typescript
interface CellAlignment {
  horizontal?: HorizontalAlignment;
  vertical?: VerticalAlignment;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;           // 0-255
  textRotation?: number;     // -90 to 90, or 255 for vertical
  readingOrder?: 'contextDependent' | 'leftToRight' | 'rightToLeft';
}

type HorizontalAlignment = 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';

type VerticalAlignment = 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
```

### NumberFormat

Number format definition.

```typescript
interface NumberFormat {
  formatCode: string;           // e.g., "0.00", "$#,##0.00"
  category?: NumberFormatCategory;
}

type NumberFormatCategory =
  | 'general' | 'number' | 'currency' | 'accounting'
  | 'date' | 'time' | 'percentage' | 'fraction'
  | 'scientific' | 'text' | 'custom';
```

### CellProtection

Cell protection settings.

```typescript
interface CellProtection {
  locked?: boolean;
  hidden?: boolean; // Hide formula
}
```

### NamedStyle

Reusable named style.

```typescript
interface NamedStyle {
  name: string;
  style: CellStyle;
  builtIn?: boolean;
}
```

## Data Validation Types

### CellValidation

Data validation rule.

```typescript
interface CellValidation {
  type: ValidationType;
  operator?: ValidationOperator;
  formula1?: string | number | Date;
  formula2?: string | number | Date;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showInputMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  showErrorMessage?: boolean;
  errorStyle?: ValidationErrorStyle;
  errorTitle?: string;
  errorMessage?: string;
}

type ValidationType = 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';

type ValidationOperator =
  | 'between' | 'notBetween'
  | 'equal' | 'notEqual'
  | 'lessThan' | 'lessThanOrEqual'
  | 'greaterThan' | 'greaterThanOrEqual';

type ValidationErrorStyle = 'stop' | 'warning' | 'information';
```

## Range Types

### RangeDefinition

Range specification.

```typescript
interface RangeDefinition {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
```

### AutoFilter

Auto filter configuration.

```typescript
interface AutoFilter {
  range: RangeDefinition;
  // Filter conditions per column can be added
}
```

### ConditionalFormatRule

Conditional formatting rule.

```typescript
interface ConditionalFormatRule {
  range: RangeDefinition;
  type: string;
  priority: number;
  style?: CellStyle;
  // Additional properties based on type
}
```

## Sheet Types

### RowConfig

Row configuration.

```typescript
interface RowConfig {
  height?: number;        // Height in points
  hidden?: boolean;
  outlineLevel?: number;  // For grouping
  style?: CellStyle;      // Default style for row
}
```

### ColumnConfig

Column configuration.

```typescript
interface ColumnConfig {
  width?: number;         // Width in characters
  hidden?: boolean;
  outlineLevel?: number;  // For grouping
  style?: CellStyle;      // Default style for column
}
```

### SheetView

Sheet view settings.

```typescript
interface SheetView {
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showZeros?: boolean;
  tabSelected?: boolean;
  zoomScale?: number;     // 10-400
  frozenRows?: number;
  frozenCols?: number;
  splitRow?: number;
  splitCol?: number;
}
```

### PageSetup

Print settings.

```typescript
interface PageSetup {
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
```

### SheetProtection

Sheet protection options.

```typescript
interface SheetProtection {
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
```

## Workbook Types

### WorkbookProperties

Workbook metadata.

```typescript
interface WorkbookProperties {
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
```

### DefinedName

Named range or constant.

```typescript
interface DefinedName {
  name: string;
  formula: string;    // Range reference or formula
  scope?: string;     // Sheet name for local scope
  comment?: string;
  hidden?: boolean;
}
```

### WorkbookView

Workbook view settings.

```typescript
interface WorkbookView {
  activeSheet?: number;
  firstSheet?: number;
  showSheetTabs?: boolean;
  tabRatio?: number;
}
```

### CalculationMode

Calculation mode.

```typescript
type CalculationMode = 'auto' | 'manual' | 'autoNoTable';
```

## Import/Export Types

### XlsxImportOptions

Options for importing Excel files.

```typescript
interface XlsxImportOptions {
  sheets?: string[] | number[] | 'all';
  importFormulas?: boolean;       // Default: true
  importStyles?: boolean;         // Default: true
  importMergedCells?: boolean;    // Default: true
  importDimensions?: boolean;     // Default: true
  importFreezePanes?: boolean;    // Default: true
  importProperties?: boolean;     // Default: true
  maxRows?: number;               // 0 = unlimited
  maxCols?: number;               // 0 = unlimited
}
```

### XlsxImportResult

Result of importing an Excel file.

```typescript
interface XlsxImportResult {
  workbook: Workbook;
  stats: {
    sheetCount: number;
    totalCells: number;
    formulaCells: number;
    mergedRanges: number;
    durationMs: number;
  };
  warnings: Array<{
    code: string;
    message: string;
    location?: string;
  }>;
}
```

### CsvImportOptions

Options for importing CSV files.

```typescript
interface CsvImportOptions {
  delimiter?: string;
  quoteChar?: string;
  sheetName?: string;
  startCell?: string;
  hasHeaders?: boolean;
  skipEmptyLines?: boolean;
  trimValues?: boolean;
  detectNumbers?: boolean;
  detectDates?: boolean;
  dateFormats?: string[];
  maxRows?: number;
  commentChar?: string;
  onProgress?: (current: number, total: number) => void;
}
```

### CsvExportOptions

Options for exporting to CSV.

```typescript
interface CsvExportOptions {
  delimiter?: string;
  rowDelimiter?: string;
  quoteChar?: string;
  quoteAllFields?: boolean;
  includeBom?: boolean;
  nullValue?: string;
  dateFormat?: string;
  range?: string;
}
```

## Utility Functions

### Column Conversion

```typescript
// Column index to letter (0 -> 'A', 25 -> 'Z', 26 -> 'AA')
function columnIndexToLetter(index: number): string

// Column letter to index ('A' -> 0, 'Z' -> 25, 'AA' -> 26)
function columnLetterToIndex(letter: string): number
```

### Address Conversion

```typescript
// Row/col to A1 notation
function addressToA1(row: number, col: number): string

// A1 notation to row/col
function a1ToAddress(a1: string): CellAddress
```

### Style Helpers

```typescript
// Create a solid fill
function createSolidFill(color: string): CellFill

// Create a border
function createBorder(style: BorderStyle, color?: string): Border

// Create uniform borders on all sides
function createUniformBorders(style: BorderStyle, color?: string): CellBorders
```

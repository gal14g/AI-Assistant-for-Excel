/**
 * Core type definitions for the AI Assistant For Excel execution engine.
 *
 * These types define the plan schema that the LLM planner produces,
 * the executor consumes, and the validator checks.
 *
 * IMPORTANT: The LLM never generates executable code — only typed JSON plans.
 * The executor maps each step to a safe Office.js wrapper.
 */

// ---------------------------------------------------------------------------
// Plan step action types – every supported Excel operation
// ---------------------------------------------------------------------------

export type StepAction =
  | "readRange"
  | "writeValues"
  | "writeFormula"
  | "matchRecords"
  | "groupSum"
  | "createTable"
  | "applyFilter"
  | "sortRange"
  | "createPivot"
  | "createChart"
  | "addConditionalFormat"
  | "cleanupText"
  | "removeDuplicates"
  | "freezePanes"
  | "findReplace"
  | "addValidation"
  | "addSheet"
  | "renameSheet"
  | "deleteSheet"
  | "copySheet"
  | "protectSheet"
  | "autoFitColumns"
  | "mergeCells"
  | "setNumberFormat"
  | "insertDeleteRows"
  | "addSparkline"
  | "formatCells"
  | "clearRange"
  | "hideShow"
  | "addComment"
  | "addHyperlink"
  | "groupRows"
  | "setRowColSize"
  | "copyPasteRange"
  | "pageLayout"
  | "insertPicture"
  | "insertShape"
  | "insertTextBox"
  | "addSlicer"
  | "splitColumn"
  | "unpivot"
  | "crossTabulate"
  | "bulkFormula"
  | "compareSheets"
  | "consolidateRanges"
  | "extractPattern"
  | "categorize"
  | "fillBlanks"
  | "subtotals"
  | "transpose"
  | "namedRange"
  | "fuzzyMatch"
  | "deleteRowsByCondition"
  | "splitByGroup"
  | "lookupAll"
  | "regexReplace"
  | "coerceDataType"
  | "normalizeDates"
  | "deduplicateAdvanced"
  | "joinSheets"
  | "frequencyDistribution"
  | "runningTotal"
  | "rankColumn"
  | "topN"
  | "percentOfTotal"
  | "growthRate"
  | "consolidateAllSheets"
  | "cloneSheetStructure"
  | "addReportHeader"
  | "alternatingRowFormat"
  | "quickFormat"
  | "refreshPivot"
  | "pivotCalculatedField"
  | "addDropdownControl"
  | "conditionalFormula"
  | "spillFormula";

// ---------------------------------------------------------------------------
// Step parameter shapes – one per action
// ---------------------------------------------------------------------------

export interface ReadRangeParams {
  range: string; // e.g. "Sheet1!A1:C20"
  includeHeaders?: boolean;
}

export interface WriteValuesParams {
  range: string;
  /** 2D array of values. Null means skip (preserve existing). */
  values: (string | number | boolean | null)[][];
  /** If true, only write values — never carry formatting. Default true. */
  valuesOnly?: boolean;
}

export interface WriteFormulaParams {
  cell: string; // single cell e.g. "Sheet1!D2"
  formula: string; // e.g. "=VLOOKUP(B2,Sheet1!A:B,2,FALSE)"
  /** Optional: fill down to this many rows */
  fillDown?: number;
}

export interface MatchRecordsParams {
  lookupRange: string; // range containing lookup keys
  sourceRange: string; // range to search in (key column)
  returnColumns?: number[]; // 1-based column offsets to return (default [1]); not required when writeValue is set
  matchType: "exact" | "approximate" | "contains";
  outputRange: string; // where to write results
  /** Prefer native VLOOKUP/XLOOKUP formulas over computed values */
  preferFormula?: boolean;
  /**
   * When set, write this constant string to outputRange for matched rows
   * and empty string for unmatched rows. Forces deterministic JS matching
   * (no formula written). Required for composite multi-column matching.
   */
  writeValue?: string;
}

export interface GroupSumParams {
  dataRange: string;
  groupByColumn: number; // 1-based
  sumColumn: number; // 1-based
  outputRange: string;
  /** Prefer SUMIF/SUMIFS formula over computed values */
  preferFormula?: boolean;
  includeHeaders?: boolean;
}

export interface CreateTableParams {
  range: string;
  tableName?: string; // auto-generated if omitted
  hasHeaders?: boolean;
  style?: string; // e.g. "TableStyleMedium2"
}

export interface ApplyFilterParams {
  tableNameOrRange: string;
  columnIndex: number; // 0-based
  criteria: FilterCriteria;
}

export interface FilterCriteria {
  filterOn: "values" | "topItems" | "custom";
  values?: string[];
  operator?: "greaterThan" | "lessThan" | "equals" | "contains";
  value?: string | number;
}

export interface SortRangeParams {
  range: string;
  sortFields?: SortField[]; // defaults to first column ascending if omitted
  hasHeaders?: boolean;
}

export interface SortField {
  columnIndex: number; // 0-based
  ascending?: boolean;
}

export interface CreatePivotParams {
  sourceRange: string;
  destinationRange?: string; // new sheet created if omitted
  pivotName?: string;        // auto-generated if omitted
  rows?: string[];           // auto-detected from headers if omitted
  columns?: string[];
  values?: PivotValue[];     // auto-detected from headers if omitted
  filters?: string[];
}

export interface PivotValue {
  field: string;
  summarizeBy: "sum" | "count" | "average" | "max" | "min";
  displayName?: string;
}

export interface CreateChartParams {
  dataRange: string;
  chartType:
    | "columnClustered"
    | "columnStacked"
    | "bar"
    | "line"
    | "pie"
    | "area"
    | "scatter"
    | "combo";
  title?: string;
  /** Sheet to place the chart on. Defaults to active sheet. */
  sheetName?: string;
  /** Position in pixels from top-left of sheet */
  position?: { left: number; top: number; width: number; height: number };
  seriesNames?: string[];
}

export interface AddConditionalFormatParams {
  range: string;
  /** "formula" uses a custom Excel formula (e.g. "=$D2=\"\"") to trigger the format */
  ruleType: "cellValue" | "colorScale" | "dataBar" | "iconSet" | "text" | "formula";
  /** For cellValue rules */
  operator?: "greaterThan" | "greaterThanOrEqualTo" | "lessThan" | "lessThanOrEqualTo" | "between" | "notBetween" | "equalTo" | "notEqualTo";
  values?: (string | number)[];
  format?: {
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
  };
  /** For text rules */
  text?: string;
  /** For formula-based rules: an Excel formula string starting with "=" */
  formula?: string;
}

export interface CleanupTextParams {
  range: string;
  operations: CleanupOperation[];
  outputRange?: string; // defaults to in-place
}

export type CleanupOperation =
  | "trim"
  | "lowercase"
  | "uppercase"
  | "properCase"
  | "removeNonPrintable"
  | "normalizeWhitespace";

export interface RemoveDuplicatesParams {
  range: string;
  columnIndexes?: number[]; // 0-based; if omitted, all columns
}

export interface FreezePanesParams {
  /** Cell below-right of the freeze. e.g. "B2" freezes row 1 and column A */
  cell: string;
  sheetName?: string;
}

export interface FindReplaceParams {
  range?: string; // if omitted, entire sheet
  sheetName?: string;
  find: string;
  replace: string;
  matchCase?: boolean;
  matchEntireCell?: boolean;
}

export interface AddValidationParams {
  range: string;
  validationType: "list" | "wholeNumber" | "decimal" | "date" | "textLength" | "custom";
  /** For list validation */
  listValues?: string[];
  /** For numeric / date validation */
  operator?: "between" | "greaterThan" | "lessThan";
  min?: number | string;
  max?: number | string;
  /** Custom formula */
  formula?: string;
  showErrorAlert?: boolean;
  errorMessage?: string;
}

export interface SheetOpParams {
  sheetName: string;
  newName?: string; // for rename
  password?: string; // for protect
}

export interface AutoFitColumnsParams {
  range?: string;
  sheetName?: string;
}

export interface MergeCellsParams {
  range: string;
  across?: boolean;    // merge across rows instead of full merge
  /** Backend Pydantic alias — "mergeAcross" maps to across=true */
  mergeType?: "merge" | "mergeAcross" | "mergeAllCells";
}

export interface SetNumberFormatParams {
  range: string;
  format: string; // e.g. "#,##0.00", "yyyy-mm-dd"
}

export interface InsertDeleteRowsParams {
  /** The range that determines which rows/columns to insert or delete */
  range: string;
  /**
   * "down"  → insert blank rows above (existing rows shift down)
   * "up"    → delete rows (remaining rows shift up)
   * "right" → insert blank columns to the left (existing shift right)
   * "left"  → delete columns (remaining columns shift left)
   */
  shiftDirection: "down" | "up" | "right" | "left";
}

export interface AddSparklineParams {
  /** Data source range — one row per sparkline cell */
  dataRange: string;
  /** Where to place the sparklines (one cell per row of dataRange) */
  locationRange: string;
  /** Sparkline type — defaults to "line" */
  sparklineType?: "line" | "column" | "winLoss";
  /** Optional hex color, e.g. "#0f6cbd" */
  color?: string;
}

export interface FormatCellsParams {
  range: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;       // e.g. "Calibri", "Arial"
  fontColor?: string;        // hex e.g. "#FF0000"
  fillColor?: string;        // hex e.g. "#FFFF00"
  horizontalAlignment?: "left" | "center" | "right" | "justify";
  verticalAlignment?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  borders?: {
    style: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "none";
    color?: string;           // hex
    edges?: ("top" | "bottom" | "left" | "right" | "all" | "outside" | "inside")[];
  };
}

export interface ClearRangeParams {
  range: string;
  /** What to clear: "contents" (values+formulas), "formats", "all" (both) */
  clearType: "contents" | "formats" | "all";
}

export interface HideShowParams {
  /** What to hide/show */
  target: "rows" | "columns" | "sheet";
  /** The range of rows/columns (e.g. "A:C" for columns, "2:5" for rows), or sheet name for target="sheet" */
  rangeOrName: string;
  /** true = hide, false = unhide */
  hide: boolean;
}

export interface AddCommentParams {
  /** Single cell or range — one comment per cell */
  cell: string;
  content: string;
  author?: string;
}

export interface AddHyperlinkParams {
  cell: string;
  url: string;
  /** Display text — defaults to url if omitted */
  displayText?: string;
}

export interface GroupRowsParams {
  /** Range of rows or columns to group, e.g. "3:8" or "B:E" */
  range: string;
  /** "group" adds outline grouping, "ungroup" removes it */
  operation: "group" | "ungroup";
}

export interface SetRowColSizeParams {
  /** Row range (e.g. "2:5") or column range (e.g. "A:C") */
  range: string;
  /** Target dimension: "rowHeight" or "columnWidth" */
  dimension: "rowHeight" | "columnWidth";
  /** Size in points (rowHeight) or characters (columnWidth) */
  size: number;
}

export interface CopyPasteRangeParams {
  /** Source range to copy from */
  sourceRange: string;
  /** Destination range (top-left cell is enough) */
  destinationRange: string;
  /** What to paste: "all" (default), "values", "formats", "formulas" */
  pasteType?: "all" | "values" | "formats" | "formulas";
}

export interface PageLayoutParams {
  sheetName?: string;
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
    header?: number;
    footer?: number;
  };
  orientation?: "portrait" | "landscape";
  paperSize?: string;
  printArea?: string;
  showGridlines?: boolean;
  printGridlines?: boolean;
}

export interface InsertPictureParams {
  sheetName?: string;
  imageBase64: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  altText?: string;
}

export interface InsertShapeParams {
  sheetName?: string;
  shapeType: string;
  left: number;
  top: number;
  width: number;
  height: number;
  fillColor?: string;
  lineColor?: string;
  lineWeight?: number;
  textContent?: string;
}

export interface InsertTextBoxParams {
  sheetName?: string;
  text: string;
  left: number;
  top: number;
  width: number;
  height: number;
  fontSize?: number;
  fontFamily?: string;
  fontColor?: string;
  fillColor?: string;
  horizontalAlignment?: string;
}

export interface AddSlicerParams {
  sheetName?: string;
  sourceType: "pivotTable" | "table";
  sourceName: string;
  sourceField: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export interface SplitColumnParams {
  /** Source column range, e.g. "A:A" or "Sheet1!A1:A100" */
  sourceRange: string;
  /** Delimiter to split on. Use "fixed" for fixed-width splitting. */
  delimiter: string;
  /** Sheet column letter where the first output column is written */
  outputStartColumn: string;
  /** Optional header row label(s) for the new columns */
  outputHeaders?: string[];
  /** How many parts to split into (default 2) */
  parts?: number;
}

export interface UnpivotParams {
  /** Range including headers, e.g. "Sheet1!A1:E20" */
  sourceRange: string;
  /** Number of columns at the left to keep as ID columns (default 1) */
  idColumns: number;
  /** Sheet or cell where the unpivoted table starts */
  outputRange: string;
  /** Header label for the "variable" column (default "Attribute") */
  variableColumnName?: string;
  /** Header label for the "value" column (default "Value") */
  valueColumnName?: string;
}

export interface CrossTabulateParams {
  /** Range of raw data including headers */
  sourceRange: string;
  /** 1-based column index for rows of the output matrix */
  rowField: number;
  /** 1-based column index for columns of the output matrix */
  columnField: number;
  /** 1-based column index for values to aggregate */
  valueField: number;
  /** Aggregation function (default "count") */
  aggregation: "count" | "sum" | "average";
  /** Where to write the cross-tab matrix */
  outputRange: string;
}

export interface BulkFormulaParams {
  /** Formula template for the first data row, e.g. "=A2*B2" */
  formula: string;
  /** Output column range, e.g. "C:C" or "Sheet1!C2:C200" */
  outputRange: string;
  /** Range that defines how many rows to fill (used to detect last row) */
  dataRange: string;
  /** Whether to include a header row (skips row 1 of output, default true) */
  hasHeaders?: boolean;
}

export interface CompareSheetsParams {
  /** First range/sheet to compare */
  rangeA: string;
  /** Second range/sheet to compare */
  rangeB: string;
  /** Where to write the diff report (new sheet created if omitted) */
  outputRange?: string;
  /** Highlight differences in-place on rangeA (default false) */
  highlightDiffs?: boolean;
  /** Background color for highlighting (default "#FFD966" yellow) */
  highlightColor?: string;
}

export interface ConsolidateRangesParams {
  /** Array of source range addresses to merge */
  sourceRanges: string[];
  /** Where to write the consolidated data */
  outputRange: string;
  /** Stack vertically (default) or join horizontally */
  direction?: "vertical" | "horizontal";
  /** Include a source-label column showing which range each row came from */
  addSourceLabel?: boolean;
  /** Whether to de-duplicate rows after consolidating */
  deduplicate?: boolean;
}

export interface ExtractPatternParams {
  /** Source range of text cells */
  sourceRange: string;
  /** Built-in pattern name or a custom regex string */
  pattern: "email" | "phone" | "url" | "date" | "number" | "currency" | string;
  /** Where to write extracted values */
  outputRange: string;
  /** If true, extract ALL matches per cell (comma-joined); default first match only */
  allMatches?: boolean;
}

export interface CategorizeParams {
  /** Source range to categorize */
  sourceRange: string;
  /** Where to write category labels */
  outputRange: string;
  /** Ordered list of rules; first match wins */
  rules: CategorizeRule[];
  /** Value to write when no rule matches (default "") */
  defaultValue?: string;
}

export interface CategorizeRule {
  /** "contains" | "equals" | "startsWith" | "endsWith" | "greaterThan" | "lessThan" | "regex" */
  operator: "contains" | "equals" | "startsWith" | "endsWith" | "greaterThan" | "lessThan" | "regex";
  value: string | number;
  label: string;
}

export interface FillBlanksParams {
  /** Range in which to fill blank cells */
  range: string;
  /** "down" = copy value from the cell above (default), "up", "constant" */
  fillMode?: "down" | "up" | "constant";
  /** Value to fill when fillMode = "constant" */
  constantValue?: string | number;
}

export interface SubtotalsParams {
  /** Data range including headers */
  dataRange: string;
  /** 1-based column index to group by */
  groupByColumn: number;
  /** Columns to subtotal (1-based) */
  subtotalColumns: number[];
  /** Aggregation function (default "sum") */
  aggregation?: "sum" | "count" | "average";
  /** Label to append to each subtotal row (default "Total") */
  subtotalLabel?: string;
}

export interface TransposeParams {
  /** Source range */
  sourceRange: string;
  /** Where to write the transposed data */
  outputRange: string;
  /** Copy formatting too (default false — values only) */
  copyFormatting?: boolean;
}

export interface NamedRangeParams {
  /** Operation: create / update / delete */
  operation: "create" | "update" | "delete";
  /** Name for the range */
  name: string;
  /** Range address (required for create/update) */
  range?: string;
  /** Scope sheet name (default = workbook-level) */
  sheetName?: string;
}

export interface FuzzyMatchParams {
  /** Column with values to match */
  lookupRange: string;
  /** Column to match against */
  sourceRange: string;
  /** Where to write results */
  outputRange: string;
  /** Similarity threshold 0-1 (default 0.7) */
  threshold: number;
  /** Constant to write instead of matched value */
  writeValue?: string;
  /** If true, write the best matching source value */
  returnBestMatch?: boolean;
}

export interface DeleteRowsByConditionParams {
  /** Data range including headers */
  range: string;
  /** 1-based column index to check */
  column: number;
  /** Condition to evaluate */
  condition: "blank" | "notBlank" | "equals" | "notEquals" | "contains" | "greaterThan" | "lessThan";
  /** Value for equals/notEquals/contains/greaterThan/lessThan */
  value?: string | number;
  /** Whether the range has headers (default true — skip header row) */
  hasHeaders?: boolean;
}

export interface SplitByGroupParams {
  /** Data range including headers */
  dataRange: string;
  /** 1-based column index to group by */
  groupByColumn: number;
  /** Include header row in each new sheet (default true) */
  keepHeaders?: boolean;
}

export interface LookupAllParams {
  /** Column with values to match */
  lookupRange: string;
  /** Range to search in */
  sourceRange: string;
  /** 1-based column index in source to return */
  returnColumn: number;
  /** Where to write results */
  outputRange: string;
  /** Delimiter for joining multiple matches (default ", ") */
  delimiter?: string;
  /** Match type (default "exact") */
  matchType?: "exact" | "contains";
}

export interface RegexReplaceParams {
  /** Range to apply replacement */
  range: string;
  /** Regex pattern */
  pattern: string;
  /** Replacement string (supports $1, $2 capture groups) */
  replacement: string;
  /** Regex flags (default "gi") */
  flags?: string;
}

export interface CoerceDataTypeParams {
  range: string;
  targetType: "number" | "text" | "date";
  dateFormat?: string;
  locale?: string;
}

export interface NormalizeDatesParams {
  range: string;
  outputFormat: string;
  inputFormat?: string;
}

export interface DeduplicateAdvancedParams {
  range: string;
  keyColumns: number[];
  keepStrategy: "first" | "last" | "mostComplete" | "newest";
  dateColumn?: number;
  hasHeaders?: boolean;
}

export interface JoinSheetsParams {
  leftRange: string;
  rightRange: string;
  leftKeyColumn: number;
  rightKeyColumn: number;
  joinType: "inner" | "left" | "right" | "full";
  outputRange: string;
}

export interface FrequencyDistributionParams {
  sourceRange: string;
  outputRange: string;
  sortBy?: "value" | "frequency";
  ascending?: boolean;
  includePercent?: boolean;
}

export interface RunningTotalParams {
  sourceRange: string;
  outputRange: string;
  hasHeaders?: boolean;
}

export interface RankColumnParams {
  sourceRange: string;
  outputRange: string;
  order?: "descending" | "ascending";
  hasHeaders?: boolean;
}

export interface TopNParams {
  dataRange: string;
  valueColumn: number;
  n: number;
  position?: "top" | "bottom";
  outputRange: string;
  hasHeaders?: boolean;
}

export interface PercentOfTotalParams {
  sourceRange: string;
  outputRange: string;
  hasHeaders?: boolean;
  formatAsPercent?: boolean;
}

export interface GrowthRateParams {
  sourceRange: string;
  outputRange: string;
  hasHeaders?: boolean;
  formatAsPercent?: boolean;
}

export interface ConsolidateAllSheetsParams {
  outputSheetName?: string;
  hasHeaders?: boolean;
  excludeSheets?: string[];
}

export interface CloneSheetStructureParams {
  sourceSheet: string;
  newSheetName: string;
}

export interface AddReportHeaderParams {
  title: string;
  sheetName?: string;
  range?: string;
  fontSize?: number;
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
}

export interface AlternatingRowFormatParams {
  range: string;
  evenColor?: string;
  oddColor?: string;
  hasHeaders?: boolean;
}

export interface QuickFormatParams {
  range: string;
  freezeHeader?: boolean;
  addFilters?: boolean;
  autoFit?: boolean;
  zebraStripe?: boolean;
  headerColor?: string;
  headerFontColor?: string;
}

export interface RefreshPivotParams {
  pivotName?: string;
  sheetName?: string;
}

export interface PivotCalculatedFieldParams {
  pivotName: string;
  sheetName?: string;
  fieldName: string;
  formula: string;
}

export interface AddDropdownControlParams {
  cell: string;
  listSource: string;
  promptMessage?: string;
  sheetName?: string;
}

export interface ConditionalFormulaParams {
  range: string;
  conditionColumn: number;
  condition: "blank" | "notBlank" | "equals" | "notEquals" | "contains" | "greaterThan" | "lessThan";
  conditionValue?: string | number;
  trueFormula: string;
  falseFormula: string;
  outputRange: string;
  hasHeaders?: boolean;
}

export interface SpillFormulaParams {
  cell: string;
  formula: string;
  sheetName?: string;
}

// ---------------------------------------------------------------------------
// Union of all param types
// ---------------------------------------------------------------------------

export type StepParams =
  | ReadRangeParams
  | WriteValuesParams
  | WriteFormulaParams
  | MatchRecordsParams
  | GroupSumParams
  | CreateTableParams
  | ApplyFilterParams
  | SortRangeParams
  | CreatePivotParams
  | CreateChartParams
  | AddConditionalFormatParams
  | CleanupTextParams
  | RemoveDuplicatesParams
  | FreezePanesParams
  | FindReplaceParams
  | AddValidationParams
  | SheetOpParams
  | AutoFitColumnsParams
  | MergeCellsParams
  | SetNumberFormatParams
  | InsertDeleteRowsParams
  | AddSparklineParams
  | FormatCellsParams
  | ClearRangeParams
  | HideShowParams
  | AddCommentParams
  | AddHyperlinkParams
  | GroupRowsParams
  | SetRowColSizeParams
  | CopyPasteRangeParams
  | PageLayoutParams
  | InsertPictureParams
  | InsertShapeParams
  | InsertTextBoxParams
  | AddSlicerParams
  | SplitColumnParams
  | UnpivotParams
  | CrossTabulateParams
  | BulkFormulaParams
  | CompareSheetsParams
  | ConsolidateRangesParams
  | ExtractPatternParams
  | CategorizeParams
  | FillBlanksParams
  | SubtotalsParams
  | TransposeParams
  | NamedRangeParams
  | FuzzyMatchParams
  | DeleteRowsByConditionParams
  | SplitByGroupParams
  | LookupAllParams
  | RegexReplaceParams
  | CoerceDataTypeParams
  | NormalizeDatesParams
  | DeduplicateAdvancedParams
  | JoinSheetsParams
  | FrequencyDistributionParams
  | RunningTotalParams
  | RankColumnParams
  | TopNParams
  | PercentOfTotalParams
  | GrowthRateParams
  | ConsolidateAllSheetsParams
  | CloneSheetStructureParams
  | AddReportHeaderParams
  | AlternatingRowFormatParams
  | QuickFormatParams
  | RefreshPivotParams
  | PivotCalculatedFieldParams
  | AddDropdownControlParams
  | ConditionalFormulaParams
  | SpillFormulaParams;

// ---------------------------------------------------------------------------
// Plan step
// ---------------------------------------------------------------------------

export interface PlanStep {
  /** Unique step ID within the plan, e.g. "step_1" */
  id: string;
  /** Human-readable description shown in the execution timeline */
  description: string;
  /** The action to perform */
  action: StepAction;
  /** Action-specific parameters */
  params: StepParams;
  /**
   * IDs of steps that must complete before this one.
   * Used for ordering in the executor. If empty, step can run early.
   */
  dependsOn?: string[];
}

// ---------------------------------------------------------------------------
// Execution plan – the top-level object produced by the LLM planner
// ---------------------------------------------------------------------------

export interface ExecutionPlan {
  /** Unique plan ID */
  planId: string;
  /** ISO timestamp of plan creation */
  createdAt: string;
  /** Original natural-language user request */
  userRequest: string;
  /** Human-readable summary of what the plan will do */
  summary: string;
  /** Ordered list of steps */
  steps: PlanStep[];
  /**
   * Whether formatting should be preserved (default true).
   * Only set to false when the user explicitly requests formatting changes.
   */
  preserveFormatting: boolean;
  /** Planner's confidence (0-1) in the plan's correctness */
  confidence: number;
  /** Optional warnings or notes from the planner */
  warnings?: string[];
}

// ---------------------------------------------------------------------------
// Execution state
// ---------------------------------------------------------------------------

export type StepStatus = "pending" | "running" | "success" | "error" | "skipped" | "preview";

export interface StepResult {
  stepId: string;
  status: StepStatus;
  message: string;
  /** Milliseconds taken */
  durationMs?: number;
  /** Data read by the step (for readRange) */
  data?: unknown;
  error?: string;
  /**
   * Output metadata from this step, available for binding in downstream steps.
   * Steps can populate fields like outputRange, sheetName, tableName, etc.
   * Downstream steps reference them via {{step_N.outputRange}}.
   */
  outputs?: Record<string, string | number | boolean>;
}

export interface ExecutionState {
  planId: string;
  status: "idle" | "previewing" | "running" | "completed" | "failed" | "rolledBack";
  stepResults: StepResult[];
  startedAt?: string;
  completedAt?: string;
}

// ---------------------------------------------------------------------------
// Snapshot for rollback
// ---------------------------------------------------------------------------

export interface CellSnapshot {
  range: string;
  values: (string | number | boolean | null)[][];
  /** We store number formats to restore them if needed */
  numberFormats?: string[][];
}

export interface PlanSnapshot {
  planId: string;
  timestamp: string;
  cells: CellSnapshot[];
}

// ---------------------------------------------------------------------------
// Chat types
// ---------------------------------------------------------------------------

export interface RangeToken {
  id: string;
  /** Display text, e.g. "[[Sheet1!A1:C20]]" */
  display: string;
  /** Normalized range address */
  address: string;
  /** Sheet name */
  sheetName: string;
}

export interface PlanOption {
  optionLabel: string;
  plan: ExecutionPlan;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant" | "system";
  content: string;
  /** Range tokens extracted from [[...]] markers in the message text */
  rangeTokens?: { address: string; sheetName: string }[];
  /** Associated plan, if any */
  plan?: ExecutionPlan;
  /** Execution state, if any */
  execution?: ExecutionState;
  /** Progress log from the plan execution, if any */
  progressLog?: { stepId: string; message: string; timestamp: string }[];
  timestamp: string;
}

// ---------------------------------------------------------------------------
// Capability registry types
// ---------------------------------------------------------------------------

export interface CapabilityMeta {
  action: StepAction;
  description: string;
  /** Whether this action writes to the workbook (needs snapshot) */
  mutates: boolean;
  /** Whether this action changes formatting */
  affectsFormatting: boolean;
  /** Office.js API set requirement, e.g. "ExcelApi 1.9" */
  requiresApiSet?: string;
}

export type CapabilityHandler = (
  context: Excel.RequestContext,
  params: StepParams,
  options: ExecutionOptions
) => Promise<StepResult>;

export interface ExecutionOptions {
  dryRun: boolean;
  preserveFormatting: boolean;
  onProgress?: (message: string) => void;
}

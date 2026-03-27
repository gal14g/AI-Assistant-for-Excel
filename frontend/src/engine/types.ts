/**
 * Core type definitions for the Excel AI Copilot execution engine.
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
  | "setNumberFormat";

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
  returnColumns: number[]; // 1-based column offsets to return
  matchType: "exact" | "approximate";
  outputRange: string; // where to write results
  /** Prefer native VLOOKUP/XLOOKUP formulas over computed values */
  preferFormula?: boolean;
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
  tableName: string;
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
  sortFields: SortField[];
  hasHeaders?: boolean;
}

export interface SortField {
  columnIndex: number; // 0-based
  ascending?: boolean;
}

export interface CreatePivotParams {
  sourceRange: string;
  destinationRange: string; // top-left cell
  pivotName: string;
  rows: string[]; // field names
  columns?: string[];
  values: PivotValue[];
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
  ruleType: "cellValue" | "colorScale" | "dataBar" | "iconSet" | "text";
  /** For cellValue rules */
  operator?: "greaterThan" | "lessThan" | "between" | "equalTo";
  values?: (string | number)[];
  format?: {
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
  };
  /** For text rules */
  text?: string;
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
  action: "add" | "rename" | "delete" | "copy" | "protect";
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
  across?: boolean; // merge across rows instead of full merge
}

export interface SetNumberFormatParams {
  range: string;
  format: string; // e.g. "#,##0.00", "yyyy-mm-dd"
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
  | SetNumberFormatParams;

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

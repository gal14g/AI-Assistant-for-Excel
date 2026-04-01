/**
 * Plan Validator
 *
 * Validates an ExecutionPlan before it is executed. Checks:
 * 1. Schema validity (required fields, correct types)
 * 2. Action existence in the capability registry
 * 3. Business rules (formatting safety, dependency ordering, range validity)
 * 4. Step dependency graph is acyclic
 *
 * The validator returns a list of errors/warnings. If errors exist, the plan
 * MUST NOT be executed. Warnings are informational.
 */

import { ExecutionPlan, PlanStep } from "./types";
import { registry } from "./capabilityRegistry";

export interface ValidationResult {
  valid: boolean;
  errors: ValidationIssue[];
  warnings: ValidationIssue[];
}

export interface ValidationIssue {
  stepId?: string;
  field?: string;
  message: string;
  code: string;
}

/** Validate an execution plan. */
export function validatePlan(plan: ExecutionPlan): ValidationResult {
  const errors: ValidationIssue[] = [];
  const warnings: ValidationIssue[] = [];

  // Top-level required fields
  if (!plan.planId) {
    errors.push({ message: "Plan must have a planId", code: "MISSING_PLAN_ID" });
  }
  if (!plan.steps || plan.steps.length === 0) {
    errors.push({ message: "Plan must have at least one step", code: "NO_STEPS" });
  }
  if (typeof plan.preserveFormatting !== "boolean") {
    warnings.push({
      message: "preserveFormatting not set; defaulting to true",
      code: "DEFAULT_PRESERVE_FMT",
    });
  }
  if (plan.confidence !== undefined && (plan.confidence < 0 || plan.confidence > 1)) {
    warnings.push({
      message: `Confidence ${plan.confidence} is outside [0,1]`,
      code: "INVALID_CONFIDENCE",
    });
  }

  if (!plan.steps) return { valid: errors.length === 0, errors, warnings };

  const stepIds = new Set<string>();

  for (const step of plan.steps) {
    validateStep(step, stepIds, plan.preserveFormatting, errors, warnings);
    stepIds.add(step.id);
  }

  // Check dependency graph is acyclic
  const cycleError = detectCycles(plan.steps);
  if (cycleError) {
    errors.push(cycleError);
  }

  return { valid: errors.length === 0, errors, warnings };
}

function validateStep(
  step: PlanStep,
  existingIds: Set<string>,
  preserveFormatting: boolean,
  errors: ValidationIssue[],
  warnings: ValidationIssue[]
): void {
  // Required fields
  if (!step.id) {
    errors.push({ message: "Step missing id", code: "MISSING_STEP_ID" });
    return;
  }
  if (existingIds.has(step.id)) {
    errors.push({
      stepId: step.id,
      message: `Duplicate step id: ${step.id}`,
      code: "DUPLICATE_STEP_ID",
    });
  }
  if (!step.action) {
    errors.push({
      stepId: step.id,
      message: "Step missing action",
      code: "MISSING_ACTION",
    });
    return;
  }
  if (!step.params) {
    errors.push({
      stepId: step.id,
      message: "Step missing params",
      code: "MISSING_PARAMS",
    });
    return;
  }

  // Check action is registered
  if (!registry.has(step.action)) {
    errors.push({
      stepId: step.id,
      message: `Unknown action: ${step.action}`,
      code: "UNKNOWN_ACTION",
    });
  }

  // Check dependencies reference valid step IDs
  if (step.dependsOn) {
    for (const dep of step.dependsOn) {
      if (!existingIds.has(dep)) {
        errors.push({
          stepId: step.id,
          message: `Dependency "${dep}" not found (must be defined before this step)`,
          code: "INVALID_DEPENDENCY",
        });
      }
    }
  }

  // Formatting safety: if preserveFormatting is true, warn on formatting actions
  if (preserveFormatting) {
    const meta = registry.getMeta(step.action);
    if (meta?.affectsFormatting) {
      warnings.push({
        stepId: step.id,
        message: `Action "${step.action}" affects formatting but preserveFormatting is true`,
        code: "FORMAT_SAFETY_WARNING",
      });
    }
  }

  // Action-specific param validation
  validateActionParams(step, errors);
}

/**
 * Human-readable descriptions for every required param field.
 * Shown in validation errors so the user/developer understands what is wrong.
 */
const FIELD_DESCRIPTIONS: Record<string, string> = {
  // Common
  range:            "the cell range to operate on (e.g. \"Sheet1!A1:C20\" or \"Sheet1!A:A\")",
  cell:             "the target cell or freeze-point cell (e.g. \"Sheet1!D2\"; for freezePanes \"B2\" freezes row 1 and column A)",
  formula:          "the Excel formula to write — must start with \"=\"",
  sheetName:        "the name of the worksheet to operate on",

  // writeValues
  values:           "a 2D array of cell values to write, e.g. [[\"Name\",\"Age\"],[\"Alice\",30]]",

  // matchRecords
  lookupRange:      "the range containing the lookup keys (the column you want to match FROM)",
  sourceRange:      "the data range to search in — for matchRecords: the key column; for createPivot: the full table including headers",
  returnColumns:    "1-based column offsets within sourceRange to return, e.g. [2] returns the 2nd column",
  outputRange:      "the destination range where matched/computed results will be written",

  // groupSum
  dataRange:        "the full data range including both the grouping column and the values column (also used as chart data range)",
  groupByColumn:    "1-based index of the column to group by (e.g. 1 = first column)",
  sumColumn:        "1-based index of the column whose values will be summed",

  // createChart
  chartType:        "the chart type — one of: columnClustered, bar, line, pie, area, scatter",

  // applyFilter
  tableNameOrRange: "the name of an Excel Table or a range address to apply the filter to",
  criteria:         "the filter criteria object — must include filterOn (\"values\", \"topItems\", or \"custom\")",

  // addConditionalFormat
  ruleType:         "the type of rule — one of: cellValue, colorScale, dataBar, iconSet, text",

  // cleanupText
  operations:       "list of text operations to apply — e.g. [\"trim\", \"uppercase\", \"normalizeWhitespace\"]",

  // findReplace
  find:             "the text to search for in the sheet",
  replace:          "the replacement text (use \"\" to delete matches)",

  // addValidation
  validationType:   "the type of validation — one of: list, wholeNumber, decimal, date, textLength, custom",
};

function validateActionParams(
  step: PlanStep,
  errors: ValidationIssue[]
): void {
  const p = step.params as Record<string, unknown>;

  switch (step.action) {
    case "readRange":
      requireField(step.id, p, "range", errors);
      break;
    case "writeValues":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "values", errors);
      if (p.values !== undefined) {
        const isFlat = Array.isArray(p.values) && p.values.length > 0 && !Array.isArray((p.values as unknown[])[0]);
        if (!Array.isArray(p.values) || isFlat) {
          errors.push({
            stepId: step.id,
            field: "values",
            message: "\"values\" must be a 2D array, e.g. [[\"Alice\", 30], [\"Bob\", 25]]",
            code: "INVALID_VALUES",
          });
        }
      }
      break;
    case "writeFormula":
      requireField(step.id, p, "cell", errors);
      requireField(step.id, p, "formula", errors);
      break;
    case "matchRecords":
      requireField(step.id, p, "lookupRange", errors);
      requireField(step.id, p, "sourceRange", errors);
      // returnColumns is not required when writeValue is set (composite key mode)
      if (!p.writeValue) {
        requireField(step.id, p, "returnColumns", errors);
      }
      requireField(step.id, p, "outputRange", errors);
      break;
    case "groupSum":
      requireField(step.id, p, "dataRange", errors);
      requireField(step.id, p, "groupByColumn", errors);
      requireField(step.id, p, "sumColumn", errors);
      requireField(step.id, p, "outputRange", errors);
      break;
    case "createTable":
      // tableName is optional — handler auto-generates it
      requireField(step.id, p, "range", errors);
      break;
    case "applyFilter":
      requireField(step.id, p, "tableNameOrRange", errors);
      requireField(step.id, p, "criteria", errors);
      break;
    case "sortRange":
      // sortFields is optional — handler defaults to first column ascending
      requireField(step.id, p, "range", errors);
      break;
    case "createPivot":
      // Only sourceRange is truly required — handler auto-detects everything else
      requireField(step.id, p, "sourceRange", errors);
      break;
    case "createChart":
      requireField(step.id, p, "dataRange", errors);
      requireField(step.id, p, "chartType", errors);
      break;
    case "addConditionalFormat":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "ruleType", errors);
      if (p.ruleType === "formula" && !p.formula && !p.text) {
        errors.push({
          stepId: step.id,
          field: "formula",
          message: "ruleType=\"formula\" requires a \"formula\" field (e.g. \"=$D2=\\\"\\\"\")",
          code: "MISSING_FIELD",
        });
      }
      break;
    case "cleanupText":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "operations", errors);
      break;
    case "removeDuplicates":
      requireField(step.id, p, "range", errors);
      break;
    case "freezePanes":
      requireField(step.id, p, "cell", errors);
      break;
    case "findReplace":
      requireField(step.id, p, "find", errors);
      requireField(step.id, p, "replace", errors);
      break;
    case "addValidation":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "validationType", errors);
      break;
    case "addSheet":
    case "renameSheet":
    case "deleteSheet":
    case "copySheet":
    case "protectSheet":
      requireField(step.id, p, "sheetName", errors);
      break;
    case "formatCells":
      requireField(step.id, p, "range", errors);
      break;
    case "clearRange":
      requireField(step.id, p, "range", errors);
      break;
    case "hideShow":
      requireField(step.id, p, "target", errors);
      requireField(step.id, p, "rangeOrName", errors);
      break;
    case "addComment":
      requireField(step.id, p, "cell", errors);
      requireField(step.id, p, "content", errors);
      break;
    case "addHyperlink":
      requireField(step.id, p, "cell", errors);
      requireField(step.id, p, "url", errors);
      break;
    case "groupRows":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "operation", errors);
      break;
    case "setRowColSize":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "dimension", errors);
      requireField(step.id, p, "size", errors);
      break;
    case "copyPasteRange":
      requireField(step.id, p, "sourceRange", errors);
      requireField(step.id, p, "destinationRange", errors);
      break;
    case "pageLayout":
      if (
        p.margins === undefined &&
        p.orientation === undefined &&
        p.paperSize === undefined &&
        p.printArea === undefined &&
        p.showGridlines === undefined &&
        p.printGridlines === undefined
      ) {
        errors.push({
          stepId: step.id,
          message: "pageLayout requires at least one of: margins, orientation, paperSize, printArea, showGridlines, printGridlines",
          code: "MISSING_FIELD",
        });
      }
      break;
    case "insertPicture":
      requireField(step.id, p, "imageBase64", errors);
      break;
    case "insertShape":
      requireField(step.id, p, "shapeType", errors);
      requireField(step.id, p, "left", errors);
      requireField(step.id, p, "top", errors);
      requireField(step.id, p, "width", errors);
      requireField(step.id, p, "height", errors);
      break;
    case "insertTextBox":
      requireField(step.id, p, "text", errors);
      requireField(step.id, p, "left", errors);
      requireField(step.id, p, "top", errors);
      requireField(step.id, p, "width", errors);
      requireField(step.id, p, "height", errors);
      break;
    case "addSlicer":
      requireField(step.id, p, "sourceName", errors);
      requireField(step.id, p, "sourceField", errors);
      break;
    case "addSparkline":
      requireField(step.id, p, "dataRange", errors);
      requireField(step.id, p, "locationRange", errors);
      break;
    case "autoFitColumns":
      // range is optional — handler defaults to used range
      break;
    case "insertDeleteRows":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "shiftDirection", errors);
      break;
    case "mergeCells":
      requireField(step.id, p, "range", errors);
      break;
    case "setNumberFormat":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "format", errors);
      break;
    default:
      // Extensible: unknown actions are caught by the registry check above
      break;
  }
}

function requireField(
  stepId: string,
  params: Record<string, unknown>,
  field: string,
  errors: ValidationIssue[]
): void {
  if (params[field] === undefined || params[field] === null) {
    const description = FIELD_DESCRIPTIONS[field];
    const detail = description ? ` — ${description}` : "";
    errors.push({
      stepId,
      field,
      message: `Missing required field "${field}"${detail}`,
      code: "MISSING_FIELD",
    });
  }
}

/**
 * Detect cycles in the step dependency graph using DFS.
 */
function detectCycles(steps: PlanStep[]): ValidationIssue | null {
  const adjacency = new Map<string, string[]>();
  for (const step of steps) {
    adjacency.set(step.id, step.dependsOn ?? []);
  }

  const visited = new Set<string>();
  const inStack = new Set<string>();

  for (const step of steps) {
    if (hasCycleDFS(step.id, adjacency, visited, inStack)) {
      return {
        message: `Dependency cycle detected involving step "${step.id}"`,
        code: "DEPENDENCY_CYCLE",
      };
    }
  }
  return null;
}

function hasCycleDFS(
  node: string,
  adj: Map<string, string[]>,
  visited: Set<string>,
  inStack: Set<string>
): boolean {
  if (inStack.has(node)) return true;
  if (visited.has(node)) return false;

  visited.add(node);
  inStack.add(node);

  for (const dep of adj.get(node) ?? []) {
    if (hasCycleDFS(dep, adj, visited, inStack)) return true;
  }

  inStack.delete(node);
  return false;
}

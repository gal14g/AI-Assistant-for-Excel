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

import { ExecutionPlan, PlanStep, StepAction } from "./types";
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
      if (p.values && !Array.isArray(p.values)) {
        errors.push({
          stepId: step.id,
          field: "values",
          message: "values must be a 2D array",
          code: "INVALID_VALUES",
        });
      }
      break;
    case "writeFormula":
      requireField(step.id, p, "cell", errors);
      requireField(step.id, p, "formula", errors);
      break;
    case "matchRecords":
      requireField(step.id, p, "lookupRange", errors);
      requireField(step.id, p, "sourceRange", errors);
      requireField(step.id, p, "returnColumns", errors);
      requireField(step.id, p, "outputRange", errors);
      break;
    case "groupSum":
      requireField(step.id, p, "dataRange", errors);
      requireField(step.id, p, "groupByColumn", errors);
      requireField(step.id, p, "sumColumn", errors);
      requireField(step.id, p, "outputRange", errors);
      break;
    case "createTable":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "tableName", errors);
      break;
    case "applyFilter":
      requireField(step.id, p, "tableNameOrRange", errors);
      requireField(step.id, p, "criteria", errors);
      break;
    case "sortRange":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "sortFields", errors);
      break;
    case "createPivot":
      requireField(step.id, p, "sourceRange", errors);
      requireField(step.id, p, "destinationRange", errors);
      requireField(step.id, p, "pivotName", errors);
      requireField(step.id, p, "rows", errors);
      requireField(step.id, p, "values", errors);
      break;
    case "createChart":
      requireField(step.id, p, "dataRange", errors);
      requireField(step.id, p, "chartType", errors);
      break;
    case "addConditionalFormat":
      requireField(step.id, p, "range", errors);
      requireField(step.id, p, "ruleType", errors);
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
    errors.push({
      stepId,
      field,
      message: `Missing required field: ${field}`,
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

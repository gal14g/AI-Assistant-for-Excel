/**
 * Execution Engine
 *
 * Runs a validated ExecutionPlan against the live workbook via Office.js.
 *
 * Key design:
 * - Steps execute in dependency order (topological sort).
 * - Before any mutating step, a snapshot is captured for rollback.
 * - Supports dry-run (preview) mode: runs read steps, simulates writes.
 * - Streams progress updates via a callback.
 * - On failure, stops and offers rollback of completed steps.
 *
 * Office.js notes:
 * - All Excel interactions happen inside Excel.run() which provides a
 *   RequestContext. We must batch operations and call context.sync()
 *   judiciously to avoid excessive round-trips.
 * - Proxy objects are only valid within their Excel.run callback.
 */

import {
  ExecutionPlan,
  ExecutionState,
  StepResult,
  PlanStep,
  ExecutionOptions,
} from "./types";
import { registry } from "./capabilityRegistry";
import { captureSnapshotBatched, createEmptySnapshot, rollbackPlan } from "./snapshot";
import { validatePlan } from "./validator";
import { meetsRequirement, parseApiRequirement } from "./apiSupport";

// ---------------------------------------------------------------------------
// Execution audit log — persistent structured failure data for debugging
// ---------------------------------------------------------------------------

export interface ExecutionAuditEntry {
  timestamp: string;
  planId: string;
  userRequest: string;
  status: "completed" | "failed" | "validation_error";
  totalSteps: number;
  completedSteps: number;
  failedStepId?: string;
  failedAction?: string;
  failedError?: string;
  /** Full step results for detailed analysis */
  stepResults: {
    stepId: string;
    action: string;
    status: string;
    message: string;
    durationMs?: number;
    error?: string;
    outputKeys?: string[];
  }[];
  durationMs: number;
}

const AUDIT_LOG_KEY = "excelCopilot_executionAudit";
const MAX_AUDIT_ENTRIES = 50;

/** Retrieve the execution audit log from localStorage. */
export function getAuditLog(): ExecutionAuditEntry[] {
  try {
    const raw = localStorage.getItem(AUDIT_LOG_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

/** Clear the audit log. */
export function clearAuditLog(): void {
  localStorage.removeItem(AUDIT_LOG_KEY);
}

function appendAuditEntry(entry: ExecutionAuditEntry): void {
  try {
    const log = getAuditLog();
    log.push(entry);
    // Keep only the most recent entries
    while (log.length > MAX_AUDIT_ENTRIES) log.shift();
    localStorage.setItem(AUDIT_LOG_KEY, JSON.stringify(log));
  } catch {
    // localStorage may be unavailable in some contexts — don't crash
  }
}

/** Get only failed executions from the audit log. */
export function getFailedExecutions(): ExecutionAuditEntry[] {
  return getAuditLog().filter((e) => e.status === "failed" || e.status === "validation_error");
}

/** Get a human-readable summary of failures for debugging. */
export function getFailureSummary(): string {
  const failures = getFailedExecutions();
  if (failures.length === 0) return "No execution failures recorded.";
  return failures
    .map((f) => {
      const stepInfo = f.failedStepId
        ? `step ${f.failedStepId} (${f.failedAction}): ${f.failedError}`
        : "validation error";
      return `[${f.timestamp}] Plan "${f.userRequest.slice(0, 60)}" — ${stepInfo}`;
    })
    .join("\n");
}

// ---------------------------------------------------------------------------
// Step output binding — resolve {{step_N.field}} references in params
// ---------------------------------------------------------------------------

/** Regex matching a binding token like {{step_1.outputRange}} */
const BINDING_RE = /\{\{(step_\d+)\.(\w+)\}\}/g;

/**
 * Deep-clone params and replace any {{step_N.field}} tokens with the
 * actual value from that step's result.outputs.
 *
 * Throws a clear error when a binding cannot be resolved instead of
 * silently leaving the literal token in place — Office.js's "invalid
 * range" error for a literal `{{step_1.outputRange}}` is unactionable,
 * but `Cannot resolve {{step_1.outputRange}}: step_1 did not produce
 * 'outputRange'` tells the user exactly what's wrong.
 */
function resolveBindings(
  params: Record<string, unknown>,
  resultsMap: Map<string, StepResult>,
): Record<string, unknown> {
  const json = JSON.stringify(params);
  const errors: string[] = [];
  const resolved = json.replace(BINDING_RE, (_match, stepId: string, field: string) => {
    const result = resultsMap.get(stepId);
    if (!result) {
      errors.push(`${_match}: step '${stepId}' has not run (missing from plan or earlier failure)`);
      return _match;
    }
    if (!result.outputs || !(field in result.outputs)) {
      const available = result.outputs ? Object.keys(result.outputs).join(", ") || "none" : "none";
      errors.push(`${_match}: step '${stepId}' did not produce '${field}' (available: ${available})`);
      return _match;
    }
    const val = result.outputs[field];
    // If the replacement sits inside a JSON string, return the raw value.
    // If it's a number/boolean that was the entire string value, we still
    // return the string form because it's embedded in a JSON string token.
    return String(val).replace(/"/g, '\\"');
  });
  if (errors.length > 0) {
    throw new Error(`Unresolved step binding(s): ${errors.join("; ")}`);
  }
  return JSON.parse(resolved) as Record<string, unknown>;
}

/**
 * Build an actionable error for an action that isn't in the registry.
 *
 * The validator catches truly invalid action names. If we still get here,
 * it means the action *is* schema-valid but no handler is wired up — almost
 * always a bundle/version mismatch between the backend planner and the
 * frontend add-in. Suggest similarly-named registered actions to help the
 * user spot stale builds.
 *
 * Exported for unit testing.
 */
export function buildUnknownActionError(action: string): string {
  const known = registry.listActions();
  const suggestions = findSimilarActions(action, known, 3);
  const suggestionText = suggestions.length > 0
    ? ` Did you mean: ${suggestions.join(", ")}?`
    : "";
  return (
    `No handler registered for action "${action}". ` +
    `This usually means the add-in bundle is out of date — the backend ` +
    `planner knows about an action that the frontend hasn't shipped a ` +
    `handler for yet. Try reloading the add-in (Ctrl+F5).` +
    suggestionText
  );
}

/** Tiny edit-distance based similarity ranker — no external deps. */
function findSimilarActions(target: string, candidates: string[], limit: number): string[] {
  const t = target.toLowerCase();
  const scored = candidates
    .map((c) => ({ name: c, dist: levenshtein(t, c.toLowerCase()) }))
    // Only suggest reasonably-close matches (within ~40% of target length)
    .filter((x) => x.dist <= Math.max(2, Math.floor(target.length * 0.4)))
    .sort((a, b) => a.dist - b.dist)
    .slice(0, limit);
  return scored.map((s) => s.name);
}

function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  let prev = new Array(b.length + 1);
  let curr = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[b.length];
}

export interface ExecutorCallbacks {
  onStepStart: (step: PlanStep) => void;
  onStepComplete: (step: PlanStep, result: StepResult) => void;
  onProgress: (stepId: string, message: string) => void;
  onPlanComplete: (state: ExecutionState) => void;
  onError: (stepId: string, error: string) => void;
}

/**
 * Execute a plan with full lifecycle management.
 */
export async function executePlan(
  plan: ExecutionPlan,
  callbacks: ExecutorCallbacks,
  options: { dryRun?: boolean } = {}
): Promise<ExecutionState> {
  const execStart = Date.now();

  // Validate first
  const validation = validatePlan(plan);
  if (!validation.valid) {
    const errorMsg = validation.errors.map((e) => e.message).join("; ");
    // Log validation failures to the audit log
    appendAuditEntry({
      timestamp: new Date().toISOString(),
      planId: plan.planId,
      userRequest: plan.userRequest,
      status: "validation_error",
      totalSteps: plan.steps.length,
      completedSteps: 0,
      failedError: errorMsg,
      stepResults: plan.steps.map((s) => ({
        stepId: s.id,
        action: s.action,
        status: "skipped",
        message: "Skipped due to validation error",
      })),
      durationMs: Date.now() - execStart,
    });
    throw new Error(`Plan validation failed: ${errorMsg}`);
  }

  const state: ExecutionState = {
    planId: plan.planId,
    status: options.dryRun ? "previewing" : "running",
    stepResults: [],
    startedAt: new Date().toISOString(),
  };

  const execOptions: ExecutionOptions = {
    dryRun: options.dryRun ?? false,
    preserveFormatting: plan.preserveFormatting ?? true,
    onProgress: undefined, // set per-step below
  };

  // Topological order (steps are already ordered, but respect dependsOn)
  const orderedSteps = topologicalSort(plan.steps);

  const completedSteps = new Set<string>();
  /** Map of stepId → result for output binding between steps */
  const stepResultsMap = new Map<string, StepResult>();

  for (const step of orderedSteps) {
    // Check dependencies are met
    if (step.dependsOn) {
      const unmet = step.dependsOn.filter((d) => !completedSteps.has(d));
      if (unmet.length > 0) {
        const result: StepResult = {
          stepId: step.id,
          status: "skipped",
          message: `Skipped: unmet dependencies [${unmet.join(", ")}]`,
        };
        state.stepResults.push(result);
        stepResultsMap.set(step.id, result);
        callbacks.onStepComplete(step, result);
        continue;
      }
    }

    // Resolve {{step_N.field}} bindings in params before execution
    let resolvedStep = step;
    const stepHasBindings = JSON.stringify(step.params).includes("{{step_");
    if (stepHasBindings) {
      try {
        const resolvedParams = resolveBindings(
          step.params as Record<string, unknown>,
          stepResultsMap,
        );
        resolvedStep = { ...step, params: resolvedParams as typeof step.params };
      } catch (err) {
        // Binding resolution failed — fail this step with an actionable
        // message instead of passing literal {{...}} tokens to Office.js,
        // which would produce a cryptic "invalid range" error.
        const errorMsg = err instanceof Error ? err.message : String(err);
        const result: StepResult = {
          stepId: step.id,
          status: "error",
          message: errorMsg,
          error: errorMsg,
        };
        state.stepResults.push(result);
        stepResultsMap.set(step.id, result);
        state.status = "failed";
        callbacks.onError(step.id, errorMsg);
        callbacks.onStepComplete(step, result);
        break;
      }
    }

    callbacks.onStepStart(resolvedStep);

    const primaryHandler = registry.getHandler(resolvedStep.action);
    if (!primaryHandler) {
      // Reaching here usually means: validator accepted the action (so the
      // name is in the schema) but no handler is wired up in the registry.
      // That's almost always a stale add-in bundle — capabilities/index.ts
      // didn't import the handler module, or the user is on an older build
      // than the backend planner. Surface that explicitly.
      const errorMsg = buildUnknownActionError(resolvedStep.action);
      const result: StepResult = {
        stepId: resolvedStep.id,
        status: "error",
        message: errorMsg,
        error: errorMsg,
      };
      state.stepResults.push(result);
      stepResultsMap.set(resolvedStep.id, result);
      state.status = "failed";
      callbacks.onError(resolvedStep.id, errorMsg);
      callbacks.onStepComplete(resolvedStep, result);
      break;
    }

    // ── API-support dispatch ────────────────────────────────────────────────
    // If the handler declares a `requiresApiSet` that this Excel build does
    // not satisfy, try the registered fallback (e.g. Excel 2016 running a
    // PivotTable plan → SUMIFS summary sheet). If no fallback is registered,
    // surface a structured "feature unavailable" error so the UI can show a
    // helpful suggestion rather than letting Office.js throw cryptically mid-
    // step.
    const meta = registry.getMeta(resolvedStep.action);
    let handler = primaryHandler;
    let usingFallback = false;
    if (meta?.requiresApiSet && !meetsRequirement(meta.requiresApiSet)) {
      const fallback = registry.getFallback(resolvedStep.action);
      if (fallback) {
        handler = fallback;
        usingFallback = true;
        callbacks.onProgress(
          resolvedStep.id,
          `Using compatibility fallback (requires ${meta.requiresApiSet}).`,
        );
      } else {
        const parsed = parseApiRequirement(meta.requiresApiSet);
        const needed = parsed ? `${parsed.setName} ${parsed.version}+` : meta.requiresApiSet;
        const errorMsg =
          `"${resolvedStep.action}" requires ${needed}, which this Excel ` +
          `version doesn't support. Upgrade to Excel 2019/2021/Microsoft 365, ` +
          `or ask the assistant for an alternative approach.`;
        const result: StepResult = {
          stepId: resolvedStep.id,
          status: "error",
          message: errorMsg,
          error: errorMsg,
        };
        state.stepResults.push(result);
        stepResultsMap.set(resolvedStep.id, result);
        state.status = "failed";
        callbacks.onError(resolvedStep.id, errorMsg);
        callbacks.onStepComplete(resolvedStep, result);
        break;
      }
    }

    try {
      const result = await executeStep(resolvedStep, handler, execOptions, callbacks, plan.planId);
      if (usingFallback && result.status === "success") {
        result.message = `${result.message} (compatibility fallback)`;
      }
      state.stepResults.push(result);
      stepResultsMap.set(resolvedStep.id, result);
      callbacks.onStepComplete(resolvedStep, result);

      if (result.status === "error") {
        state.status = "failed";
        callbacks.onError(resolvedStep.id, result.error ?? "Unknown error");
        break;
      }

      completedSteps.add(resolvedStep.id);
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      const result: StepResult = {
        stepId: resolvedStep.id,
        status: "error",
        message: `Exception: ${errorMsg}`,
        error: errorMsg,
      };
      state.stepResults.push(result);
      stepResultsMap.set(resolvedStep.id, result);
      state.status = "failed";
      callbacks.onError(resolvedStep.id, errorMsg);
      callbacks.onStepComplete(resolvedStep, result);
      break;
    }
  }

  if (state.status === "running") {
    state.status = "completed";
  } else if (state.status === "previewing") {
    state.status = "completed";
  }

  state.completedAt = new Date().toISOString();

  // Write audit log entry
  const failedResult = state.stepResults.find((r) => r.status === "error");
  const failedStep = failedResult
    ? plan.steps.find((s) => s.id === failedResult.stepId)
    : undefined;
  if (!options.dryRun) {
    appendAuditEntry({
      timestamp: new Date().toISOString(),
      planId: plan.planId,
      userRequest: plan.userRequest,
      status: state.status === "failed" ? "failed" : "completed",
      totalSteps: plan.steps.length,
      completedSteps: state.stepResults.filter((r) => r.status === "success").length,
      failedStepId: failedResult?.stepId,
      failedAction: failedStep?.action,
      failedError: failedResult?.error,
      stepResults: state.stepResults.map((r) => {
        const step = plan.steps.find((s) => s.id === r.stepId);
        return {
          stepId: r.stepId,
          action: step?.action ?? "unknown",
          status: r.status,
          message: r.message,
          durationMs: r.durationMs,
          error: r.error,
          outputKeys: r.outputs ? Object.keys(r.outputs) : undefined,
        };
      }),
      durationMs: Date.now() - execStart,
    });
  }

  callbacks.onPlanComplete(state);
  return state;
}

/**
 * Execute a single step within an Excel.run context.
 */
async function executeStep(
  step: PlanStep,
  handler: import("./types").CapabilityHandler,
  options: ExecutionOptions,
  callbacks: ExecutorCallbacks,
  planId: string
): Promise<StepResult> {
  const startTime = Date.now();

  // Set up per-step progress callback
  const stepOptions: ExecutionOptions = {
    ...options,
    onProgress: (message: string) => callbacks.onProgress(step.id, message),
  };

  return Excel.run(async (context) => {
    // Snapshot before mutating steps (unless dry run). If the step has no
    // range params (structural ops like addSheet / tabColor / sheetPosition),
    // we still push an EMPTY snapshot so the handler can register an inverse
    // op on it — otherwise undo can't reverse structural changes.
    const meta = registry.getMeta(step.action);
    if (meta?.mutates && !options.dryRun) {
      const rangeAddresses = extractRangesFromParams(step);
      if (rangeAddresses.length > 0) {
        await captureSnapshotBatched(context, planId, rangeAddresses);
      } else {
        createEmptySnapshot(planId);
      }
    }

    // Execute the capability handler
    const result = await handler(context, step.params, stepOptions);
    result.durationMs = Date.now() - startTime;

    // In dry-run mode, mark as preview instead of success
    if (options.dryRun && result.status === "success") {
      result.status = "preview";
      result.message = `[Preview] ${result.message}`;
    }

    return result;
  });
}

/**
 * Rollback the most recent plan execution.
 */
export async function undoLastPlan(planId: string): Promise<boolean> {
  return Excel.run(async (context) => {
    return rollbackPlan(context, planId);
  });
}

/**
 * Extract range addresses from step params for snapshot capture.
 * Looks for common range fields in the params object.
 */
function extractRangesFromParams(step: PlanStep): string[] {
  const params = step.params as Record<string, unknown>;
  const ranges: string[] = [];

  // Check common range field names
  for (const key of [
    "range", "outputRange", "cell", "destinationRange",
    "lookupRange", "sourceRange", "dataRange", "locationRange",
    "printArea", "rangeOrName", "tableNameOrRange",
  ]) {
    if (typeof params[key] === "string") {
      ranges.push(params[key] as string);
    }
  }

  // For actions that operate on the entire sheet when no range is given
  // (e.g. findReplace without a range), snapshot the used range of the target sheet.
  if (ranges.length === 0) {
    const sheetName = params.sheetName as string | undefined;
    // Use full-sheet marker — snapshot.ts will resolve via getUsedRange
    ranges.push(sheetName ? `${sheetName}!A:XFD` : "A:XFD");
  }

  return ranges;
}

/**
 * Topological sort of plan steps respecting dependsOn.
 * Steps without dependencies come first. Falls back to original order.
 */
function topologicalSort(steps: PlanStep[]): PlanStep[] {
  const stepMap = new Map(steps.map((s) => [s.id, s]));
  const inDegree = new Map<string, number>();
  const adjacency = new Map<string, string[]>();

  for (const step of steps) {
    inDegree.set(step.id, 0);
    adjacency.set(step.id, []);
  }

  for (const step of steps) {
    if (step.dependsOn) {
      for (const dep of step.dependsOn) {
        adjacency.get(dep)?.push(step.id);
        inDegree.set(step.id, (inDegree.get(step.id) ?? 0) + 1);
      }
    }
  }

  const queue: string[] = [];
  for (const [id, deg] of inDegree) {
    if (deg === 0) queue.push(id);
  }

  const result: PlanStep[] = [];
  while (queue.length > 0) {
    const id = queue.shift()!;
    const step = stepMap.get(id);
    if (step) result.push(step);

    for (const neighbor of adjacency.get(id) ?? []) {
      const newDeg = (inDegree.get(neighbor) ?? 1) - 1;
      inDegree.set(neighbor, newDeg);
      if (newDeg === 0) queue.push(neighbor);
    }
  }

  // If the sort didn't include all steps (shouldn't happen after validation), append remainder
  if (result.length < steps.length) {
    const sorted = new Set(result.map((s) => s.id));
    for (const step of steps) {
      if (!sorted.has(step.id)) result.push(step);
    }
  }

  return result;
}

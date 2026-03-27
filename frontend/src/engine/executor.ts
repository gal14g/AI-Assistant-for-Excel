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
  StepStatus,
  PlanStep,
  ExecutionOptions,
} from "./types";
import { registry } from "./capabilityRegistry";
import { captureSnapshotBatched, rollbackPlan } from "./snapshot";
import { validatePlan } from "./validator";

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
  // Validate first
  const validation = validatePlan(plan);
  if (!validation.valid) {
    const errorMsg = validation.errors.map((e) => e.message).join("; ");
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
        callbacks.onStepComplete(step, result);
        continue;
      }
    }

    callbacks.onStepStart(step);

    const handler = registry.getHandler(step.action);
    if (!handler) {
      const result: StepResult = {
        stepId: step.id,
        status: "error",
        message: `No handler registered for action: ${step.action}`,
        error: `Unknown action: ${step.action}`,
      };
      state.stepResults.push(result);
      state.status = "failed";
      callbacks.onError(step.id, result.error!);
      callbacks.onStepComplete(step, result);
      break;
    }

    try {
      const result = await executeStep(step, handler, execOptions, callbacks, plan.planId);
      state.stepResults.push(result);
      callbacks.onStepComplete(step, result);

      if (result.status === "error") {
        state.status = "failed";
        callbacks.onError(step.id, result.error ?? "Unknown error");
        break;
      }

      completedSteps.add(step.id);
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      const result: StepResult = {
        stepId: step.id,
        status: "error",
        message: `Exception: ${errorMsg}`,
        error: errorMsg,
      };
      state.stepResults.push(result);
      state.status = "failed";
      callbacks.onError(step.id, errorMsg);
      callbacks.onStepComplete(step, result);
      break;
    }
  }

  if (state.status === "running") {
    state.status = "completed";
  } else if (state.status === "previewing") {
    state.status = "completed";
  }

  state.completedAt = new Date().toISOString();
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
    // Snapshot before mutating steps (unless dry run)
    const meta = registry.getMeta(step.action);
    if (meta?.mutates && !options.dryRun) {
      const rangeAddresses = extractRangesFromParams(step);
      if (rangeAddresses.length > 0) {
        await captureSnapshotBatched(context, planId, rangeAddresses);
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
  for (const key of ["range", "outputRange", "cell", "destinationRange"]) {
    if (typeof params[key] === "string") {
      ranges.push(params[key] as string);
    }
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

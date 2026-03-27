/**
 * usePlanExecution – Manages plan preview, execution, and rollback.
 */

import { useState, useCallback } from "react";
import { ExecutionPlan, ExecutionState, StepResult, PlanStep } from "../../engine/types";
import { executePlan, undoLastPlan, ExecutorCallbacks } from "../../engine/executor";
import { validatePlan, ValidationResult } from "../../engine/validator";

interface PlanExecutionState {
  /** Current execution state */
  executionState: ExecutionState | null;
  /** Validation result for the current plan */
  validationResult: ValidationResult | null;
  /** Whether currently executing */
  isExecuting: boolean;
  /** Whether currently in preview mode */
  isPreviewing: boolean;
  /** Step-by-step progress messages */
  progressLog: { stepId: string; message: string; timestamp: string }[];
  /** Last error */
  lastError: string | null;
}

interface PlanExecutionActions {
  /** Validate and preview a plan (dry run) */
  previewPlan: (plan: ExecutionPlan) => Promise<ExecutionState | null>;
  /** Execute a plan for real */
  runPlan: (plan: ExecutionPlan) => Promise<ExecutionState | null>;
  /** Undo the last executed plan */
  undoLast: (planId: string) => Promise<boolean>;
  /** Reset execution state */
  reset: () => void;
}

export function usePlanExecution(): PlanExecutionState & PlanExecutionActions {
  const [executionState, setExecutionState] = useState<ExecutionState | null>(null);
  const [validationResult, setValidationResult] = useState<ValidationResult | null>(null);
  const [isExecuting, setIsExecuting] = useState(false);
  const [isPreviewing, setIsPreviewing] = useState(false);
  const [progressLog, setProgressLog] = useState<
    { stepId: string; message: string; timestamp: string }[]
  >([]);
  const [lastError, setLastError] = useState<string | null>(null);

  const addProgress = useCallback((stepId: string, message: string) => {
    setProgressLog((prev) => [
      ...prev,
      { stepId, message, timestamp: new Date().toISOString() },
    ]);
  }, []);

  const createCallbacks = useCallback((): ExecutorCallbacks => {
    return {
      onStepStart: (step: PlanStep) => {
        addProgress(step.id, `Starting: ${step.description}`);
        setExecutionState((prev) =>
          prev
            ? {
                ...prev,
                stepResults: [
                  ...prev.stepResults,
                  { stepId: step.id, status: "running", message: "Running..." },
                ],
              }
            : prev
        );
      },
      onStepComplete: (step: PlanStep, result: StepResult) => {
        addProgress(step.id, `${result.status}: ${result.message}`);
        setExecutionState((prev) => {
          if (!prev) return prev;
          const results = prev.stepResults.map((r) =>
            r.stepId === step.id ? result : r
          );
          return { ...prev, stepResults: results };
        });
      },
      onProgress: (stepId: string, message: string) => {
        addProgress(stepId, message);
      },
      onPlanComplete: (state: ExecutionState) => {
        setExecutionState(state);
      },
      onError: (stepId: string, error: string) => {
        addProgress(stepId, `ERROR: ${error}`);
        setLastError(error);
      },
    };
  }, [addProgress]);

  const previewPlan = useCallback(
    async (plan: ExecutionPlan): Promise<ExecutionState | null> => {
      setIsPreviewing(true);
      setLastError(null);
      setProgressLog([]);

      // Validate first
      const validation = validatePlan(plan);
      setValidationResult(validation);

      if (!validation.valid) {
        setIsPreviewing(false);
        setLastError(
          `Validation failed: ${validation.errors.map((e) => e.message).join("; ")}`
        );
        return null;
      }

      try {
        const state = await executePlan(plan, createCallbacks(), { dryRun: true });
        setIsPreviewing(false);
        return state;
      } catch (err) {
        setLastError(err instanceof Error ? err.message : "Preview failed");
        setIsPreviewing(false);
        return null;
      }
    },
    [createCallbacks]
  );

  const runPlan = useCallback(
    async (plan: ExecutionPlan): Promise<ExecutionState | null> => {
      setIsExecuting(true);
      setLastError(null);
      setProgressLog([]);

      // Validate
      const validation = validatePlan(plan);
      setValidationResult(validation);

      if (!validation.valid) {
        setIsExecuting(false);
        setLastError(
          `Validation failed: ${validation.errors.map((e) => e.message).join("; ")}`
        );
        return null;
      }

      try {
        const state = await executePlan(plan, createCallbacks());
        setIsExecuting(false);
        return state;
      } catch (err) {
        setLastError(err instanceof Error ? err.message : "Execution failed");
        setIsExecuting(false);
        return null;
      }
    },
    [createCallbacks]
  );

  const undoLast = useCallback(async (planId: string): Promise<boolean> => {
    try {
      const success = await undoLastPlan(planId);
      if (success) {
        addProgress("undo", "Successfully rolled back last execution");
        setExecutionState((prev) =>
          prev ? { ...prev, status: "rolledBack" } : prev
        );
      }
      return success;
    } catch (err) {
      setLastError(err instanceof Error ? err.message : "Undo failed");
      return false;
    }
  }, [addProgress]);

  const reset = useCallback(() => {
    setExecutionState(null);
    setValidationResult(null);
    setIsExecuting(false);
    setIsPreviewing(false);
    setProgressLog([]);
    setLastError(null);
  }, []);

  return {
    executionState,
    validationResult,
    isExecuting,
    isPreviewing,
    progressLog,
    lastError,
    previewPlan,
    runPlan,
    undoLast,
    reset,
  };
}

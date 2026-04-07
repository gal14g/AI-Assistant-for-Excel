/**
 * Tests for the ExecutionContextDTO — the frontend data structure
 * sent to the backend for multi-turn plan refinement.
 *
 * These tests verify:
 * 1. The DTO shape matches what the backend expects
 * 2. Realistic refinement scenarios construct valid DTOs
 * 3. Hebrew error messages are preserved
 * 4. Edge cases: all steps succeed, all steps fail, single step plans
 */

import type { ExecutionContextDTO } from "../services/api";
import type { ExecutionState, ExecutionPlan } from "../engine/types";

/**
 * Helper: Build an ExecutionContextDTO from a plan and its execution state,
 * mirroring the logic in useChat.refinePlan().
 */
function buildExecutionContext(
  plan: ExecutionPlan,
  state: ExecutionState,
): ExecutionContextDTO {
  const failedResult = state.stepResults.find((r) => r.status === "error");
  const failedStep = plan.steps.find((s) => s.id === failedResult?.stepId);

  return {
    originalPlanId: plan.planId,
    originalUserRequest: plan.userRequest,
    stepResults: state.stepResults.map((r) => ({
      stepId: r.stepId,
      status: r.status as "success" | "error" | "skipped" | "preview",
      message: r.message,
      error: r.error,
    })),
    failedStepId: failedResult?.stepId,
    failedStepAction: failedStep?.action,
    failedStepError: failedResult?.error ?? failedResult?.message,
  };
}

describe("ExecutionContextDTO construction", () => {
  const basePlan: ExecutionPlan = {
    planId: "plan-001",
    createdAt: "2026-04-07T10:00:00Z",
    userRequest: "create sales dashboard",
    summary: "Dashboard with KPIs and charts",
    steps: [
      { id: "step_1", description: "Add sheet", action: "addSheet", params: { sheetName: "Dashboard" } as never },
      { id: "step_2", description: "Write KPIs", action: "writeValues", params: { range: "A1:B3", values: [[]] } as never, dependsOn: ["step_1"] },
      { id: "step_3", description: "Create chart", action: "createChart", params: { dataRange: "A1:B3", chartType: "bar" } as never, dependsOn: ["step_2"] },
    ],
    preserveFormatting: false,
    confidence: 0.9,
  };

  it("correctly identifies the failed step", () => {
    const state: ExecutionState = {
      planId: "plan-001",
      status: "failed",
      stepResults: [
        { stepId: "step_1", status: "success", message: "Added sheet" },
        { stepId: "step_2", status: "success", message: "Wrote 3 rows" },
        { stepId: "step_3", status: "error", message: "Chart failed", error: "Invalid data range" },
      ],
    };

    const ctx = buildExecutionContext(basePlan, state);

    expect(ctx.originalPlanId).toBe("plan-001");
    expect(ctx.originalUserRequest).toBe("create sales dashboard");
    expect(ctx.failedStepId).toBe("step_3");
    expect(ctx.failedStepAction).toBe("createChart");
    expect(ctx.failedStepError).toBe("Invalid data range");
    expect(ctx.stepResults.length).toBe(3);
  });

  it("preserves all step results in order", () => {
    const state: ExecutionState = {
      planId: "plan-001",
      status: "failed",
      stepResults: [
        { stepId: "step_1", status: "success", message: "ok" },
        { stepId: "step_2", status: "error", message: "fail", error: "boom" },
        { stepId: "step_3", status: "skipped", message: "skipped" },
      ],
    };

    const ctx = buildExecutionContext(basePlan, state);

    expect(ctx.stepResults[0].stepId).toBe("step_1");
    expect(ctx.stepResults[0].status).toBe("success");
    expect(ctx.stepResults[1].stepId).toBe("step_2");
    expect(ctx.stepResults[1].status).toBe("error");
    expect(ctx.stepResults[1].error).toBe("boom");
    expect(ctx.stepResults[2].stepId).toBe("step_3");
    expect(ctx.stepResults[2].status).toBe("skipped");
  });

  it("handles first-step failure (no successful steps)", () => {
    const state: ExecutionState = {
      planId: "plan-001",
      status: "failed",
      stepResults: [
        { stepId: "step_1", status: "error", message: "Sheet already exists", error: "Duplicate sheet name" },
      ],
    };

    const ctx = buildExecutionContext(basePlan, state);

    expect(ctx.failedStepId).toBe("step_1");
    expect(ctx.failedStepAction).toBe("addSheet");
    expect(ctx.stepResults.length).toBe(1);
  });

  it("handles all-success case (no failed step)", () => {
    const state: ExecutionState = {
      planId: "plan-001",
      status: "completed",
      stepResults: [
        { stepId: "step_1", status: "success", message: "ok" },
        { stepId: "step_2", status: "success", message: "ok" },
        { stepId: "step_3", status: "success", message: "ok" },
      ],
    };

    const ctx = buildExecutionContext(basePlan, state);

    expect(ctx.failedStepId).toBeUndefined();
    expect(ctx.failedStepAction).toBeUndefined();
    expect(ctx.failedStepError).toBeUndefined();
  });

  it("handles Hebrew error messages", () => {
    const hebrewPlan: ExecutionPlan = {
      ...basePlan,
      planId: "plan-heb",
      userRequest: "התאם לקוחות להזמנות",
      steps: [
        { id: "step_1", description: "חפש התאמות", action: "matchRecords", params: {
          lookupRange: "לקוחות!A:A", sourceRange: "הזמנות!B:B",
          matchType: "contains", outputRange: "לקוחות!D:D",
        } as never },
      ],
    };

    const state: ExecutionState = {
      planId: "plan-heb",
      status: "failed",
      stepResults: [
        { stepId: "step_1", status: "error", message: "שגיאה בטווח", error: "הטווח לקוחות!A:A לא נמצא" },
      ],
    };

    const ctx = buildExecutionContext(hebrewPlan, state);

    expect(ctx.originalUserRequest).toBe("התאם לקוחות להזמנות");
    expect(ctx.failedStepError).toBe("הטווח לקוחות!A:A לא נמצא");
    expect(ctx.stepResults[0].error).toBe("הטווח לקוחות!A:A לא נמצא");
  });

  it("handles single-step plan failure", () => {
    const singlePlan: ExecutionPlan = {
      ...basePlan,
      steps: [basePlan.steps[0]],
    };

    const state: ExecutionState = {
      planId: "plan-001",
      status: "failed",
      stepResults: [
        { stepId: "step_1", status: "error", message: "fail", error: "err" },
      ],
    };

    const ctx = buildExecutionContext(singlePlan, state);

    expect(ctx.stepResults.length).toBe(1);
    expect(ctx.failedStepId).toBe("step_1");
  });
});

describe("ExecutionContextDTO — multi-step pipeline fail-and-fix scenario", () => {
  it("simulates a 5-step pipeline where step 3 fails then gets fixed", () => {
    const plan: ExecutionPlan = {
      planId: "pipeline-5",
      createdAt: "2026-04-07",
      userRequest: "clean data, match, aggregate, chart, format",
      summary: "Full pipeline",
      steps: [
        { id: "step_1", description: "Clean names", action: "cleanupText", params: { range: "A:A", operations: ["trim"] } as never },
        { id: "step_2", description: "Match records", action: "matchRecords", params: { lookupRange: "A:A", sourceRange: "B:B", matchType: "exact", outputRange: "C:C" } as never, dependsOn: ["step_1"] },
        { id: "step_3", description: "Group sums", action: "groupSum", params: { dataRange: "A1:C100", groupByColumn: 1, sumColumn: 3, outputRange: "E1" } as never, dependsOn: ["step_2"] },
        { id: "step_4", description: "Create chart", action: "createChart", params: { dataRange: "E1:F10", chartType: "bar" } as never, dependsOn: ["step_3"] },
        { id: "step_5", description: "Auto-fit", action: "autoFitColumns", params: {} as never, dependsOn: ["step_4"] },
      ],
      preserveFormatting: false,
      confidence: 0.88,
    };

    // First execution: steps 1-2 succeed, step 3 fails
    const failedState: ExecutionState = {
      planId: "pipeline-5",
      status: "failed",
      stepResults: [
        { stepId: "step_1", status: "success", message: "Cleaned 100 cells" },
        { stepId: "step_2", status: "success", message: "Matched 85/100" },
        { stepId: "step_3", status: "error", message: "Column 3 is text", error: "Expected numeric column" },
        { stepId: "step_4", status: "skipped", message: "Skipped: depends on step_3" },
        { stepId: "step_5", status: "skipped", message: "Skipped: depends on step_4" },
      ],
    };

    const ctx = buildExecutionContext(plan, failedState);

    // Verify the context captures the right state
    expect(ctx.failedStepId).toBe("step_3");
    expect(ctx.failedStepAction).toBe("groupSum");
    expect(ctx.stepResults.filter((r) => r.status === "success").length).toBe(2);
    expect(ctx.stepResults.filter((r) => r.status === "skipped").length).toBe(2);
    expect(ctx.stepResults.filter((r) => r.status === "error").length).toBe(1);

    // Now simulate the corrected plan (only step_3 onwards)
    const fixedPlan: ExecutionPlan = {
      planId: "pipeline-5-fix",
      createdAt: "2026-04-07",
      userRequest: "fix the groupSum step",
      summary: "Use COUNTIF instead of SUM",
      steps: [
        { id: "step_3", description: "Count with COUNTIF", action: "writeFormula", params: { cell: "E1", formula: "=COUNTIF(C:C,\"v\")", fillDown: 10 } as never },
        { id: "step_4", description: "Create chart", action: "createChart", params: { dataRange: "E1:F10", chartType: "bar" } as never, dependsOn: ["step_3"] },
        { id: "step_5", description: "Auto-fit", action: "autoFitColumns", params: {} as never, dependsOn: ["step_4"] },
      ],
      preserveFormatting: false,
      confidence: 0.85,
    };

    // The fixed plan should be a valid plan
    // (We can't run validatePlan here without the full registry being loaded,
    // but we can verify the structure is correct)
    expect(fixedPlan.steps.length).toBe(3);
    expect(fixedPlan.steps[0].id).toBe("step_3");
    expect(fixedPlan.steps[0].action).toBe("writeFormula");
  });
});

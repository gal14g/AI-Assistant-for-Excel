/**
 * Tests for StepResult.outputs — the new field for step output binding.
 *
 * Verifies that StepResult correctly carries output metadata
 * and that downstream consumers can access the fields.
 */

import type { StepResult } from "../engine/types";

describe("StepResult.outputs", () => {
  it("supports string output values", () => {
    const result: StepResult = {
      stepId: "step_1",
      status: "success",
      message: "Added sheet",
      outputs: { sheetName: "Dashboard" },
    };
    expect(result.outputs?.sheetName).toBe("Dashboard");
  });

  it("supports numeric output values", () => {
    const result: StepResult = {
      stepId: "step_2",
      status: "success",
      message: "Wrote 42 rows",
      outputs: { rowsWritten: 42 },
    };
    expect(result.outputs?.rowsWritten).toBe(42);
  });

  it("supports boolean output values", () => {
    const result: StepResult = {
      stepId: "step_1",
      status: "success",
      message: "Created",
      outputs: { createdNew: true },
    };
    expect(result.outputs?.createdNew).toBe(true);
  });

  it("supports multiple output fields", () => {
    const result: StepResult = {
      stepId: "step_1",
      status: "success",
      message: "ok",
      outputs: {
        sheetName: "Data",
        outputRange: "Data!A1:C50",
        rowsWritten: 50,
      },
    };
    expect(result.outputs?.sheetName).toBe("Data");
    expect(result.outputs?.outputRange).toBe("Data!A1:C50");
    expect(result.outputs?.rowsWritten).toBe(50);
  });

  it("outputs field is optional (backward compatible)", () => {
    const result: StepResult = {
      stepId: "step_1",
      status: "success",
      message: "ok",
    };
    expect(result.outputs).toBeUndefined();
  });

  it("handles Hebrew values in outputs", () => {
    const result: StepResult = {
      stepId: "step_1",
      status: "success",
      message: "ok",
      outputs: { sheetName: "נתונים", tableName: "טבלה_מכירות" },
    };
    expect(result.outputs?.sheetName).toBe("נתונים");
    expect(result.outputs?.tableName).toBe("טבלה_מכירות");
  });

  it("error results can also have outputs (partial success)", () => {
    const result: StepResult = {
      stepId: "step_1",
      status: "error",
      message: "Partial failure",
      error: "Chart creation failed",
      outputs: { outputRange: "A1:B10" },
    };
    expect(result.status).toBe("error");
    expect(result.outputs?.outputRange).toBe("A1:B10");
  });
});

describe("ExecutionState with outputs", () => {
  it("stepResults array preserves outputs from each step", () => {
    const results: StepResult[] = [
      {
        stepId: "step_1",
        status: "success",
        message: "Added sheet",
        outputs: { sheetName: "Report" },
      },
      {
        stepId: "step_2",
        status: "success",
        message: "Wrote data",
        outputs: { outputRange: "Report!A1:C20", rowsWritten: 20 },
      },
      {
        stepId: "step_3",
        status: "error",
        message: "Chart failed",
        error: "Invalid range",
      },
    ];

    // Simulate what executor does: build a map for downstream binding
    const resultsMap = new Map(results.map((r) => [r.stepId, r]));

    expect(resultsMap.get("step_1")?.outputs?.sheetName).toBe("Report");
    expect(resultsMap.get("step_2")?.outputs?.outputRange).toBe("Report!A1:C20");
    expect(resultsMap.get("step_3")?.outputs).toBeUndefined();
  });
});

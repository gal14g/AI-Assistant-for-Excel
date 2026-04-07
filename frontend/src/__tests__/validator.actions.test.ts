/**
 * Tests that all 76 StepAction values are handled in the frontend validator.
 *
 * Also validates multi-step pipeline plans with dependencies, binding tokens,
 * Hebrew sheet names, and realistic data scenarios.
 */

// Import capability index FIRST so all 76 handlers get registered before
// the validator checks registry.has(action).
import "../engine/capabilities/index";
import { validatePlan } from "../engine/validator";
import type { ExecutionPlan, PlanStep } from "../engine/types";

function makePlan(steps: PlanStep[], overrides?: Partial<ExecutionPlan>): ExecutionPlan {
  return {
    planId: "test-plan",
    createdAt: "2026-04-07T00:00:00Z",
    userRequest: "test",
    summary: "test plan",
    steps,
    preserveFormatting: true,
    confidence: 0.9,
    ...overrides,
  };
}

function makeStep(overrides: Partial<PlanStep> & { id: string; action: PlanStep["action"] | string }): PlanStep {
  return {
    description: `Step ${overrides.id}`,
    params: {},
    ...overrides,
  } as PlanStep;
}

describe("validatePlan — all 76 actions", () => {
  // The complete list of 76 actions
  const ALL_ACTIONS = [
    "readRange", "writeValues", "writeFormula", "matchRecords", "groupSum",
    "createTable", "applyFilter", "sortRange", "createPivot", "createChart",
    "addConditionalFormat", "cleanupText", "removeDuplicates", "freezePanes",
    "findReplace", "addValidation", "addSheet", "renameSheet", "deleteSheet",
    "copySheet", "protectSheet", "autoFitColumns", "mergeCells", "setNumberFormat",
    "insertDeleteRows", "addSparkline", "formatCells", "clearRange", "hideShow",
    "addComment", "addHyperlink", "groupRows", "setRowColSize", "copyPasteRange",
    "pageLayout", "insertPicture", "insertShape", "insertTextBox", "addSlicer",
    "splitColumn", "unpivot", "crossTabulate", "bulkFormula", "compareSheets",
    "consolidateRanges", "extractPattern", "categorize", "fillBlanks", "subtotals",
    "transpose", "namedRange",
    // New batch 2
    "fuzzyMatch", "deleteRowsByCondition", "splitByGroup", "lookupAll",
    "regexReplace", "coerceDataType", "normalizeDates", "deduplicateAdvanced",
    "joinSheets", "frequencyDistribution", "runningTotal", "rankColumn",
    "topN", "percentOfTotal", "growthRate", "consolidateAllSheets",
    "cloneSheetStructure", "addReportHeader", "alternatingRowFormat",
    "quickFormat", "refreshPivot", "pivotCalculatedField", "addDropdownControl",
    "conditionalFormula", "spillFormula",
  ];

  it("should have exactly 76 known actions", () => {
    expect(ALL_ACTIONS.length).toBe(76);
  });

  it("should accept a valid plan for each of the 76 actions", () => {
    // For each action, create a minimal plan with dummy params and ensure
    // the validator doesn't reject with UNKNOWN_ACTION.
    for (const action of ALL_ACTIONS) {
      const plan = makePlan([
        makeStep({ id: "step_1", action: action as PlanStep["action"], params: { range: "A1:B10" } as never }),
      ]);
      const result = validatePlan(plan);
      const unknownErrors = result.errors.filter((e) => e.code === "UNKNOWN_ACTION");
      expect(unknownErrors).toEqual([]);
    }
  });
});

describe("validatePlan — multi-step pipeline validation", () => {
  it("accepts a 5-step dashboard pipeline with dependencies", () => {
    const plan = makePlan([
      makeStep({ id: "step_1", action: "addSheet", params: { sheetName: "Dashboard" } as never }),
      makeStep({
        id: "step_2", action: "writeValues",
        params: { range: "Dashboard!A1:B3", values: [["Metric", "Value"], ["Total", ""], ["Avg", ""]] } as never,
        dependsOn: ["step_1"],
      }),
      makeStep({
        id: "step_3", action: "writeFormula",
        params: { cell: "Dashboard!B2", formula: "=SUM(Data!C:C)" } as never,
        dependsOn: ["step_2"],
      }),
      makeStep({
        id: "step_4", action: "createChart",
        params: { dataRange: "Dashboard!A1:B3", chartType: "bar" } as never,
        dependsOn: ["step_3"],
      }),
      makeStep({
        id: "step_5", action: "autoFitColumns",
        params: { sheetName: "Dashboard" } as never,
        dependsOn: ["step_4"],
      }),
    ], { preserveFormatting: false });
    const result = validatePlan(plan);
    expect(result.valid).toBe(true);
  });

  it("rejects a plan with a dependency cycle", () => {
    const plan = makePlan([
      makeStep({ id: "step_1", action: "readRange", params: { range: "A1:B10" } as never, dependsOn: ["step_2"] }),
      makeStep({ id: "step_2", action: "readRange", params: { range: "C1:D10" } as never, dependsOn: ["step_1"] }),
    ]);
    const result = validatePlan(plan);
    expect(result.valid).toBe(false);
    expect(result.errors.some((e) => e.code === "DEPENDENCY_CYCLE")).toBe(true);
  });

  it("rejects a plan with a missing dependency", () => {
    const plan = makePlan([
      makeStep({ id: "step_1", action: "readRange", params: { range: "A1:B10" } as never, dependsOn: ["nonexistent"] }),
    ]);
    const result = validatePlan(plan);
    expect(result.valid).toBe(false);
    expect(result.errors.some((e) => e.code === "INVALID_DEPENDENCY")).toBe(true);
  });

  it("rejects duplicate step IDs", () => {
    const plan = makePlan([
      makeStep({ id: "step_1", action: "readRange", params: { range: "A1:B10" } as never }),
      makeStep({ id: "step_1", action: "readRange", params: { range: "C1:D10" } as never }),
    ]);
    const result = validatePlan(plan);
    expect(result.valid).toBe(false);
    expect(result.errors.some((e) => e.code === "DUPLICATE_STEP_ID")).toBe(true);
  });
});

describe("validatePlan — binding tokens are passthrough", () => {
  it("accepts params containing {{step_N.field}} binding tokens", () => {
    const plan = makePlan([
      makeStep({ id: "step_1", action: "addSheet", params: { sheetName: "Report" } as never }),
      makeStep({
        id: "step_2", action: "writeValues",
        params: { range: "Report!A1:B5", values: [["A", "B"]] } as never,
        dependsOn: ["step_1"],
      }),
      makeStep({
        id: "step_3", action: "createChart",
        params: { dataRange: "{{step_2.outputRange}}", chartType: "line" } as never,
        dependsOn: ["step_2"],
      }),
    ], { preserveFormatting: false });
    const result = validatePlan(plan);
    expect(result.valid).toBe(true);
  });
});

describe("validatePlan — Hebrew sheet names in plans", () => {
  it("accepts plans with Hebrew range addresses", () => {
    const plan = makePlan([
      makeStep({
        id: "step_1", action: "matchRecords",
        params: {
          lookupRange: "נתונים!A:A",
          sourceRange: "הזמנות!B:B",
          matchType: "contains",
          outputRange: "נתונים!D:D",
          writeValue: "v",
        } as never,
      }),
    ]);
    const result = validatePlan(plan);
    // Should not have UNKNOWN_ACTION or structural errors
    const structuralErrors = result.errors.filter((e) =>
      ["UNKNOWN_ACTION", "DUPLICATE_STEP_ID", "DEPENDENCY_CYCLE", "INVALID_DEPENDENCY"].includes(e.code),
    );
    expect(structuralErrors).toEqual([]);
  });

  it("accepts Hebrew sheet name in addSheet", () => {
    const plan = makePlan([
      makeStep({ id: "step_1", action: "addSheet", params: { sheetName: "דוח מכירות" } as never }),
    ]);
    const result = validatePlan(plan);
    expect(result.valid).toBe(true);
  });
});

describe("validatePlan — new action param validation", () => {
  it("validates normalizeDates params (range + outputFormat required)", () => {
    const plan = makePlan([
      makeStep({
        id: "step_1", action: "normalizeDates",
        params: { range: "A:A", outputFormat: "dd/mm/yyyy" } as never,
      }),
    ]);
    const result = validatePlan(plan);
    // Should not have MISSING_REQUIRED errors for these params
    const missingErrors = result.errors.filter(
      (e) => e.code === "MISSING_REQUIRED" && e.stepId === "step_1",
    );
    expect(missingErrors).toEqual([]);
  });

  it("validates frequencyDistribution params", () => {
    const plan = makePlan([
      makeStep({
        id: "step_1", action: "frequencyDistribution",
        params: { sourceRange: "A:A", outputRange: "C1" } as never,
      }),
    ]);
    const result = validatePlan(plan);
    const missingErrors = result.errors.filter(
      (e) => e.code === "MISSING_REQUIRED" && e.stepId === "step_1",
    );
    expect(missingErrors).toEqual([]);
  });

  it("validates deleteRowsByCondition params", () => {
    const plan = makePlan([
      makeStep({
        id: "step_1", action: "deleteRowsByCondition",
        params: { range: "A1:D100", column: 2, condition: "blank" } as never,
      }),
    ]);
    const result = validatePlan(plan);
    const missingErrors = result.errors.filter(
      (e) => e.code === "MISSING_REQUIRED" && e.stepId === "step_1",
    );
    expect(missingErrors).toEqual([]);
  });

  it("validates joinSheets params", () => {
    const plan = makePlan([
      makeStep({
        id: "step_1", action: "joinSheets",
        params: {
          leftRange: "Sheet1!A1:C50",
          rightRange: "Sheet2!A1:B50",
          leftKeyColumn: 1,
          rightKeyColumn: 1,
          joinType: "inner",
          outputRange: "Sheet3!A1",
        } as never,
      }),
    ]);
    const result = validatePlan(plan);
    const missingErrors = result.errors.filter(
      (e) => e.code === "MISSING_REQUIRED" && e.stepId === "step_1",
    );
    expect(missingErrors).toEqual([]);
  });
});

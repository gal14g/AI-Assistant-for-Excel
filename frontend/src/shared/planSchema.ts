/**
 * Zod schema for ExecutionPlan validation.
 *
 * This provides runtime validation of plans received from the backend.
 * It's a complementary layer to the TypeScript types (which are compile-time only).
 */

import { z } from "zod";

export const StepActionSchema = z.enum([
  "readRange",
  "writeValues",
  "writeFormula",
  "matchRecords",
  "groupSum",
  "createTable",
  "applyFilter",
  "sortRange",
  "createPivot",
  "createChart",
  "addConditionalFormat",
  "cleanupText",
  "removeDuplicates",
  "freezePanes",
  "findReplace",
  "addValidation",
  "addSheet",
  "renameSheet",
  "deleteSheet",
  "copySheet",
  "protectSheet",
  "autoFitColumns",
  "mergeCells",
  "setNumberFormat",
  "insertDeleteRows",
  "addSparkline",
  "formatCells",
  "clearRange",
  "hideShow",
  "addComment",
  "addHyperlink",
  "groupRows",
  "setRowColSize",
  "copyPasteRange",
  "pageLayout",
  "insertPicture",
  "insertShape",
  "insertTextBox",
  "addSlicer",
  "splitColumn",
  "unpivot",
  "crossTabulate",
  "bulkFormula",
  "compareSheets",
  "consolidateRanges",
  "extractPattern",
  "categorize",
  "fillBlanks",
  "subtotals",
  "transpose",
  "namedRange",
]);

export const PlanStepSchema = z.object({
  id: z.string().min(1),
  description: z.string().min(1),
  action: StepActionSchema,
  params: z.record(z.unknown()),
  dependsOn: z.array(z.string()).optional(),
});

export const ExecutionPlanSchema = z.object({
  planId: z.string().min(1),
  createdAt: z.string(),
  userRequest: z.string(),
  summary: z.string(),
  steps: z.array(PlanStepSchema).min(1),
  preserveFormatting: z.boolean().default(true),
  confidence: z.number().min(0).max(1),
  warnings: z.array(z.string()).optional(),
});

/**
 * Validate a raw JSON object against the plan schema.
 * Returns the typed plan or throws a ZodError.
 */
export function parsePlan(raw: unknown) {
  return ExecutionPlanSchema.parse(raw);
}

/**
 * Safe parse that returns a result object instead of throwing.
 */
export function safeParsePlan(raw: unknown) {
  return ExecutionPlanSchema.safeParse(raw);
}

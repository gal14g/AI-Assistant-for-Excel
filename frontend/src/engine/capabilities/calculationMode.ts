/**
 * calculationMode — set workbook calculation mode.
 *
 * Office.js: workbook.application.calculationMode (ExcelApi 1.8+).
 * Values: "Automatic" | "AutomaticExceptTables" | "Manual".
 *
 * Useful as a bracket around large bulk-writes to avoid recalc thrashing.
 */

import { CapabilityMeta, CalculationModeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "calculationMode",
  description: "Set workbook calculation mode (manual / automatic / automaticExceptTables)",
  mutates: false,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.8",
};

async function handler(
  context: Excel.RequestContext,
  params: CalculationModeParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { mode } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would set calculation mode to ${mode}` };
  }

  const modeMap: Record<string, Excel.CalculationMode> = {
    manual:                 Excel.CalculationMode.manual,
    automatic:              Excel.CalculationMode.automatic,
    automaticExceptTables:  Excel.CalculationMode.automaticExceptTables,
  };
  const target = modeMap[mode];
  if (!target) {
    return { stepId: "", status: "error", message: `Unknown calculation mode: ${mode}` };
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (context.workbook as any).application.calculationMode = target;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Calculation mode set to ${mode}.`,
  };
}

// Fallback: pre-1.8 has no programmatic calc mode
async function fallback(
  _context: Excel.RequestContext,
  params: CalculationModeParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: "Would skip calc mode (legacy fallback)." };
  }
  return {
    stepId: "",
    status: "success",
    message:
      `Calculation mode request (${params.mode}) skipped — Application.calculationMode requires ExcelApi 1.8+. ` +
      `Set manually via Formulas > Calculation Options (legacy-Excel fallback).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };

/**
 * addSparkline – Add sparkline mini-charts to a range of cells.
 *
 * Office.js notes:
 * - Sparklines require ExcelApi 1.9+.
 * - The correct API is worksheet.sparklineGroups.add(type, sourceData, locationRange)
 *   NOT range.group() — that method is for row/column outline grouping.
 * - SparklineType enum is not exported in all @types/office-js versions,
 *   so we pass the string values ("Line", "Column", "WinLoss") via cast.
 *
 * Typical use-case:
 *   dataRange:     "Sheet1!B2:M10"   (12 months × 9 products)
 *   locationRange: "Sheet1!N2:N10"   (one sparkline per product row in col N)
 */

import { CapabilityMeta, AddSparklineParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addSparkline",
  description: "Add sparkline mini-charts to cells",
  mutates: false,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.9",
};

async function handler(
  context: Excel.RequestContext,
  params: AddSparklineParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { dataRange, locationRange, sparklineType = "line" } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add ${sparklineType} sparklines to ${locationRange} from ${dataRange}`,
    };
  }

  options.onProgress?.(`Adding ${sparklineType} sparklines...`);

  const locationRng = resolveRange(context, locationRange);
  const dataRng    = resolveRange(context, dataRange);

  // Excel API string values for SparklineType (ExcelApi 1.9+).
  // We use strings instead of the Excel.SparklineType enum because the enum
  // is not exported in all versions of @types/office-js.
  const typeMap: Record<string, string> = {
    line:    "Line",
    column:  "Column",
    winLoss: "WinLoss",
  };
  const excelType = typeMap[sparklineType] ?? "Line";

  // sparklineGroups.add(type, sourceData, locationRange) — ExcelApi 1.9
  // Cast worksheet to any because SparklineGroupCollection typing may be absent.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (locationRng.worksheet as any).sparklineGroups.add(excelType, dataRng, locationRng);

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added ${sparklineType} sparklines to ${locationRange} from ${dataRange}`,
  };
}

registry.register(meta, handler as any);
export { meta };

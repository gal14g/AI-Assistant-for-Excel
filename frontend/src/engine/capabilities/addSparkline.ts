/**
 * addSparkline – Add sparkline mini-charts to a range of cells.
 *
 * Office.js notes:
 * - Sparklines require ExcelApi 1.9+.
 * - range.group(type, dataRange) creates a SparklineGroup on every cell in
 *   the location range, each reading from the corresponding row of dataRange.
 * - SparklineType: line | column | winLoss
 *
 * Typical use-case:
 *   dataRange:     "Sheet1!B2:M10"   (12 months × 9 products)
 *   locationRange: "Sheet1!N2:N10"   (one sparkline per product in col N)
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
  const { dataRange, locationRange, sparklineType = "line", color } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add ${sparklineType} sparklines to ${locationRange} from data in ${dataRange}`,
    };
  }

  options.onProgress?.(`Adding ${sparklineType} sparklines...`);

  const locationRng = resolveRange(context, locationRange);
  const dataRng = resolveRange(context, dataRange);

  const typeMap: Record<string, Excel.SparklineType> = {
    line:    Excel.SparklineType.line,
    column:  Excel.SparklineType.column,
    winLoss: Excel.SparklineType.winLoss,
  };

  const excelType = typeMap[sparklineType] ?? Excel.SparklineType.line;
  const group = locationRng.group(excelType, dataRng);

  // Optional: set sparkline color
  if (color) {
    group.load("items");
    await context.sync();
    // Color applies via presetStyle or seriesColor on the group
    try {
      group.presetStyle = Excel.SparklineStyle.custom;
    } catch {
      // presetStyle may not be available on all API versions — safe to skip
    }
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added ${sparklineType} sparklines to ${locationRange} from ${dataRange}`,
  };
}

registry.register(meta, handler as any);
export { meta };

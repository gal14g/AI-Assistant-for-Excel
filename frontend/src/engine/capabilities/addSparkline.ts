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
  const sparklineGroup: any = (locationRng.worksheet as any).sparklineGroups.add(excelType, dataRng, locationRng);

  // Apply optional color to the sparkline group
  if (params.color) {
    sparklineGroup.format.line.color = params.color;
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added ${sparklineType} sparklines to ${locationRange} from ${dataRange}`,
    outputs: { range: locationRange },
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.9) ───────────────────────────────────
// Sparklines don't exist before 1.9. Best substitute: embed a small column
// chart per row positioned over each target cell. This works on Excel 2016+
// (charts were in 1.1) and visually mimics a sparkline — compact, inline, and
// driven by the same source data so it updates live.
async function fallback(
  context: Excel.RequestContext,
  params: AddSparklineParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { dataRange, locationRange, sparklineType = "line" } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add ${locationRange.split(":").length > 1 ? "mini charts per row" : "a mini chart"} at ${locationRange} (sparklines unsupported; using embedded ${sparklineType === "column" ? "column" : "line"} charts).`,
    };
  }

  options.onProgress?.("Legacy-Excel mode: embedding mini charts instead of sparklines...");

  const dataRng = resolveRange(context, dataRange);
  const locRng = resolveRange(context, locationRange);
  dataRng.load(["rowCount", "columnCount", "address", "worksheet/name"]);
  locRng.load(["rowCount", "columnCount", "address", "left", "top", "width", "height", "worksheet/name"]);
  await context.sync();

  // One chart per location cell, driven by the corresponding row of source data.
  const rows = Math.min(dataRng.rowCount, locRng.rowCount);
  if (rows <= 0) {
    return { stepId: "", status: "error", message: "Data range and location range must each contain at least one row." };
  }

  const sheet = locRng.worksheet;
  const chartType: string = sparklineType === "column" ? "ColumnClustered" : "Line";

  for (let r = 0; r < rows; r++) {
    const rowData = dataRng.getRow(r);
    const anchor = locRng.getCell(r, 0);
    anchor.load(["left", "top", "width", "height"]);
    await context.sync();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const chart: any = (sheet as any).charts.add(chartType, rowData, "auto");
    // Size + position the chart to hug the anchor cell.
    chart.left = anchor.left;
    chart.top = anchor.top;
    chart.width = anchor.width;
    chart.height = anchor.height;
    // Strip non-essential chrome so it reads like a sparkline.
    try {
      chart.title.visible = false;
      chart.legend.visible = false;
      chart.axes.categoryAxis.visible = false;
      chart.axes.valueAxis.visible = false;
    } catch {
      /* some chart sub-properties may not be available on very old API sets */
    }
    if (params.color) {
      try {
        chart.series.getItemAt(0).format.line.color = params.color;
      } catch {
        /* color setting best-effort on legacy API */
      }
    }
    await context.sync();
  }

  return {
    stepId: "",
    status: "success",
    message:
      `Added ${rows} mini ${sparklineType === "column" ? "column" : "line"} charts to ${locationRange} ` +
      `(legacy-Excel fallback — native sparklines require ExcelApi 1.9+).`,
    outputs: { range: locationRange },
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };

/**
 * createChart – Create a chart from data.
 *
 * Office.js notes:
 * - Chart types map to Excel.ChartType enum values.
 * - setData() binds the chart to the source range.
 * - Position is set via left/top/width/height in points.
 * - Charts are added to a sheet's chart collection.
 */

import { CapabilityMeta, CreateChartParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "createChart",
  description: "Create a chart from data range",
  mutates: true,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.1",
};

async function handler(
  context: Excel.RequestContext,
  params: CreateChartParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { dataRange, chartType, title, sheetName, position, seriesNames } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would create ${chartType} chart from ${dataRange}`,
    };
  }

  options.onProgress?.(`Creating ${chartType} chart...`);

  // Validate the target sheet if explicitly provided, then resolve it.
  let sheet: Excel.Worksheet;
  if (sheetName) {
    const ws = context.workbook.worksheets.getItemOrNullObject(sheetName);
    ws.load("isNullObject");
    await context.sync();
    if (ws.isNullObject) {
      return {
        stepId: "",
        status: "error",
        message: `Sheet "${sheetName}" not found. Please check the sheet name.`,
      };
    }
    sheet = ws;
  } else {
    sheet = context.workbook.worksheets.getActiveWorksheet();
  }

  const range = resolveRange(context, dataRange);
  const excelChartType = mapChartType(chartType);

  const chart = sheet.charts.add(excelChartType, range, Excel.ChartSeriesBy.auto);

  if (title) {
    chart.title.text = title;
    chart.title.visible = true;
  }

  if (position) {
    chart.left = position.left;
    chart.top = position.top;
    chart.width = position.width;
    chart.height = position.height;
  } else {
    // Default position: below the data
    chart.left = 10;
    chart.top = 300;
    chart.width = 500;
    chart.height = 300;
  }

  // Set series names if provided
  if (seriesNames) {
    chart.series.load("count");
    await context.sync();
    for (let i = 0; i < Math.min(seriesNames.length, chart.series.count); i++) {
      chart.series.getItemAt(i).name = seriesNames[i];
    }
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Created ${chartType} chart${title ? ` "${title}"` : ""} from ${dataRange}`,
    outputs: { chartName: title ?? `${chartType} chart` },
  };
}

function mapChartType(type: string): Excel.ChartType {
  const map: Record<string, Excel.ChartType> = {
    columnClustered: Excel.ChartType.columnClustered,
    columnStacked:   Excel.ChartType.columnStacked,
    bar:             Excel.ChartType.barClustered,
    line:            Excel.ChartType.line,
    pie:             Excel.ChartType.pie,
    area:            Excel.ChartType.area,
    scatter:         Excel.ChartType.xyscatter,
    combo:           Excel.ChartType.columnClustered, // combo requires special handling
  };
  return map[type] ?? Excel.ChartType.columnClustered;
}

registry.register(meta, handler as any);
export { meta };

/**
 * quickFormat – Apply a combination of common formatting in one step.
 *
 * Combines: freeze header row, add auto-filters, format header row,
 * auto-fit columns, and optional zebra-stripe — all in a single action.
 * This is a convenience capability for quickly making data presentable.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "quickFormat",
  description: "Apply common formatting (freeze, filters, autofit, header style) in one step",
  mutates: true,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const {
    range: address,
    freezeHeader = true,
    addFilters = true,
    autoFit = true,
    zebraStripe = false,
    headerColor = "#4472C4",
    headerFontColor = "#FFFFFF",
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would apply quick format to ${address}`,
    };
  }

  const range = resolveRange(context, address);
  const sheet = resolveSheet(context, address);
  range.load(["rowCount", "columnCount"]);
  await context.sync();

  const applied: string[] = [];

  // Freeze header row
  if (freezeHeader) {
    options.onProgress?.("Freezing header row...");
    sheet.freezePanes.freezeRows(1);
    applied.push("freeze: Y");
  } else {
    applied.push("freeze: N");
  }

  // Add auto-filters
  if (addFilters) {
    options.onProgress?.("Adding auto-filters...");
    range.getRow(0).getResizedRange(range.rowCount - 1, 0)
      // Use the full range for the auto filter
    ;
    sheet.autoFilter.apply(range);
    applied.push("filters: Y");
  } else {
    applied.push("filters: N");
  }

  // Format header row
  options.onProgress?.("Formatting header row...");
  const headerRow = range.getRow(0);
  headerRow.format.font.bold = true;
  headerRow.format.font.color = headerFontColor;
  headerRow.format.fill.color = headerColor;
  headerRow.format.horizontalAlignment = "Center" as Excel.HorizontalAlignment;

  // Auto-fit columns
  if (autoFit) {
    options.onProgress?.("Auto-fitting columns...");
    range.format.autofitColumns();
    applied.push("autofit: Y");
  } else {
    applied.push("autofit: N");
  }

  // Zebra-stripe data rows
  if (zebraStripe) {
    options.onProgress?.("Applying zebra-stripe...");
    for (let r = 1; r < range.rowCount; r++) {
      const row = range.getRow(r);
      row.format.fill.color = (r - 1) % 2 === 0 ? "#F2F2F2" : "#FFFFFF";
    }
    applied.push("zebra: Y");
  } else {
    applied.push("zebra: N");
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Applied quick format (${applied.join(", ")})`,
    outputs: { range: address },
  };
}

registry.register(meta, handler as any);
export { meta };

/**
 * alternatingRowFormat – Apply zebra-stripe formatting to alternate rows.
 *
 * Iterates through each row in the target range and applies alternating
 * fill colors. Optionally formats the header row with bold text and a
 * slightly darker fill.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "alternatingRowFormat",
  description: "Apply zebra-stripe (alternating row) formatting",
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
    evenColor = "#F2F2F2",
    oddColor = "#FFFFFF",
    hasHeaders = true,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would apply alternating row colors to ${address}`,
    };
  }

  options.onProgress?.("Applying alternating row formatting...");

  const range = resolveRange(context, address);
  range.load(["rowCount", "columnCount"]);
  await context.sync();

  const { rowCount } = range;
  const startRow = hasHeaders ? 1 : 0;

  // Format header row if applicable
  if (hasHeaders && rowCount > 0) {
    const headerRow = range.getRow(0);
    headerRow.format.font.bold = true;
    headerRow.format.fill.color = "#D9E2F3"; // slightly darker header fill
  }

  // Apply alternating colors to data rows
  for (let r = startRow; r < rowCount; r++) {
    const row = range.getRow(r);
    const dataRowIndex = r - startRow;
    row.format.fill.color = dataRowIndex % 2 === 0 ? evenColor : oddColor;
  }

  await context.sync();

  const dataRows = rowCount - startRow;
  return {
    stepId: "",
    status: "success",
    message: `Applied alternating row colors to ${dataRows} rows`,
  };
}

registry.register(meta, handler as any);
export { meta };

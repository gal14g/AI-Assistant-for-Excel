/**
 * highlightDuplicates — add a conditional-formatting rule that colors
 * duplicate values in a range.
 *
 * One-step equivalent of addConditionalFormat with a COUNTIF formula. Uses
 * a duplicateValues rule on ExcelApi 1.6+; falls back to a COUNTIF-based
 * custom formula on older versions.
 */

import {
  CapabilityMeta,
  HighlightDuplicatesParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "highlightDuplicates",
  description: "Highlight duplicate values in a range with a conditional-formatting rule",
  mutates: false,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.6",
};

async function handler(
  context: Excel.RequestContext,
  params: HighlightDuplicatesParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, fillColor = "#FFCCCC", fontColor = "#C50F1F" } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would highlight duplicates in ${address}` };
  }

  options.onProgress?.("Highlighting duplicates...");

  const range = resolveRange(context, address);
  range.load("address");
  await context.sync();

  // Parse the resolved address into an absolute range (for the COUNTIF's
  // FIRST argument — must stay put as the CF is applied to each cell) and a
  // relative cell ref (for the SECOND argument — must shift so the formula
  // evaluates against each cell's OWN value).
  //
  // Previous bug: we used `addr` directly (relative) so when the CF was
  // applied to the second cell, COUNTIF's range shifted by one row and
  // ended up counting values that weren't really duplicates.
  const addr = range.address;
  const sheetPrefix = addr.includes("!") ? addr.substring(0, addr.lastIndexOf("!") + 1) : "";
  const addrOnly = addr.includes("!") ? addr.substring(addr.lastIndexOf("!") + 1) : addr;

  // Build absolute range: "$A$1:$A$100"
  const absoluteRange = addrOnly.split(":").map((part) => {
    const m = part.match(/^\$?([A-Z]+)\$?(\d+)?$/);
    if (!m) return part;
    return `$${m[1]}${m[2] ? `$${m[2]}` : ""}`;
  }).join(":");
  const absRangeRef = `${sheetPrefix}${absoluteRange}`;

  // First cell as a RELATIVE reference so it shifts with each CF application.
  const firstCellMatch = addrOnly.split(":")[0].match(/^\$?([A-Z]+)\$?(\d+)/);
  const firstCell = firstCellMatch ? `${firstCellMatch[1]}${firstCellMatch[2]}` : "A1";

  // The Office.js `ConditionalFormatType` enum in older @types/office-js
  // doesn't expose `duplicateValues`, so we use the `custom` type with a
  // COUNTIF formula — works on ExcelApi 1.6+ and is equivalent visually.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cf: any = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
  cf.custom.rule.formula = `=COUNTIF(${absRangeRef},${firstCell})>1`;
  cf.custom.format.fill.color = fillColor;
  cf.custom.format.font.color = fontColor;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added duplicate-highlighting rule to ${address}.`,
    outputs: { range: address },
  };
}

registry.register(meta, handler as any);
export { meta };

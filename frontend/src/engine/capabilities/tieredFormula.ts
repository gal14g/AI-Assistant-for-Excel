/**
 * tieredFormula — generate tier-based formulas (tax brackets, grading,
 * commission, discount thresholds) as a single IFS per output cell.
 *
 * mode="lookup": pick the tier whose threshold ≤ source cell. Example —
 *   source 75000, tiers [{0,0.1},{50000,0.2},{100000,0.3}] → 0.2
 *   Formula: =IFS(A1>=100000,0.3, A1>=50000,0.2, A1>=0,0.1, TRUE, 0)
 *
 * mode="tax": cumulative tier tax. Each tier's value is a RATE applied to
 *   the slice of the source between that tier's threshold and the next.
 *   Example — source 120000, tiers [{0,0.1},{50000,0.2},{100000,0.3}]:
 *     0-50000  @ 10% = 5000
 *     50000-100000 @ 20% = 10000
 *     100000+ @ 30% = 6000 (of the 20k above 100k)
 *     total = 21000
 */

import {
  CapabilityMeta,
  TieredFormulaParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "tieredFormula",
  description: "Generate tier-based IFS formulas for tax brackets, grading, commission tiers, etc.",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.2",
};

/** Extract the first-cell reference from an address range — used as the
 *  anchor for the formula template. */
function firstCellRef(address: string): string | null {
  const addrOnly = address.includes("!") ? address.split("!").pop()! : address;
  const m = addrOnly.split(":")[0].match(/^\$?([A-Z]+)\$?(\d+)$/);
  if (!m) return null;
  return `${m[1]}${m[2]}`;
}

/** Strip any absolute-row marker so the formula can be auto-filled down. */
function relativeRef(cell: string): string {
  return cell.replace(/\$/g, "");
}

async function handler(
  context: Excel.RequestContext,
  params: TieredFormulaParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    sourceRange,
    outputRange,
    tiers,
    mode = "lookup",
    defaultValue = 0,
    hasHeaders = true,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would write ${mode} tier formula for ${tiers.length} tier(s) from ${sourceRange} → ${outputRange}.`,
    };
  }

  if (tiers.length === 0) {
    return { stepId: "", status: "error", message: "At least one tier is required." };
  }

  // Sort ascending by threshold for stable tax-mode computation, then we iterate descending for lookup mode.
  const sorted = [...tiers].sort((a, b) => a.threshold - b.threshold);

  // Resolve source/output and determine the data row range (skip headers if any).
  const src = resolveRange(context, sourceRange);
  const out = resolveRange(context, outputRange);
  src.load(["rowCount", "columnCount", "address"]);
  out.load(["rowCount", "columnCount", "address"]);
  await context.sync();

  if (src.columnCount !== 1 || out.columnCount !== 1) {
    return {
      stepId: "",
      status: "error",
      message: "sourceRange and outputRange must each be a single column.",
    };
  }

  // Anchor cell refs — used to build the formula template. We use the top-left
  // data cell of the source as the reference in the template (e.g. "A2"), and
  // autoFill handles the rest.
  const srcTopLeft = firstCellRef(src.address);
  if (!srcTopLeft) {
    return { stepId: "", status: "error", message: `Could not parse source address: ${src.address}` };
  }
  const m = srcTopLeft.match(/^([A-Z]+)(\d+)$/)!;
  const srcCol = m[1];
  const srcTopRow = Number(m[2]);
  const firstDataRow = hasHeaders ? srcTopRow + 1 : srcTopRow;
  const firstDataCell = `${srcCol}${firstDataRow}`;

  // Number of rows of output formula = min(src data rows, out rows).
  const srcDataRows = Math.max(0, src.rowCount - (hasHeaders ? 1 : 0));
  const outDataRows = Math.max(0, out.rowCount - (hasHeaders ? 1 : 0));
  const nRows = Math.min(srcDataRows, outDataRows);
  if (nRows === 0) {
    return { stepId: "", status: "success", message: "No data rows to fill." };
  }

  // Build the formula template in terms of `firstDataCell` (relative ref).
  let formula: string;
  const src0 = relativeRef(firstDataCell);
  if (mode === "lookup") {
    // Descending thresholds — first matching tier wins.
    const parts: string[] = [];
    for (let i = sorted.length - 1; i >= 0; i--) {
      parts.push(`${src0}>=${sorted[i].threshold},${sorted[i].value}`);
    }
    formula = `=IFS(${parts.join(",")},TRUE,${defaultValue})`;
  } else {
    // Tax mode — cumulative tier tax. Build as a sum of MAX(0, MIN(src, nextThreshold) - threshold) * rate
    // For the last tier, MIN with INF → just MAX(0, src - threshold).
    const segments: string[] = [];
    for (let i = 0; i < sorted.length; i++) {
      const t = sorted[i];
      const nextT = i + 1 < sorted.length ? sorted[i + 1].threshold : null;
      const upper = nextT === null ? src0 : `MIN(${src0},${nextT})`;
      segments.push(`MAX(0,${upper}-${t.threshold})*${t.value}`);
    }
    formula = `=${segments.join("+")}`;
  }

  // Write the formula at the top-left data cell of the output range, then
  // autofill down.
  const outTopLeft = firstCellRef(out.address);
  if (!outTopLeft) {
    return { stepId: "", status: "error", message: `Could not parse output address: ${out.address}` };
  }
  const outMatch = outTopLeft.match(/^([A-Z]+)(\d+)$/)!;
  const outFirstRow = Number(outMatch[2]) + (hasHeaders ? 1 : 0);
  const outCol = outMatch[1];
  const outSheetName = out.address.includes("!") ? out.address.split("!")[0] : null;

  try {
    const sheet = outSheetName
      ? context.workbook.worksheets.getItem(outSheetName.replace(/^'|'$/g, ""))
      : out.worksheet;
    const anchor = sheet.getRange(`${outCol}${outFirstRow}`);
    anchor.formulas = [[formula]];
    await context.sync();
    if (nRows > 1) {
      const filldown = anchor.getResizedRange(nRows - 1, 0);
      anchor.autoFill(filldown, Excel.AutoFillType.fillDefault);
      await context.sync();
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write tier formulas: ${msg}`, error: msg };
  }

  return {
    stepId: "",
    status: "success",
    message: `Wrote ${mode}-mode tier formulas (${tiers.length} tier${tiers.length === 1 ? "" : "s"}) to ${outputRange} over ${nRows} row(s).`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };

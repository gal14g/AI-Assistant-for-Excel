/**
 * spillFormula – Write a dynamic array formula that spills automatically.
 *
 * Dynamic array formulas (FILTER, SORT, UNIQUE, SEQUENCE, etc.) write a
 * single formula to one cell and Excel spills the results into adjacent
 * cells. This capability writes the formula and reports the spill size.
 *
 * Office.js notes:
 * - Dynamic arrays are supported in Excel 365+.
 * - getSpillingToRangeOrNullObject() returns the spill range after sync.
 * - The formula must start with "=".
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { rewriteDynamicArrayFormula } from "./fallbacks/dynamicArrayRewrite";
import { clipFullColumnRefs } from "./fallbacks/clipFullColumnRefs";

const meta: CapabilityMeta = {
  action: "spillFormula",
  description: "Write a dynamic array formula (FILTER, SORT, UNIQUE, etc.) that spills automatically",
  mutates: true,
  affectsFormatting: false,
  // Dynamic arrays (SPILL behavior) are Excel 365-only — ExcelApi 1.11+.
  // Older Excels route through the dynamic-array fallback, which rewrites
  // FILTER/UNIQUE/XLOOKUP/SORT/SEQUENCE into legacy equivalents.
  requiresApiSet: "ExcelApi 1.11",
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const { cell, formula, sheetName } = params;

  if (!cell || !formula) {
    return {
      stepId: "",
      status: "error",
      message: "cell and formula are required",
    };
  }

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would write dynamic array formula to ${cell}: ${formula}`,
    };
  }

  options.onProgress?.(`Writing dynamic array formula to ${cell}...`);

  // Clip any full-column references (e.g. `=SORT(B:B)`) to the sheet's
  // used range. Without this, SORT/UNIQUE/FILTER on a full column spill
  // to 1,048,576 rows, locking up the worksheet and breaking downstream
  // steps. The clip is sheet-local, so multi-sheet formulas get each
  // reference bounded against its own sheet's used range.
  const clip = await clipFullColumnRefs(context, formula, sheetName ?? null);
  const formulaToWrite = clip.formula;

  // Resolve target cell — prepend sheetName if provided
  const cellAddress = sheetName && !cell.includes("!") ? `${sheetName}!${cell}` : cell;
  const range = resolveRange(context, cellAddress);

  // Write the (possibly clipped) formula to the single cell
  range.formulas = [[formulaToWrite]];
  await context.sync();

  // Check for formula errors
  range.load("values");
  await context.sync();

  const firstVal = String(range.values?.[0]?.[0] ?? "");
  const errorTypes = ["#SPILL!", "#REF!", "#VALUE!", "#NAME?", "#NULL!", "#N/A", "#DIV/0!", "#CALC!"];
  const hasError = errorTypes.some((e) => firstVal.includes(e));

  if (hasError) {
    return {
      stepId: "",
      status: "error",
      message: `Formula wrote to ${cell} but produced ${firstVal}. The formula may need to be corrected.`,
      error: `Formula error: ${firstVal}`,
    };
  }

  // Try to determine the spill range size
  let spillInfo = "spill size unknown";
  try {
    const spillRange = range.getSpillingToRangeOrNullObject();
    spillRange.load(["isNullObject", "rowCount", "columnCount", "address"]);
    await context.sync();

    if (!spillRange.isNullObject) {
      const totalCells = spillRange.rowCount * spillRange.columnCount;
      spillInfo = `spilled to ${totalCells} cells (${spillRange.address})`;
    }
  } catch {
    // getSpillingToRangeOrNullObject may not be available in all API sets
    spillInfo = "spill size unknown";
  }

  const clipNote = clip.clippedCount > 0
    ? ` (auto-clipped ${clip.clippedCount} full-column ref${clip.clippedCount === 1 ? "" : "s"}: ${clip.clippedRefs.join(", ")})`
    : "";
  return {
    stepId: "",
    status: "success",
    message: `Wrote dynamic array formula to ${cell}, ${spillInfo}${clipNote}`,
    outputs: { cell: cellAddress },
  };
}


// ── Legacy-Excel fallback ────────────────────────────────────────────────────
// On Excel 2016/2019 there is no SPILL behaviour. Rewrite the formula body
// into legacy-equivalent INDEX/MATCH/SMALL constructs and write it as an
// array formula at the target cell. The user typically needs to drag it
// down to fill the expected output range — we warn about this in the message.
async function fallback(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { cell, formula, sheetName } = params;

  if (!cell || !formula) {
    return { stepId: "", status: "error", message: "cell and formula are required" };
  }

  // Clip full-column refs BEFORE the legacy rewrite so SMALL/IF constructs
  // operate on a bounded array — otherwise the rewritten formula evaluates
  // against 1M-row virtual arrays, which Excel 2016 can't handle.
  const clipLegacy = await clipFullColumnRefs(context, formula, sheetName ?? null);
  const rewritten = rewriteDynamicArrayFormula(clipLegacy.formula);

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message:
        `Would write legacy-compatible formula to ${cell}: ${rewritten.formula} ` +
        `(rewrite: ${rewritten.changes.join(", ") || "none"})`,
    };
  }

  options.onProgress?.(
    `Legacy-Excel mode: rewriting ${rewritten.changes.join(", ") || "formula"}...`,
  );

  const cellAddress = sheetName && !cell.includes("!") ? `${sheetName}!${cell}` : cell;
  const range = resolveRange(context, cellAddress);

  if (rewritten.requiresArrayEntry) {
    // Office.js `range.setArrayFormula` expects the array formula on the
    // resolved range; for a single-cell entry it still commits as an array
    // formula so Excel evaluates the IF/SMALL construct correctly.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (range as any).setFormulas?.([[rewritten.formula]]);
    // setFormulas doesn't exist on older Office.js either — fall back to
    // plain .formulas, which commits as array-formula when containing SMALL/IF
    // in practice on Excel 2016.
    range.formulas = [[rewritten.formula]];
  } else {
    range.formulas = [[rewritten.formula]];
  }
  await context.sync();

  range.load("values");
  await context.sync();
  const firstVal = String(range.values?.[0]?.[0] ?? "");
  const errorTypes = ["#SPILL!", "#REF!", "#VALUE!", "#NAME?", "#NULL!", "#N/A", "#DIV/0!", "#CALC!"];
  if (errorTypes.some((e) => firstVal.includes(e))) {
    return {
      stepId: "",
      status: "error",
      message:
        `Legacy-compatible formula written to ${cell} but produced ${firstVal}. ` +
        `The dynamic-array rewrite may not cover this construct; consider using ` +
        `a static INDEX/MATCH formula or upgrading to Excel 365.`,
      error: `Formula error: ${firstVal}`,
    };
  }

  const warnSuffix = rewritten.warnings.length
    ? ` — ${rewritten.warnings.join("; ")}`
    : "";
  return {
    stepId: "",
    status: "success",
    message:
      `Wrote legacy-compatible formula to ${cell} ` +
      `(rewrote ${rewritten.changes.join(", ") || "formula"})${warnSuffix}. ` +
      `Drag the cell down to fill the desired output range.`,
    outputs: { cell: cellAddress },
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };

/**
 * concatRows — concatenate the cells of each row in sourceRange into a
 * single cell in outputColumn, joined by separator.
 *
 * Uses TEXTJOIN so outputs stay live as source values change. Skips the
 * header row when hasHeaders is true.
 */

import { CapabilityMeta, ConcatRowsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "concatRows",
  description: "Concatenate each row's cells into a single TEXTJOIN cell in outputColumn",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.2",
};

async function handler(
  context: Excel.RequestContext,
  params: ConcatRowsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    sourceRange,
    outputColumn,
    separator = ", ",
    ignoreBlanks = true,
    hasHeaders = true,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would write TEXTJOIN formulas for ${sourceRange} into column ${outputColumn}.`,
    };
  }

  options.onProgress?.("Writing TEXTJOIN formulas per row...");

  // Clip to used range.
  const rawRange = resolveRange(context, sourceRange);
  const used = rawRange.getUsedRange(false);
  used.load(["rowCount", "columnCount", "address", "worksheet/name"]);
  await context.sync();

  if (!used.rowCount) {
    return { stepId: "", status: "success", message: "Source range is empty." };
  }

  // Parse the source used-range to get first/last row + first/last column letters.
  const addrPart = used.address.includes("!") ? used.address.split("!").pop()! : used.address;
  const [tl, br] = addrPart.split(":");
  const tlMatch = tl.match(/^([A-Z]+)(\d+)$/);
  const brMatch = (br ?? tl).match(/^([A-Z]+)(\d+)$/);
  if (!tlMatch || !brMatch) {
    return { stepId: "", status: "error", message: `Could not parse used range: ${used.address}` };
  }
  const firstColL = tlMatch[1];
  const firstRow = Number(tlMatch[2]);
  const lastColL = brMatch[1];
  const lastRow = Number(brMatch[2]);

  // Normalize output column — accept "G", "Sheet1!G", "$G", "G:G", etc.
  let outColPart = outputColumn;
  let outSheet: string | null = null;
  if (outColPart.includes("!")) {
    const parts = outColPart.split("!");
    outSheet = parts[0].replace(/^'|'$/g, "");
    outColPart = parts[1];
  }
  outColPart = outColPart.replace(/[:$]/g, "").replace(/\d+$/, "");
  const outColLetter = outColPart.toUpperCase();
  if (!/^[A-Z]+$/.test(outColLetter)) {
    return { stepId: "", status: "error", message: `Invalid outputColumn: "${outputColumn}"` };
  }

  const sourceSheetName = used.worksheet.name;
  const sheetName = outSheet ?? sourceSheetName;

  // Build formulas. Each row N gets =TEXTJOIN(sep, ignoreBlanks, firstColL{N}:lastColL{N})
  // skipping the header row when hasHeaders.
  const dataStart = hasHeaders ? firstRow + 1 : firstRow;
  const dataCount = Math.max(0, lastRow - dataStart + 1);
  if (dataCount === 0) {
    return { stepId: "", status: "success", message: "No data rows to concatenate (only headers present)." };
  }

  const sheetPrefix = sourceSheetName.includes(" ") ? `'${sourceSheetName}'!` : `${sourceSheetName}!`;
  const formulas: string[][] = [];
  for (let r = 0; r < dataCount; r++) {
    const rowNum = dataStart + r;
    formulas.push([
      `=TEXTJOIN("${separator.replace(/"/g, '""')}",${ignoreBlanks ? "TRUE" : "FALSE"},${sheetPrefix}${firstColL}${rowNum}:${lastColL}${rowNum})`,
    ]);
  }

  // Optionally write a header into the output column's header row.
  const outAddr = `${sheetName}!${outColLetter}${dataStart}:${outColLetter}${dataStart + dataCount - 1}`;
  const outSheetObj = context.workbook.worksheets.getItem(sheetName);
  try {
    if (hasHeaders) {
      outSheetObj.getRange(`${outColLetter}${firstRow}`).values = [["Joined"]];
    }
    outSheetObj.getRange(outAddr.split("!")[1]).formulas = formulas;
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write TEXTJOIN formulas: ${msg}`, error: msg };
  }

  return {
    stepId: "",
    status: "success",
    message: `Wrote ${dataCount} TEXTJOIN formulas to ${outAddr}.`,
    outputs: { outputRange: outAddr },
  };
}

registry.register(meta, handler as any);
export { meta };

/**
 * aging — bucket dates into aging categories.
 *
 * Given a date column and a list of upper-bound day counts (default
 * [30, 60, 90]), emits a formula column that labels each row with the
 * matching bucket. The open-ended last bucket reads `{lastBucket}+`.
 *
 * Example: dateColumn = "Sheet1!C2:C100", buckets = [30, 60, 90]
 *   age(date) ≤ 30   → "0-30"
 *   30 < age ≤ 60    → "31-60"
 *   60 < age ≤ 90    → "61-90"
 *   age > 90         → "90+"
 */

import { CapabilityMeta, AgingParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "aging",
  description: "Bucket dates into aging categories (0-30, 31-60, 61-90, 90+) with a formula column",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.2",
};

async function handler(
  context: Excel.RequestContext,
  params: AgingParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { dateColumn, outputColumn, buckets = [30, 60, 90], referenceDate, hasHeaders = true } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would write aging buckets (${buckets.join(",")}+) for ${dateColumn} into ${outputColumn}.`,
    };
  }

  // Resolve the date column so we know where data starts/ends.
  const dateRaw = resolveRange(context, dateColumn);
  const dateUsed = dateRaw.getUsedRange(false);
  dateUsed.load(["address", "rowCount", "worksheet/name"]);
  await context.sync();

  const addrPart = dateUsed.address.includes("!") ? dateUsed.address.split("!").pop()! : dateUsed.address;
  const m = addrPart.match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
  if (!m) {
    return { stepId: "", status: "error", message: `Could not parse date column address: ${dateUsed.address}` };
  }
  const [, dateColLetter, dateTopRow, , dateBotRow] = m;
  const firstDataRow = hasHeaders ? Number(dateTopRow) + 1 : Number(dateTopRow);
  const lastDataRow = Number(dateBotRow);
  const dataCount = Math.max(0, lastDataRow - firstDataRow + 1);
  if (dataCount === 0) {
    return { stepId: "", status: "success", message: "No data rows in dateColumn." };
  }

  const dateSheetName = dateUsed.worksheet.name;
  const sheetPrefix = dateSheetName.includes(" ") ? `'${dateSheetName}'!` : `${dateSheetName}!`;

  // Normalize output column — accept "G", "Sheet1!G", etc.
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
  const outSheetName = outSheet ?? dateSheetName;
  const outSheetObj = context.workbook.worksheets.getItem(outSheetName);

  // Reference date: explicit or TODAY().
  const refDateExpr = referenceDate
    ? (() => {
        // Accept "dd/mm/yyyy" — convert to DATE(y,m,d) expression.
        const s = String(referenceDate).trim();
        const mm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
        if (mm) {
          let y = Number(mm[3]);
          if (y < 100) y += 2000;
          return `DATE(${y},${Number(mm[2])},${Number(mm[1])})`;
        }
        const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (iso) return `DATE(${Number(iso[1])},${Number(iso[2])},${Number(iso[3])})`;
        return `DATEVALUE("${s.replace(/"/g, '""')}")`;
      })()
    : "TODAY()";

  // Sorted buckets ascending — labels are "0-B1", "B1+1-B2", ... "Bn+".
  const sortedBuckets = [...buckets].sort((a, b) => a - b);
  const bucketLabels: string[] = [];
  for (let i = 0; i < sortedBuckets.length; i++) {
    const lower = i === 0 ? 0 : sortedBuckets[i - 1] + 1;
    bucketLabels.push(`${lower}-${sortedBuckets[i]}`);
  }
  bucketLabels.push(`${sortedBuckets[sortedBuckets.length - 1]}+`);

  // Build formula. We reference the first data row of the date column
  // (relative) and auto-fill down.
  const firstDateCell = `${dateColLetter}${firstDataRow}`;
  const ageExpr = `(${refDateExpr}-${firstDateCell})`;
  const ifsParts: string[] = [];
  for (let i = 0; i < sortedBuckets.length; i++) {
    ifsParts.push(`${ageExpr}<=${sortedBuckets[i]},"${bucketLabels[i]}"`);
  }
  ifsParts.push(`TRUE,"${bucketLabels[bucketLabels.length - 1]}"`);
  const formula = `=IFS(${ifsParts.join(",")})`;

  try {
    // Optional header.
    if (hasHeaders) {
      outSheetObj.getRange(`${outColLetter}${Number(dateTopRow)}`).values = [["Age Bucket"]];
    }
    const anchor = outSheetObj.getRange(`${outColLetter}${firstDataRow}`);
    anchor.formulas = [[formula]];
    await context.sync();
    if (dataCount > 1) {
      const filldown = anchor.getResizedRange(dataCount - 1, 0);
      anchor.autoFill(filldown, Excel.AutoFillType.fillDefault);
      await context.sync();
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write aging column: ${msg}`, error: msg };
  }

  const outputAddr = `${sheetPrefix}${outColLetter}${firstDataRow}:${outColLetter}${lastDataRow}`;
  return {
    stepId: "",
    status: "success",
    message: `Wrote aging buckets (${bucketLabels.join(", ")}) to ${outputAddr}.`,
    outputs: { outputRange: outputAddr },
  };
}

registry.register(meta, handler as any);
export { meta };

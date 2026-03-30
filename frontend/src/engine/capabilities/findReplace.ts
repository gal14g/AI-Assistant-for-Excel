/**
 * findReplace – Find and replace text in a range or sheet.
 *
 * Office.js notes:
 * - Range.replaceAll() is available in ExcelApi 1.9+ (currently preview).
 *   We fall back to a read-modify-write approach for broader compatibility.
 * - This approach preserves formatting since we only write values back.
 */

import { CapabilityMeta, FindReplaceParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "findReplace",
  description: "Find and replace text values in a range",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: FindReplaceParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, sheetName, find, replace, matchCase = false, matchEntireCell = false } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would replace "${find}" with "${replace}" in ${address ?? "entire sheet"}`,
    };
  }

  options.onProgress?.(`Finding "${find}"...`);

  // Validate and resolve target sheet
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

  // If address includes a sheet qualifier (e.g. "תוכנה!A:I"), use resolveRange
  // so it is correctly resolved regardless of active sheet.
  // If address is plain (e.g. "A1:C10"), use the already-resolved sheet object.
  const range = address
    ? address.includes("!")
      ? resolveRange(context, address)
      : sheet.getRange(address)
    : sheet.getUsedRange();
  // Load both raw values (numbers for dates) and displayed text (what the user sees).
  // Matching is done against text so "19/04/2026" matches regardless of whether the
  // cell stores a serial number or a literal string.
  range.load(["values", "text"]);
  await context.sync();

  const values = range.values ?? [];
  const texts = range.text ?? [];
  let replacements = 0;

  const findStr = matchCase ? find : find.toLowerCase();

  // Pre-parse find/replace as dates for serial-number replacement
  const findDate = parseDateString(find);
  const replaceDate = parseDateString(replace);
  const replaceSerial = findDate && replaceDate ? dateToExcelSerial(replaceDate) : null;

  const isFormula = replace.startsWith("=");

  for (let ri = 0; ri < texts.length; ri++) {
    const textRow = texts[ri] ?? [];
    const valRow = values[ri] ?? [];

    for (let ci = 0; ci < textRow.length; ci++) {
      const displayText = textRow[ci];
      if (typeof displayText !== "string" || displayText === "") continue;

      const displayStr = matchCase ? displayText : displayText.toLowerCase();
      let matched = false;

      if (matchEntireCell) {
        matched = displayStr === findStr;
      } else {
        matched = displayStr.includes(findStr);
      }

      if (!matched) continue;

      const rawVal = valRow[ci];
      const cell = range.getCell(ri, ci);

      // Decide what value to write based on the cell's underlying type:
      // - If the raw value is a number (date serial), write a new serial to preserve formatting.
      // - If the raw value is a string, do a normal string replacement.
      if (typeof rawVal === "number" && replaceSerial != null) {
        cell.values = [[replaceSerial]];
      } else if (isFormula) {
        cell.formulas = [[replace]];
      } else if (matchEntireCell) {
        cell.values = [[replace]];
      } else {
        // Partial string replacement within the displayed text
        const regex = new RegExp(escapeRegex(find), matchCase ? "g" : "gi");
        const original = typeof rawVal === "string" ? rawVal : displayText;
        cell.values = [[original.replace(regex, replace)]];
      }

      replacements++;
    }
  }

  if (replacements > 0) {
    await context.sync();
  }

  return {
    stepId: "",
    status: "success",
    message: `Replaced ${replacements} occurrence(s) of "${find}" with "${replace}"`,
  };
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ── Date helpers ──────────────────────────────────────────────────────────────
// Excel stores dates as serial numbers (days since 1900-01-00, with Lotus bug).
// Users search by formatted text (e.g. "19/04/2026") but .values returns a
// number.  These helpers bridge the gap.

const DATE_PATTERNS = [
  // dd/mm/yyyy  or  dd-mm-yyyy  or  dd.mm.yyyy
  /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/,
  // yyyy-mm-dd  or  yyyy/mm/dd
  /^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$/,
];

/** Try to parse a user-typed date string into {day, month, year}. */
function parseDateString(s: string): { day: number; month: number; year: number } | null {
  // dd/mm/yyyy  dd-mm-yyyy  dd.mm.yyyy
  let m = s.match(DATE_PATTERNS[0]);
  if (m) return { day: +m[1], month: +m[2], year: +m[3] };
  // yyyy-mm-dd  yyyy/mm/dd
  m = s.match(DATE_PATTERNS[1]);
  if (m) return { day: +m[3], month: +m[2], year: +m[1] };
  return null;
}

/** Convert Excel serial number → JS Date (accounts for Lotus 1-2-3 bug). */
function excelSerialToDate(serial: number): Date | null {
  if (serial < 1 || serial > 2958465) return null; // out of range
  // Excel epoch: serial 1 = Jan 1 1900, so serial 0 = Dec 31 1899 ("Jan 0 1900").
  // Serial 60 = non-existent Feb 29 1900 (Lotus 1-2-3 bug) — skip it.
  const adjusted = serial > 60 ? serial - 1 : serial;
  const ms = Date.UTC(1899, 11, 31) + adjusted * 86400000;
  return new Date(ms);
}

/** Convert {day, month, year} → Excel serial number. */
function dateToExcelSerial(d: { day: number; month: number; year: number }): number {
  const ms = Date.UTC(d.year, d.month - 1, d.day);
  const epoch = Date.UTC(1899, 11, 31); // Dec 31 1899 = Excel's "Jan 0 1900"
  let serial = Math.round((ms - epoch) / 86400000);
  if (serial >= 60) serial += 1; // account for Lotus 1-2-3 fake Feb 29 1900
  return serial;
}

/**
 * Check if an Excel numeric cell value matches the find date string.
 * Returns true when the serial date represents the same calendar date.
 */
function numericCellMatchesDate(
  cellValue: number,
  findDate: { day: number; month: number; year: number }
): boolean {
  const d = excelSerialToDate(cellValue);
  if (!d) return false;
  return (
    d.getUTCFullYear() === findDate.year &&
    d.getUTCMonth() + 1 === findDate.month &&
    d.getUTCDate() === findDate.day
  );
}

registry.register(meta, handler as any);
export { meta };

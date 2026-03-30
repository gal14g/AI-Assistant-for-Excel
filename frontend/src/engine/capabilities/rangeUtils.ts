/**
 * Shared range resolution utility.
 *
 * Handles plain, sheet-qualified, and workbook-qualified addresses:
 *   "A1:C10"                                     → active sheet
 *   "Sheet1!A1:C10"                               → Sheet1
 *   "'My Sheet'!A1:C10"                           → My Sheet (quoted)
 *   "[WorkbookName.xlsx]Sheet1!A1:C10"            → Sheet1 in this workbook
 *   "[WorkbookName.xlsx]'My Sheet'!A1:C10"        → My Sheet (quoted + workbook)
 *
 * Also strips LLM-generated [[...]] token markers that may appear literally:
 *   "[[Sheet1!A:A]]"                              → "Sheet1!A:A"
 *   "[[WorkbookName.xlsx]Sheet1!A:A]"             → "[WorkbookName.xlsx]Sheet1!A:A"
 *
 * Office.js getItem() only accepts the plain sheet name, so we strip any
 * leading [WorkbookName...] prefix before calling it.
 */

/**
 * Remove [[...]] token notation that the LLM sometimes includes literally
 * when it copies range tokens from the user message into plan params.
 *
 * Two forms:
 *  "[[Sheet!Range]]"          → "Sheet!Range"          (non-qualified)
 *  "[[WB.xlsx]Sheet!Range]"   → "[WB.xlsx]Sheet!Range" (workbook-qualified)
 */
function normalizeAddress(address: string): string {
  address = address.trim();
  if (!address.startsWith("[[")) return address;

  // Try double-closing-bracket form first: "[[...]]"
  // This correctly handles BOTH:
  //   "[[Sheet!Range]]"              → "Sheet!Range"
  //   "[[[WB.xlsx]Sheet!Range]]"     → "[WB.xlsx]Sheet!Range"
  // The greedy (.+) stops before the final "]]", so workbook brackets inside
  // the capture are preserved and no trailing "]" is accidentally included.
  const plainMatch = address.match(/^\[\[(.+)\]\]$/);
  if (plainMatch) return plainMatch[1];

  // Single-closing-bracket form: "[[WB.xlsx]Sheet!Range]" (older LLM format)
  // Must be checked after plainMatch because it would also match "[[...]]"
  // but would include an extra trailing "]" in the capture group.
  const wbMatch = address.match(/^\[\[(\[.+)\]$/);
  if (wbMatch) return wbMatch[1];

  // Fallback: strip one leading "[" and trailing "]"
  return address.slice(1, address.endsWith("]") ? -1 : undefined);
}

/**
 * Strip the workbook qualifier from an address so it can be used in
 * Excel formula strings that reference ranges in the same workbook.
 *   "[WorkbookName.xlsx]Sheet1!A:A" → "Sheet1!A:A"
 */
export function stripWorkbookQualifier(address: string): string {
  return normalizeAddress(address).replace(/^\[.*?\]/, "");
}

/**
 * Ensure the sheet name portion of a range reference is single-quoted
 * when used inside an Excel formula string, which is required whenever
 * the sheet name contains:
 *   - Non-ASCII characters (Hebrew, Arabic, accented chars, etc.)
 *   - Spaces
 *   - Starts with a digit
 *   - Any punctuation except underscore and dot
 *
 * Examples:
 *   "Sheet1!A:A"       → "Sheet1!A:A"      (safe — unchanged)
 *   "גיליון1!A:A"      → "'גיליון1'!A:A"   (Hebrew — must quote)
 *   "My Sheet!A1"      → "'My Sheet'!A1"    (space — must quote)
 *   "'Sheet1'!A:A"     → "'Sheet1'!A:A"     (already quoted — unchanged)
 *   "A:A"              → "A:A"              (no sheet part — unchanged)
 */
export function quoteSheetInRef(ref: string): string {
  const bangIdx = ref.lastIndexOf("!");
  if (bangIdx === -1) return ref; // no sheet qualifier
  const sheetPart = ref.substring(0, bangIdx);
  const cellPart  = ref.substring(bangIdx + 1);
  // Already quoted
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) return ref;
  // Only pure ASCII identifier characters need no quoting
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetPart)) return ref;
  // Escape any internal single quotes as '' then wrap
  const escaped = sheetPart.replace(/'/g, "''");
  return `'${escaped}'!${cellPart}`;
}

/**
 * Get the Worksheet object for the sheet in the given address.
 * Works with plain, sheet-qualified, and workbook-qualified addresses.
 */
export function resolveSheet(
  context: Excel.RequestContext,
  address: string
): Excel.Worksheet {
  address = normalizeAddress(address);
  if (!address.includes("!")) {
    return context.workbook.worksheets.getActiveWorksheet();
  }
  const bangIdx = address.lastIndexOf("!");
  let sheetPart = address.substring(0, bangIdx);
  const wbMatch = sheetPart.match(/^\[.*?\](.+)$/);
  if (wbMatch) sheetPart = wbMatch[1];
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) sheetPart = sheetPart.slice(1, -1);
  return context.workbook.worksheets.getItem(sheetPart);
}

export function resolveRange(
  context: Excel.RequestContext,
  address: string
): Excel.Range {
  address = normalizeAddress(address);
  if (!address.includes("!")) {
    // No sheet qualifier — use the active worksheet
    return context.workbook.worksheets.getActiveWorksheet().getRange(address);
  }

  const bangIdx = address.lastIndexOf("!");
  let sheetPart = address.substring(0, bangIdx);
  const cellPart = address.substring(bangIdx + 1);

  // Strip workbook qualifier: "[WorkbookName.xlsx]SheetName" → "SheetName"
  const wbMatch = sheetPart.match(/^\[.*?\](.+)$/);
  if (wbMatch) {
    sheetPart = wbMatch[1];
  }

  // Strip surrounding single quotes from sheet names with spaces
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
    sheetPart = sheetPart.slice(1, -1);
  }

  return context.workbook.worksheets.getItem(sheetPart).getRange(cellPart);
}

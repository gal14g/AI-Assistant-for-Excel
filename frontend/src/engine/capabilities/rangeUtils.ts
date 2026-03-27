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
  // Workbook-qualified: inner address starts with "[", so outer token is "[[...]...]"
  // Match "[[" then capture "[..." then single "]" at end
  const wbMatch = address.match(/^\[\[(\[.+)\]$/);
  if (wbMatch) return wbMatch[1];
  // Non-qualified: outer token is "[[...]]"
  const plainMatch = address.match(/^\[\[(.+)\]\]$/);
  if (plainMatch) return plainMatch[1];
  // Fallback: strip leading "[" and trailing "]" (partial bracket inclusion)
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

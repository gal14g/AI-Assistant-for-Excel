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
 * Office.js getItem() only accepts the plain sheet name, so we strip any
 * leading [WorkbookName...] prefix before calling it.
 */
/**
 * Strip the workbook qualifier from an address so it can be used in
 * Excel formula strings that reference ranges in the same workbook.
 *   "[WorkbookName.xlsx]Sheet1!A:A" → "Sheet1!A:A"
 */
export function stripWorkbookQualifier(address: string): string {
  return address.replace(/^\[.*?\]/, "");
}

/**
 * Get the Worksheet object for the sheet in the given address.
 * Works with plain, sheet-qualified, and workbook-qualified addresses.
 */
export function resolveSheet(
  context: Excel.RequestContext,
  address: string
): Excel.Worksheet {
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

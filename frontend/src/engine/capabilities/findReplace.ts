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

  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  const range = address ? sheet.getRange(address) : sheet.getUsedRange();
  range.load("values");
  await context.sync();

  const values = range.values ?? [];
  let replacements = 0;

  const findStr = matchCase ? find : find.toLowerCase();

  const newValues = values.map((row) =>
    row.map((cell) => {
      if (typeof cell !== "string") return cell;
      const cellStr = matchCase ? cell : cell.toLowerCase();

      if (matchEntireCell) {
        if (cellStr === findStr) {
          replacements++;
          return replace;
        }
        return cell;
      }

      if (cellStr.includes(findStr)) {
        // Use regex for case-insensitive replacement
        const regex = new RegExp(escapeRegex(find), matchCase ? "g" : "gi");
        const result = cell.replace(regex, replace);
        if (result !== cell) replacements++;
        return result;
      }
      return cell;
    })
  );

  if (replacements > 0) {
    range.values = newValues;
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

registry.register(meta, handler as any);
export { meta };

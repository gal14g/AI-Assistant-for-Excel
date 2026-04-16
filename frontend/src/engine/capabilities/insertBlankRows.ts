/**
 * insertBlankRows — insert blank rows at explicit positions or every Nth
 * row within a range.
 *
 * Two modes:
 *  - explicit positions: insert `count` blank rows ABOVE each 1-based row
 *    number in `positions`. Positions are applied in DESCENDING order so
 *    earlier inserts don't shift later positions.
 *  - interval: insert `count` blank rows after every `every` data rows in
 *    `range`.
 */

import { CapabilityMeta, InsertBlankRowsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { registerInverseOp } from "../snapshot";

const meta: CapabilityMeta = {
  action: "insertBlankRows",
  description: "Insert blank rows at explicit row numbers or every Nth row in a range",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.1",
};

async function handler(
  context: Excel.RequestContext,
  params: InsertBlankRowsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sheetName, positions, every, range: rangeAddr, count = 1 } = params;

  if (options.dryRun) {
    const desc = positions?.length
      ? `${positions.length} explicit position(s)`
      : every && rangeAddr
      ? `every ${every} rows within ${rangeAddr}`
      : "(no valid spec)";
    return { stepId: "", status: "success", message: `Would insert blank rows at ${desc}` };
  }

  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  // Build list of 1-based row numbers to insert at. Applied in DESCENDING
  // order so later inserts don't shift earlier positions.
  let targets: number[] = [];
  if (positions && positions.length > 0) {
    targets = [...positions];
  } else if (every && rangeAddr) {
    const range = sheet.getRange(rangeAddr);
    range.load(["rowCount", "address"]);
    await context.sync();
    const addrPart = range.address.includes("!") ? range.address.split("!").pop()! : range.address;
    const tlMatch = addrPart.split(":")[0].match(/^[A-Z]+?(\d+)$/);
    if (!tlMatch) {
      return { stepId: "", status: "error", message: `Could not parse range: ${range.address}` };
    }
    const firstRow = Number(tlMatch[1]);
    for (let i = every; i < range.rowCount; i += every) targets.push(firstRow + i);
  } else {
    return {
      stepId: "",
      status: "error",
      message: "Provide either `positions` OR `every` + `range`.",
    };
  }

  targets.sort((a, b) => b - a); // descending

  try {
    for (const row of targets) {
      const rangeAt = sheet.getRange(`${row}:${row + count - 1}`);
      rangeAt.insert(Excel.InsertShiftDirection.down);
    }
    await context.sync();
    // Undo = delete the rows we just inserted. Order matters: we inserted
    // in DESCENDING order so we register inverse ops in ASCENDING order —
    // the rollback iterates inverse ops in reverse, bringing us back to
    // descending deletes that don't shift later positions.
    sheet.load("name");
    await context.sync();
    const ascending = [...targets].sort((a, b) => a - b);
    for (const row of ascending) {
      registerInverseOp({
        kind: "deleteRows",
        sheetName: sheet.name,
        rangeAddress: `${row}:${row + count - 1}`,
      });
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to insert blank rows: ${msg}`, error: msg };
  }

  return {
    stepId: "",
    status: "success",
    message: `Inserted ${targets.length * count} blank row(s) at ${targets.length} position(s).`,
    outputs: { rowsInserted: targets.length * count },
  };
}

registry.register(meta, handler as any);
export { meta };

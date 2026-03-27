/**
 * Snapshot & Rollback
 *
 * Before any write operation, the executor captures a snapshot of the
 * affected cell ranges. If the user requests an undo or the plan fails
 * partway, the snapshot is used to restore the original state.
 *
 * IMPORTANT Office.js notes:
 * - context.sync() must be called to flush reads before we can capture values.
 * - We store values and number formats. We do NOT snapshot formulas separately
 *   because writing back values is the safe rollback path (formulas may have
 *   had volatile references).
 * - Formatting (fill, font, borders) is NOT snapshotted because our design
 *   principle is to never touch formatting unless explicitly requested.
 */

import { CellSnapshot, PlanSnapshot } from "./types";

/** In-memory stack of snapshots. Most recent at the end. */
const snapshotStack: PlanSnapshot[] = [];

const MAX_SNAPSHOTS = 20;

/**
 * Capture a snapshot of one or more ranges before a write operation.
 * Must be called inside an Excel.run context.
 */
export async function captureSnapshot(
  context: Excel.RequestContext,
  planId: string,
  rangeAddresses: string[]
): Promise<PlanSnapshot> {
  const cells: CellSnapshot[] = [];

  for (const address of rangeAddresses) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load(["values", "numberFormat", "address"]);
  }

  await context.sync();

  // Re-get and read the loaded values
  for (const address of rangeAddresses) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load(["values", "numberFormat", "address"]);
  }
  await context.sync();

  // Rebuild from loaded proxy objects
  for (const address of rangeAddresses) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load(["values", "numberFormat", "address"]);
    await context.sync();

    cells.push({
      range: range.address,
      values: range.values as (string | number | boolean | null)[][],
      numberFormats: range.numberFormat as string[][],
    });
  }

  const snapshot: PlanSnapshot = {
    planId,
    timestamp: new Date().toISOString(),
    cells,
  };

  snapshotStack.push(snapshot);

  // Evict oldest snapshots if we exceed the limit
  while (snapshotStack.length > MAX_SNAPSHOTS) {
    snapshotStack.shift();
  }

  return snapshot;
}

/**
 * Optimized snapshot capture that batches all range loads into a single sync.
 */
export async function captureSnapshotBatched(
  context: Excel.RequestContext,
  planId: string,
  rangeAddresses: string[]
): Promise<PlanSnapshot> {
  const ranges: Excel.Range[] = [];

  for (const address of rangeAddresses) {
    // Parse sheet name from address if present (e.g. "Sheet1!A1:B5")
    const sheet = parseSheetFromAddress(address, context);
    const rangeRef = address.includes("!") ? address.split("!")[1] : address;
    const range = sheet.getRange(rangeRef);
    range.load(["values", "numberFormat", "address"]);
    ranges.push(range);
  }

  await context.sync();

  const cells: CellSnapshot[] = ranges.map((range) => ({
    range: range.address,
    values: range.values as (string | number | boolean | null)[][],
    numberFormats: range.numberFormat as string[][],
  }));

  const snapshot: PlanSnapshot = {
    planId,
    timestamp: new Date().toISOString(),
    cells,
  };

  snapshotStack.push(snapshot);
  while (snapshotStack.length > MAX_SNAPSHOTS) {
    snapshotStack.shift();
  }

  return snapshot;
}

/**
 * Restore the most recent snapshot (undo last plan execution).
 */
export async function rollbackLastSnapshot(
  context: Excel.RequestContext
): Promise<PlanSnapshot | null> {
  const snapshot = snapshotStack.pop();
  if (!snapshot) return null;

  for (const cell of snapshot.cells) {
    const sheet = parseSheetFromAddress(cell.range, context);
    const rangeRef = cell.range.includes("!") ? cell.range.split("!")[1] : cell.range;
    const range = sheet.getRange(rangeRef);
    range.values = cell.values;
    if (cell.numberFormats) {
      range.numberFormat = cell.numberFormats;
    }
  }

  await context.sync();
  return snapshot;
}

/**
 * Rollback a specific plan by ID.
 */
export async function rollbackPlan(
  context: Excel.RequestContext,
  planId: string
): Promise<boolean> {
  const idx = snapshotStack.findIndex((s) => s.planId === planId);
  if (idx === -1) return false;

  // Rollback all snapshots from this plan forward (LIFO order)
  const toRollback = snapshotStack.splice(idx);
  for (let i = toRollback.length - 1; i >= 0; i--) {
    const snapshot = toRollback[i];
    for (const cell of snapshot.cells) {
      const sheet = parseSheetFromAddress(cell.range, context);
      const rangeRef = cell.range.includes("!") ? cell.range.split("!")[1] : cell.range;
      const range = sheet.getRange(rangeRef);
      range.values = cell.values;
      if (cell.numberFormats) {
        range.numberFormat = cell.numberFormats;
      }
    }
  }

  await context.sync();
  return true;
}

/** Get the count of available snapshots */
export function getSnapshotCount(): number {
  return snapshotStack.length;
}

/** Get the most recent snapshot's plan ID */
export function getLastSnapshotPlanId(): string | undefined {
  return snapshotStack[snapshotStack.length - 1]?.planId;
}

/** Clear all snapshots */
export function clearSnapshots(): void {
  snapshotStack.length = 0;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function parseSheetFromAddress(
  address: string,
  context: Excel.RequestContext
): Excel.Worksheet {
  if (address.includes("!")) {
    let sheetName = address.split("!")[0];
    // Remove surrounding quotes if present (e.g. "'My Sheet'!A1")
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
      sheetName = sheetName.slice(1, -1);
    }
    return context.workbook.worksheets.getItem(sheetName);
  }
  return context.workbook.worksheets.getActiveWorksheet();
}

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
import { resolveRange } from "./capabilities/rangeUtils";

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
  // For each address, get only the used portion to avoid loading 1M rows for
  // full-column references like "Sheet1!A:A".
  // resolveRange handles workbook-qualified addresses correctly.
  const usedRanges: Excel.Range[] = [];

  for (const address of rangeAddresses) {
    try {
      const rng = resolveRange(context, address);
      // getUsedRange() (no args) includes formatting-only cells; safe on empty ranges
      const used = rng.getUsedRange(false);
      used.load(["values", "numberFormat", "address"]);
      usedRanges.push(used);
    } catch {
      // If the range can't be resolved, skip snapshotting it
    }
  }

  try {
    await context.sync();
  } catch {
    // If any range failed (e.g. empty sheet), clear it out so we still proceed
    usedRanges.length = 0;
  }

  const cells: CellSnapshot[] = usedRanges.map((range) => ({
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
  if (!address.includes("!")) {
    return context.workbook.worksheets.getActiveWorksheet();
  }
  const bangIdx = address.lastIndexOf("!");
  let sheetPart = address.substring(0, bangIdx);
  // Strip workbook qualifier: "[WorkbookName.xlsx]Sheet1" → "Sheet1"
  const wbMatch = sheetPart.match(/^\[.*?\](.+)$/);
  if (wbMatch) sheetPart = wbMatch[1];
  // Strip surrounding quotes: "'My Sheet'" → "My Sheet"
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
    sheetPart = sheetPart.slice(1, -1);
  }
  return context.workbook.worksheets.getItem(sheetPart);
}

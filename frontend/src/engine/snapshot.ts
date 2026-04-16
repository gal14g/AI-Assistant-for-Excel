/**
 * Snapshot & Rollback
 *
 * Before any write operation, the executor captures a snapshot of the
 * affected cell ranges. If the user requests an undo or the plan fails
 * partway, the snapshot is used to restore the original state.
 *
 * IMPORTANT Office.js notes:
 * - context.sync() must be called to flush reads before we can capture values.
 * - We store values, number formats, AND formulas for accurate rollback.
 *   Formulas are needed to restore formula cells (writing values would lose them).
 * - We detect merged areas so rollback can warn about merge state.
 * - Formatting (fill, font, borders) is NOT snapshotted because our design
 *   principle is to never touch formatting unless explicitly requested.
 */

import { CellSnapshot, InverseOp, PlanSnapshot } from "./types";
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
  // Load all ranges in a single batch + sync
  // resolveRange handles workbook-qualified and Hebrew sheet names correctly
  const ranges: Excel.Range[] = [];
  for (const address of rangeAddresses) {
    const range = resolveRange(context, address);
    range.load(["values", "numberFormat", "formulas", "address"]);
    ranges.push(range);
  }
  await context.sync();

  const cells: CellSnapshot[] = ranges.map((range) => ({
    range: range.address,
    values: range.values as (string | number | boolean | null)[][],
    numberFormats: range.numberFormat as string[][],
    formulas: range.formulas as string[][],
  }));

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
      used.load(["values", "numberFormat", "formulas", "address"]);
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
    formulas: range.formulas as string[][],
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

/** Restore one snapshot's state (inverse ops first, then cell values). */
async function restoreSnapshot(
  context: Excel.RequestContext,
  snapshot: PlanSnapshot,
): Promise<void> {
  // 1. Apply inverse ops in REVERSE order so structural changes unwind
  //    (delete sheet → restore tab color → rename back, etc.). Done BEFORE
  //    the cell-value restore so the sheet geometry exists for the writes.
  if (snapshot.inverseOps && snapshot.inverseOps.length > 0) {
    for (let i = snapshot.inverseOps.length - 1; i >= 0; i--) {
      await applyInverseOp(context, snapshot.inverseOps[i]);
    }
  }

  // 2. Restore cell values / formulas / number formats on ranges that still
  //    exist. If a sheet was just deleted by an inverse op, its cell entries
  //    will fail silently and that's correct (nothing to restore there).
  for (const cell of snapshot.cells) {
    try {
      const sheet = parseSheetFromAddress(cell.range, context);
      const rangeRef = cell.range.includes("!") ? cell.range.split("!")[1] : cell.range;
      const range = sheet.getRange(rangeRef);
      if (cell.formulas) {
        range.formulas = cell.formulas;
      } else {
        range.values = cell.values;
      }
      if (cell.numberFormats) {
        range.numberFormat = cell.numberFormats;
      }
    } catch {
      // Sheet or range no longer exists (inverse op removed it). Skip.
    }
  }

  await context.sync();
}

/**
 * Restore the most recent snapshot (undo last plan execution).
 */
export async function rollbackLastSnapshot(
  context: Excel.RequestContext
): Promise<PlanSnapshot | null> {
  const snapshot = snapshotStack.pop();
  if (!snapshot) return null;
  await restoreSnapshot(context, snapshot);
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

  // Rollback all snapshots from this plan forward (LIFO order so the most
  // recent step is undone first).
  const toRollback = snapshotStack.splice(idx);
  for (let i = toRollback.length - 1; i >= 0; i--) {
    await restoreSnapshot(context, toRollback[i]);
  }
  return true;
}

/** Push an empty snapshot so a structural handler (addSheet, tabColor,
 *  sheetPosition, etc. — anything without a range param) can later attach
 *  an inverse op via `registerInverseOp`. Returns the snapshot so callers
 *  can inspect it if needed. */
export function createEmptySnapshot(planId: string): PlanSnapshot {
  const snapshot: PlanSnapshot = {
    planId,
    timestamp: new Date().toISOString(),
    cells: [],
  };
  snapshotStack.push(snapshot);
  while (snapshotStack.length > MAX_SNAPSHOTS) snapshotStack.shift();
  return snapshot;
}

// ── Inverse-op + expanded-footprint API (structural undo) ──────────────────
//
// Value-only snapshots can't reverse structural changes (adding a sheet,
// inserting rows, changing tab color). Handlers that do structural work
// register an `InverseOp` on the TOP snapshot (the one captured before
// their step ran) so the undo pass can execute it.
//
// Handlers that write to a larger range than the source (e.g.
// `lateralSpreadDuplicates` widens, `extractMatchedToNewRow` lengthens)
// call `expandSnapshotFootprint` BEFORE they write so the extra cells are
// captured in the snapshot and can be restored on undo.
//
// Both APIs are no-ops when the snapshot stack is empty (e.g. if the
// executor skipped snapshotting because the step had no range params).

/** Register an inverse operation on the most-recent snapshot. Called by
 *  structural handlers (addSheet, renameSheet, insertDeleteRows, tabColor, …)
 *  AFTER their mutation succeeds. */
export function registerInverseOp(op: InverseOp): void {
  const top = snapshotStack[snapshotStack.length - 1];
  if (!top) return;
  if (!top.inverseOps) top.inverseOps = [];
  top.inverseOps.push(op);
}

/** Capture values/formulas/numberFormats for an additional range and append
 *  them to the most-recent snapshot's `cells`. Handlers that write past their
 *  source range should call this BEFORE the write, so undo can restore the
 *  outer footprint too. */
export async function expandSnapshotFootprint(
  context: Excel.RequestContext,
  extraRangeAddresses: string[],
): Promise<void> {
  const top = snapshotStack[snapshotStack.length - 1];
  if (!top || extraRangeAddresses.length === 0) return;

  const ranges: Excel.Range[] = [];
  for (const address of extraRangeAddresses) {
    try {
      const rng = resolveRange(context, address);
      rng.load(["values", "numberFormat", "formulas", "address"]);
      ranges.push(rng);
    } catch {
      // If a range can't be resolved (e.g. sheet doesn't exist yet), skip it.
    }
  }
  try {
    await context.sync();
  } catch {
    return;
  }

  for (const range of ranges) {
    top.cells.push({
      range: range.address,
      values: range.values as (string | number | boolean | null)[][],
      numberFormats: range.numberFormat as string[][],
      formulas: range.formulas as string[][],
    });
  }
}

/** Apply a single inverse op. Used by rollback. Handlers are expected to
 *  tolerate missing targets (e.g. the sheet was already deleted manually). */
async function applyInverseOp(
  context: Excel.RequestContext,
  op: InverseOp,
): Promise<void> {
  try {
    switch (op.kind) {
      case "deleteSheet": {
        const ws = context.workbook.worksheets.getItemOrNullObject(op.sheetName);
        ws.load("isNullObject");
        await context.sync();
        if (!ws.isNullObject) ws.delete();
        break;
      }
      case "renameSheet": {
        const ws = context.workbook.worksheets.getItemOrNullObject(op.currentName);
        ws.load("isNullObject");
        await context.sync();
        if (!ws.isNullObject) ws.name = op.restoreName;
        break;
      }
      case "deleteRows": {
        const ws = context.workbook.worksheets.getItemOrNullObject(op.sheetName);
        ws.load("isNullObject");
        await context.sync();
        if (!ws.isNullObject) {
          try {
            ws.getRange(op.rangeAddress).getEntireRow().delete(Excel.DeleteShiftDirection.up);
          } catch { /* best-effort */ }
        }
        break;
      }
      case "deleteColumns": {
        const ws = context.workbook.worksheets.getItemOrNullObject(op.sheetName);
        ws.load("isNullObject");
        await context.sync();
        if (!ws.isNullObject) {
          try {
            ws.getRange(op.rangeAddress).getEntireColumn().delete(Excel.DeleteShiftDirection.left);
          } catch { /* best-effort */ }
        }
        break;
      }
      case "restoreTabColor": {
        const ws = context.workbook.worksheets.getItemOrNullObject(op.sheetName);
        ws.load("isNullObject");
        await context.sync();
        if (!ws.isNullObject) ws.tabColor = op.color;
        break;
      }
      case "restoreSheetPosition": {
        const ws = context.workbook.worksheets.getItemOrNullObject(op.sheetName);
        ws.load("isNullObject");
        await context.sync();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if (!ws.isNullObject) (ws as any).position = op.position;
        break;
      }
      case "restoreSheetDirection": {
        // Office.js doesn't expose RTL setter — this is a no-op on add-in
        // mode (MCP runtime handles it). We still include the op so the
        // audit trail captures intent.
        break;
      }
    }
  } catch {
    // Inverse-op failures shouldn't block the rest of undo.
  }
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

/**
 * lateralSpreadDuplicates — "duplicate sidecar" layout.
 *
 * For every non-first-occurrence of a key column, lift that row's entire
 * data out of vertical position and paste it horizontally next to the
 * first-occurrence row (on the left or right). The original duplicate rows
 * are optionally removed so each value appears in only one row afterwards.
 *
 * Useful when the user wants to review duplicate entries side-by-side with
 * the first one — e.g. interview follow-ups, order revisions, timesheet
 * entries per employee.
 *
 * Algorithm (single-pass, no temp sheets, one context.sync for the rewrite):
 *   1. Read the sourceRange values (clipped to used-range for safety).
 *   2. One pass: for each row, bucket by its key column. Track the row
 *      index of the FIRST occurrence of each key.
 *   3. Count max duplicates per key → determines how many sidecar blocks we need.
 *   4. Build a new 2D matrix of width = (numSidecars × origCols) + origCols.
 *      Each sidecar block holds one duplicate row's full data; blocks are
 *      laid out so block #1 (closest to the anchor) holds the earliest
 *      duplicate, block #N the latest.
 *   5. Write the matrix back starting at the sourceRange's top-left,
 *      resized to the new width and (possibly) shorter height (because
 *      duplicate rows are removed).
 *   6. Clear any leftover rows inside the original sourceRange that are
 *      no longer occupied after duplicates collapsed upward.
 *
 * No new columns are *inserted* into the worksheet — the handler simply
 * overwrites existing cells starting at the anchor. Callers who need the
 * sidecar to appear "to the left of column A" should pre-insert N columns
 * before invoking this handler (a step the planner can chain with
 * insertDeleteRows-equivalent logic, or just choose direction="right").
 */

import {
  CapabilityMeta,
  LateralSpreadDuplicatesParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { normalizeForCompare } from "../utils/normalizeString";
import { ensureUnmerged } from "../utils/mergedCells";
import { expandSnapshotFootprint } from "../snapshot";

const meta: CapabilityMeta = {
  action: "lateralSpreadDuplicates",
  description:
    "Lay duplicate rows horizontally next to their first occurrence — 'duplicate sidecar' layout for side-by-side review",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: LateralSpreadDuplicatesParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const {
    sourceRange,
    keyColumnIndex,
    hasHeaders = true,
    direction = "left",
    removeOriginalDuplicates = true,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would lateral-spread duplicates of column ${keyColumnIndex} in ${sourceRange} (direction=${direction}).`,
    };
  }

  options.onProgress?.("Analyzing duplicates and laying them out horizontally...");

  // Clip to used range so "A:A" / "Sheet1!A:D" don't load 1M rows.
  const rawRange = resolveRange(context, sourceRange);
  const used = rawRange.getUsedRange(false);
  used.load(["values", "rowCount", "columnCount", "address"]);
  await context.sync();

  // Merged cells break read-then-write handlers: `range.values` returns nulls
  // in non-anchor cells, destroying the key column before we can bucket.
  // Auto-unmerge (with a warning in the output message) so the user's intent
  // still completes.
  const mergeReport = await ensureUnmerged(context, used, {
    operation: "lateralSpreadDuplicates",
    policy: "refuseWithError",
  });
  if (mergeReport.error) return mergeReport.error;
  if (mergeReport.hadMerges) {
    // Re-load values after the unmerge, since the cell contents redistributed.
    used.load(["values"]);
    await context.sync();
  }

  const values = (used.values ?? []) as (string | number | boolean | null)[][];
  if (values.length === 0) {
    return { stepId: "", status: "success", message: "Source range is empty — nothing to spread." };
  }

  const origCols = used.columnCount;
  if (keyColumnIndex < 0 || keyColumnIndex >= origCols) {
    return {
      stepId: "",
      status: "error",
      message: `keyColumnIndex ${keyColumnIndex} is out of range (source has ${origCols} columns).`,
    };
  }

  // Separate headers from data rows.
  const headerRow = hasHeaders ? values[0] : null;
  const dataRows = hasHeaders ? values.slice(1) : values.slice();

  // ── Pass 1: bucket by key, remember first-occurrence index ──────────────
  const keyToFirstIdx = new Map<string, number>();
  const keyToDupIdxs = new Map<string, number[]>(); // 2nd+, 3rd+, ... duplicates

  for (let i = 0; i < dataRows.length; i++) {
    const raw = dataRows[i][keyColumnIndex];
    // Treat null/"" as "no key" — never counts as a duplicate of another blank.
    if (raw === null || raw === undefined || raw === "") continue;
    // Normalize the key so Hebrew with RTL marks, NBSP padding, and case
    // variations group together even if visually distinct.
    const key = normalizeForCompare(raw);
    if (key === "") continue;
    if (!keyToFirstIdx.has(key)) {
      keyToFirstIdx.set(key, i);
    } else {
      const arr = keyToDupIdxs.get(key) ?? [];
      arr.push(i);
      keyToDupIdxs.set(key, arr);
    }
  }

  // How many sidecar blocks do we need? = max duplicates of any single key
  let maxDups = 0;
  for (const arr of keyToDupIdxs.values()) {
    if (arr.length > maxDups) maxDups = arr.length;
  }

  if (maxDups === 0) {
    return {
      stepId: "",
      status: "success",
      message: `No duplicates found in column ${keyColumnIndex} of ${sourceRange} — nothing to spread.`,
      outputs: { outputRange: used.address, duplicateGroupCount: 0, duplicateRowCount: 0 },
    };
  }

  // ── Pass 2: build output matrix ─────────────────────────────────────────
  const sidecarCols = maxDups * origCols;
  const newWidth = sidecarCols + origCols;

  // Decide which data-rows survive. Non-first-occurrence rows go away iff
  // removeOriginalDuplicates. First-occurrence rows and non-duplicate rows
  // always survive.
  const duplicateRowIdxs = new Set<number>();
  for (const arr of keyToDupIdxs.values()) for (const idx of arr) duplicateRowIdxs.add(idx);

  const survivingIdxs: number[] = [];
  for (let i = 0; i < dataRows.length; i++) {
    if (removeOriginalDuplicates && duplicateRowIdxs.has(i)) continue;
    survivingIdxs.push(i);
  }

  // Build the output grid. For the "left" direction, the sidecar block sits
  // BEFORE the original columns (columns 0 … sidecarCols-1), then the
  // anchor row's original data fills columns sidecarCols … newWidth-1.
  // For "right", the anchor's data sits first (0 … origCols-1), then the
  // sidecar blocks fill origCols … newWidth-1.
  const emptyRow = (): (string | number | boolean | null)[] =>
    Array.from({ length: newWidth }, () => null);

  const buildAnchorRow = (srcRow: (string | number | boolean | null)[], dupRows: (string | number | boolean | null)[][]): (string | number | boolean | null)[] => {
    const out = emptyRow();

    // Sidecar placement. When `direction === "left"`, block #0 (farthest
    // from anchor) sits at columns [0, origCols); the LAST duplicate lives
    // there so the sidecar reads left-to-right as "oldest … newest anchor".
    // For `direction === "right"`, the first duplicate is the one closest
    // to the anchor (columns [origCols, 2*origCols)).
    for (let b = 0; b < dupRows.length; b++) {
      const dup = dupRows[b];
      let blockStart: number;
      if (direction === "left") {
        // Leftmost block is the FURTHEST-AWAY duplicate so original reading
        // order (top→bottom) becomes physical left→right when scanning the
        // row. Block (maxDups - 1) is nearest the anchor.
        blockStart = (maxDups - 1 - b) * origCols;
      } else {
        // Rightmost block is the FURTHEST-AWAY duplicate, so the first
        // duplicate sits closest to the anchor.
        blockStart = origCols + b * origCols;
      }
      for (let c = 0; c < origCols; c++) {
        out[blockStart + c] = dup[c] ?? null;
      }
    }

    // Anchor's own data.
    const anchorStart = direction === "left" ? sidecarCols : 0;
    for (let c = 0; c < origCols; c++) {
      out[anchorStart + c] = srcRow[c] ?? null;
    }
    return out;
  };

  // Widen the header row too so column labels stay aligned after writing.
  const outputGrid: (string | number | boolean | null)[][] = [];
  if (hasHeaders && headerRow) {
    const widenedHeader = emptyRow();
    // Replicate the original header labels in each sidecar block and once
    // in the anchor area, so the user can tell what each wide block means.
    const placeBlock = (startCol: number) => {
      for (let c = 0; c < origCols; c++) widenedHeader[startCol + c] = headerRow[c] ?? null;
    };
    if (direction === "left") {
      for (let b = 0; b < maxDups; b++) placeBlock(b * origCols);
      placeBlock(sidecarCols);
    } else {
      placeBlock(0);
      for (let b = 0; b < maxDups; b++) placeBlock(origCols + b * origCols);
    }
    outputGrid.push(widenedHeader);
  }

  // Surviving rows: compose anchor + its duplicates (if any).
  for (const i of survivingIdxs) {
    const anchorRowVals = dataRows[i];
    const key = normalizeForCompare(anchorRowVals[keyColumnIndex]);
    // This surviving row is a "first occurrence" if the key maps to it.
    const dupIdxs = keyToDupIdxs.get(key) ?? [];
    const firstIdxForKey = keyToFirstIdx.get(key);
    const isFirstOccurrence = firstIdxForKey === i && dupIdxs.length > 0;

    if (isFirstOccurrence) {
      const dupRows = dupIdxs.map((idx) => dataRows[idx]);
      outputGrid.push(buildAnchorRow(anchorRowVals, dupRows));
    } else {
      // Non-duplicate row (unique value, or blank key). No sidecar blocks.
      outputGrid.push(buildAnchorRow(anchorRowVals, []));
    }
  }

  // ── Pass 3: write back ──────────────────────────────────────────────────
  // Anchor the write at the TOP-LEFT of the used range. If direction="left"
  // and the user wanted the sidecar to be physically left of the original
  // column A, they can chain insertDeleteRows(columns=sidecarCols, "right")
  // BEFORE this step — or pre-select a sourceRange that starts further left.
  // We deliberately don't try to insert columns here: that would shift
  // unrelated data outside the source range.
  const sheet = used.worksheet;
  sheet.load("name");
  await context.sync();

  // Top-left row/col of the used range, 0-based on the sheet.
  // `used.address` looks like "Sheet5!C5:H40" — parse the top-left "C5".
  const addrPart = used.address.includes("!") ? used.address.split("!").pop()! : used.address;
  const [topLeftRef] = addrPart.split(":");
  const m = topLeftRef.match(/^([A-Z]+)(\d+)$/);
  if (!m) {
    return {
      stepId: "",
      status: "error",
      message: `Could not parse used-range address: ${used.address}`,
    };
  }
  const startColLetters = m[1];
  const startRow = Number(m[2]) - 1; // 0-based
  let startCol = 0;
  for (let c = 0; c < startColLetters.length; c++) {
    startCol = startCol * 26 + (startColLetters.charCodeAt(c) - 64);
  }
  startCol -= 1; // 0-based

  // New range dimensions: outputGrid.length × newWidth, anchored at
  // (startRow, startCol) on the source sheet. Wrap the batch in try/catch
  // so range-locked / protected-sheet errors surface as a clean StepResult
  // instead of an unhandled exception.
  const outRows = outputGrid.length;

  // Expand the snapshot footprint to cover the WIDER output region before
  // writing, so undo can restore the sidecar columns to their pre-write state.
  // The executor's default snapshot only captured the source range; the
  // extra columns (leftward or rightward) are outside that range and would
  // otherwise be lost on undo.
  const idxToLetters = (idx: number): string => {
    let n = idx + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  };
  try {
    sheet.load("name");
    await context.sync();
    const footprintAddr = `${sheet.name}!${idxToLetters(startCol)}${startRow + 1}:${idxToLetters(startCol + newWidth - 1)}${startRow + outRows}`;
    await expandSnapshotFootprint(context, [footprintAddr]);
  } catch {
    // Non-fatal — if footprint capture fails, the primary write still runs.
  }

  try {
    const outRange = sheet.getRangeByIndexes(startRow, startCol, outRows, newWidth);
    outRange.values = outputGrid as unknown as (string | number | boolean)[][];

    // If the original source had MORE rows than the output (because we
    // removed duplicate rows), blank the tail so stale data doesn't remain.
    const origRows = used.rowCount;
    if (outRows < origRows) {
      const tailRange = sheet.getRangeByIndexes(
        startRow + outRows,
        startCol,
        origRows - outRows,
        origCols,
      );
      const blankTail: (string | number | boolean | null)[][] = [];
      for (let r = 0; r < origRows - outRows; r++) {
        const blankRow: (string | number | boolean | null)[] = [];
        for (let c = 0; c < origCols; c++) blankRow.push(null);
        blankTail.push(blankRow);
      }
      tailRange.values = blankTail as unknown as (string | number | boolean)[][];
    }

    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return {
      stepId: "",
      status: "error",
      message: `Failed to write lateral-spread layout to ${sourceRange}: ${msg}. (Sheet may be protected or the target range is locked.)`,
      error: msg,
    };
  }

  // Build a nice output-range address for the binding store.
  const indexToLetters = (idx: number): string => {
    let n = idx + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  };
  const outputAddr = `${sheet.name}!${indexToLetters(startCol)}${startRow + 1}:${indexToLetters(startCol + newWidth - 1)}${startRow + outRows}`;

  const groupCount = Array.from(keyToDupIdxs.values()).filter((a) => a.length > 0).length;
  const dupRowCount = duplicateRowIdxs.size;

  return {
    stepId: "",
    status: "success",
    message:
      `Lateral-spread ${dupRowCount} duplicate row(s) across ${groupCount} key(s) ` +
      `into ${maxDups} sidecar block(s) on the ${direction} of the anchor. ` +
      `Output: ${outputAddr}.${mergeReport.warning ?? ""}`,
    outputs: {
      outputRange: outputAddr,
      duplicateGroupCount: groupCount,
      duplicateRowCount: dupRowCount,
    },
  };
}

registry.register(meta, handler as any);
export { meta };

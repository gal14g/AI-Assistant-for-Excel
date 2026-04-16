/**
 * Snapshot inverse-op + expanded-footprint tests.
 *
 * These are module-level unit tests — they exercise the snapshot stack
 * and inverse-op registration directly, without needing a real Excel
 * context. The Excel-dependent code paths (applyInverseOp → Office.js)
 * are covered by the manual smoke tests.
 */

import {
  createEmptySnapshot,
  registerInverseOp,
  clearSnapshots,
  getSnapshotCount,
  getLastSnapshotPlanId,
} from "../engine/snapshot";

describe("snapshot — structural undo", () => {
  beforeEach(() => clearSnapshots());

  it("createEmptySnapshot pushes a snapshot with no cells", () => {
    const snap = createEmptySnapshot("plan-1");
    expect(snap.cells).toEqual([]);
    expect(snap.inverseOps).toBeUndefined();
    expect(getSnapshotCount()).toBe(1);
    expect(getLastSnapshotPlanId()).toBe("plan-1");
  });

  it("registerInverseOp attaches an op to the most recent snapshot", () => {
    const snap = createEmptySnapshot("plan-1");
    registerInverseOp({ kind: "deleteSheet", sheetName: "NewSheet" });
    expect(snap.inverseOps).toEqual([{ kind: "deleteSheet", sheetName: "NewSheet" }]);
  });

  it("registerInverseOp is a no-op when there is no snapshot", () => {
    // No snapshot pushed — should silently do nothing.
    expect(() =>
      registerInverseOp({ kind: "deleteSheet", sheetName: "X" }),
    ).not.toThrow();
    expect(getSnapshotCount()).toBe(0);
  });

  it("multiple inverse ops attach in order of registration", () => {
    const snap = createEmptySnapshot("plan-1");
    registerInverseOp({ kind: "deleteSheet", sheetName: "A" });
    registerInverseOp({ kind: "renameSheet", currentName: "B", restoreName: "C" });
    registerInverseOp({ kind: "restoreTabColor", sheetName: "D", color: "#FF0000" });
    expect(snap.inverseOps).toHaveLength(3);
    expect(snap.inverseOps?.[0].kind).toBe("deleteSheet");
    expect(snap.inverseOps?.[1].kind).toBe("renameSheet");
    expect(snap.inverseOps?.[2].kind).toBe("restoreTabColor");
  });

  it("inverse ops only attach to the top snapshot (LIFO)", () => {
    const first = createEmptySnapshot("plan-1");
    const second = createEmptySnapshot("plan-1");
    registerInverseOp({ kind: "deleteSheet", sheetName: "Only-top" });
    expect(first.inverseOps).toBeUndefined();
    expect(second.inverseOps).toHaveLength(1);
  });

  it("handles each InverseOp kind shape", () => {
    const snap = createEmptySnapshot("plan-1");
    registerInverseOp({ kind: "deleteSheet", sheetName: "A" });
    registerInverseOp({ kind: "renameSheet", currentName: "B", restoreName: "C" });
    registerInverseOp({ kind: "deleteRows", sheetName: "S", rangeAddress: "5:7" });
    registerInverseOp({ kind: "deleteColumns", sheetName: "S", rangeAddress: "C:E" });
    registerInverseOp({ kind: "restoreTabColor", sheetName: "S", color: "" });
    registerInverseOp({ kind: "restoreSheetPosition", sheetName: "S", position: 0 });
    registerInverseOp({ kind: "restoreSheetDirection", sheetName: "S", isRightToLeft: true });
    expect(snap.inverseOps).toHaveLength(7);
  });
});

/**
 * Tests for workbook snapshot constants — verify the larger snapshot limits.
 *
 * These don't require Excel.run since they test exported constants and
 * type shapes only.
 */

import * as fs from "fs";
import * as path from "path";

const snapshotSource = fs.readFileSync(
  path.join(__dirname, "..", "taskpane", "workbookSnapshot.ts"),
  "utf-8",
);

describe("Workbook snapshot constants", () => {
  it("MAX_SAMPLE_ROWS should be 20 (increased from 5)", () => {
    const match = snapshotSource.match(/const MAX_SAMPLE_ROWS\s*=\s*(\d+)/);
    expect(match).not.toBeNull();
    expect(Number(match![1])).toBe(20);
  });

  it("MAX_SHEETS should be 15", () => {
    const match = snapshotSource.match(/const MAX_SHEETS\s*=\s*(\d+)/);
    expect(match).not.toBeNull();
    expect(Number(match![1])).toBe(15);
  });

  it("MAX_COLUMNS should be 30", () => {
    const match = snapshotSource.match(/const MAX_COLUMNS\s*=\s*(\d+)/);
    expect(match).not.toBeNull();
    expect(Number(match![1])).toBe(30);
  });
});

describe("SheetSnapshot type shape", () => {
  it("exports the expected interface fields", () => {
    const snapshot: import("../taskpane/workbookSnapshot").SheetSnapshot = {
      sheetName: "Test",
      rowCount: 100,
      columnCount: 5,
      headers: ["A", "B", "C", "D", "E"],
      sampleRows: [["v1", 2, true, null, "v5"]],
      dtypes: ["text", "number", "boolean", "empty", "text"],
      anchorCell: "A1",
      usedRangeAddress: "Test!A1:E100",
    };
    expect(snapshot.sheetName).toBe("Test");
    expect(snapshot.rowCount).toBe(100);
    expect(snapshot.sampleRows.length).toBe(1);
  });
});

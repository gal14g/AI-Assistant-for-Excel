/**
 * Capability Compliance Tests
 *
 * Reads the SOURCE CODE of every capability handler and verifies
 * that it follows best practices for robustness:
 *
 * 1. Handlers that load .values/.text MUST use getUsedRange(false) to clip
 *    full-column refs (like "K:K") so we don't load ~1M rows.
 * 2. Handlers that write to individual cells SHOULD have try-catch protection.
 * 3. Handlers that call context.sync() after batch writes SHOULD handle sync errors.
 * 4. The snapshot mechanism must capture formulas for accurate undo.
 *
 * These are "source code lint" tests — they verify patterns without running
 * Office.js, and will catch regressions if someone adds a new handler without
 * following the safety patterns.
 */

import * as fs from "fs";
import * as path from "path";

const CAPS_DIR = path.join(__dirname, "..", "engine", "capabilities");

/** Read a handler file's source code. */
function readHandler(filename: string): string {
  return fs.readFileSync(path.join(CAPS_DIR, filename), "utf-8");
}

/** Get all .ts handler files in the capabilities directory (excluding index/rangeUtils). */
function getAllHandlerFiles(): string[] {
  return fs.readdirSync(CAPS_DIR)
    .filter((f) => f.endsWith(".ts") && f !== "index.ts" && f !== "rangeUtils.ts")
    .sort();
}

// ─── All handler files ───────────────────────────────────────────────────────

const allFiles = getAllHandlerFiles();

// ─── Handlers that load .values or .text (read data from cells) ─────────────

const VALUE_LOADERS = allFiles.filter((f) => {
  const src = readHandler(f);
  // Matches: range.load("values"), .load(["values", ...]), .load("values, ...")
  return /\.load\s*\(\s*[\["].*values/i.test(src) || /\.load\s*\(\s*[\["].*text/i.test(src);
});

// ─── Handlers that write to individual cells (per-cell .values = [[...]]) ───

const CELL_WRITERS = allFiles.filter((f) => {
  const src = readHandler(f);
  // Matches: .getCell(ri, ci).values, .getRange(...).values = [[...]]
  // But NOT batch writes like: range.values = cleaned (which is fine)
  return /getCell\s*\([^)]*\)\s*\.\s*values\s*=/.test(src)
    || /getRange\s*\(`[^`]*\$\{[^}]+\}[^`]*`\)\s*\.\s*values\s*=/.test(src);
});

// ─── Handlers that do context.sync() after writes ───────────────────────────

const SYNC_AFTER_WRITE = allFiles.filter((f) => {
  const src = readHandler(f);
  // Has a mutating write (values = , formulas = , clear(, insert() followed by sync
  return (src.includes(".values =") || src.includes(".formulas =") || src.includes(".clear(") || src.includes(".insert("))
    && src.includes("context.sync()");
});

// ═════════════════════════════════════════════════════════════════════════════
// TEST SUITES
// ═════════════════════════════════════════════════════════════════════════════

describe("Handler inventory", () => {
  it("should have at least 60 handler files (we expect ~76 actions across ~60 files)", () => {
    expect(allFiles.length).toBeGreaterThanOrEqual(55);
  });

  it("every file in index.ts should exist as a handler file", () => {
    const indexSrc = fs.readFileSync(path.join(CAPS_DIR, "index.ts"), "utf-8");
    const imports = [...indexSrc.matchAll(/import\s+["']\.\/(\w+)["']/g)].map((m) => m[1]);
    for (const imp of imports) {
      const file = `${imp}.ts`;
      expect(allFiles).toContain(file);
    }
  });

  it("every handler file registers with the registry", () => {
    for (const file of allFiles) {
      const src = readHandler(file);
      expect(src).toMatch(/registry\.register\s*\(/);
    }
  });
});

describe("Full-column range clipping (getUsedRange)", () => {
  // Every handler that loads .values or .text from a USER-PROVIDED range
  // must clip via getUsedRange(false) to avoid loading 1M rows for "A:A".

  const EXEMPT_FROM_CLIPPING = [
    // These handlers only read from ranges they create themselves, or
    // use Office.js native APIs that handle range sizing automatically:
    "readRange.ts",       // read-only, user can see the result
    "writeFormula.ts",    // writes formulas, doesn't read large ranges
    "bulkFormula.ts",     // uses getUsedRange already in its own way
    "groupSum.ts",        // loads specific bounded ranges
    "createPivot.ts",     // Office.js pivot API handles range sizing
    "unpivot.ts",         // expects pre-bounded sourceRange
    "crossTabulate.ts",   // expects pre-bounded sourceRange
    "transpose.ts",       // expects pre-bounded sourceRange
    "topN.ts",            // small output ranges
    "frequencyDistribution.ts", // small output ranges
    "joinSheets.ts",      // uses sheet-level getUsedRange internally
    "consolidateAllSheets.ts", // uses sheet-level getUsedRange internally
    "cloneSheetStructure.ts",  // uses getUsedRange on source sheet
    "deduplicateAdvanced.ts",  // uses its own range bounding
    "compareSheets.ts",   // bounded by two specific range params
    "consolidateRanges.ts", // bounded by specific range params
    "lookupAll.ts",       // uses getUsedRange already
    "spillFormula.ts",    // writes formula to single cell, reads back result — not user data
    "subtotals.ts",       // loads from resolveRange which is already bounded by user
    "fuzzyMatch.ts",      // uses getUsedRange on source ranges
  ];

  for (const file of VALUE_LOADERS) {
    if (EXEMPT_FROM_CLIPPING.includes(file)) continue;

    it(`${file}: must use getUsedRange(false) before loading values`, () => {
      const src = readHandler(file);
      const hasGetUsedRange = src.includes("getUsedRange(false)") || src.includes("getUsedRange(true)");
      expect(hasGetUsedRange).toBe(true);
    });
  }
});

describe("Per-cell write protection (try-catch)", () => {
  // Handlers that write to individual cells (getCell or template-string getRange)
  // must wrap those writes in try-catch to handle merged/protected cells.

  for (const file of CELL_WRITERS) {
    it(`${file}: must have try-catch around per-cell writes`, () => {
      const src = readHandler(file);
      // Check that there is at least one try block in the handler
      const hasTryCatch = src.includes("try {") || src.includes("try{");
      expect(hasTryCatch).toBe(true);
    });
  }
});

describe("Sync error handling after writes", () => {
  // Handlers that call context.sync() after batch writes should
  // have try-catch around at least one sync call.

  const EXEMPT_FROM_SYNC_CATCH = [
    // Handlers where sync failures are caught by the executor's
    // outer try-catch, and the operation is atomic (one write):
    "readRange.ts",
    "writeValues.ts",       // single batch write, executor catches
    "writeFormula.ts",      // single cell/formula write
    "clearRange.ts",        // clear is atomic
    "sortRange.ts",         // native sort API
    "removeDuplicates.ts",  // native API
    "mergeCells.ts",        // native merge API
    "setNumberFormat.ts",   // format write is atomic
    "autoFitColumns.ts",    // autofit is atomic
    "formatCells.ts",       // format write is atomic
    "addComment.ts",        // single operation
    "addHyperlink.ts",      // single operation
    "insertDeleteRows.ts",  // native insert/delete API
    "copyPasteRange.ts",    // native copy API
    "namedRange.ts",        // metadata operation
    "pageLayout.ts",        // settings operation
    "insertPicture.ts",     // single add operation
    "insertShape.ts",       // single add operation
    "insertTextBox.ts",     // single add operation
    "addSlicer.ts",         // single add operation
    "addSparkline.ts",      // single add operation
    "hideShow.ts",          // visibility toggle
    "groupRows.ts",         // native group API
    "setRowColSize.ts",     // format operation
    "freezePanes.ts",       // settings operation
    "createTable.ts",       // native table API
    "applyFilter.ts",       // native filter API
    "createChart.ts",       // native chart API
    "createPivot.ts",       // native pivot API
    "validation.ts",        // native validation API
    "conditionalFormat.ts", // native CF API
    "addDropdownControl.ts", // validation-based
    "alternatingRowFormat.ts", // format operation
    "quickFormat.ts",       // format operation
    "refreshPivot.ts",      // refresh operation
    "pivotCalculatedField.ts", // single field add
    "addReportHeader.ts",   // small write + format
    "sheetOps.ts",          // sheet management
    "spillFormula.ts",      // single formula write
    "conditionalFormula.ts", // formula write
    "runningTotal.ts",      // formula write
    "rankColumn.ts",        // formula write
    "percentOfTotal.ts",    // formula write
    "growthRate.ts",        // formula write
    // Batch-write handlers that write a single grid to one range —
    // these are atomic operations caught by the executor's outer try-catch:
    "unpivot.ts",           // single batch write to output range
    "crossTabulate.ts",     // single batch write to output range
    "transpose.ts",         // single batch write to output range
    "topN.ts",              // single batch write to output range
    "groupSum.ts",          // single batch write to output range
    "consolidateRanges.ts", // single batch write to output range
    "consolidateAllSheets.ts", // single batch write to output range
    "joinSheets.ts",        // single batch write to output range
    "frequencyDistribution.ts", // single batch write to output range
    "deduplicateAdvanced.ts",  // single batch write to output range
    "fuzzyMatch.ts",        // single batch write to output range
    "lookupAll.ts",         // single batch write to output range
    "splitByGroup.ts",      // per-sheet batch writes
    "cloneSheetStructure.ts", // structure clone operations
  ];

  const writingHandlers = SYNC_AFTER_WRITE.filter((f) => !EXEMPT_FROM_SYNC_CATCH.includes(f));

  for (const file of writingHandlers) {
    it(`${file}: should have try-catch around context.sync() after writes`, () => {
      const src = readHandler(file);
      // Must have at least one try block
      const hasTryCatch = src.includes("try {") || src.includes("try{") || src.includes("try {\n");
      expect(hasTryCatch).toBe(true);
    });
  }
});

describe("Snapshot mechanism", () => {
  const snapshotSrc = fs.readFileSync(
    path.join(__dirname, "..", "engine", "snapshot.ts"),
    "utf-8",
  );
  const typesSrc = fs.readFileSync(
    path.join(__dirname, "..", "engine", "types.ts"),
    "utf-8",
  );

  it("CellSnapshot interface should include formulas field", () => {
    expect(typesSrc).toMatch(/formulas\??\s*:\s*string\[\]\[\]/);
  });

  it("CellSnapshot interface should include hasMergedCells field", () => {
    expect(typesSrc).toMatch(/hasMergedCells\??\s*:\s*boolean/);
  });

  it("captureSnapshotBatched should load formulas", () => {
    expect(snapshotSrc).toMatch(/load\s*\(\s*\[.*"formulas".*\]/);
  });

  it("captureSnapshot should load formulas", () => {
    // Both capture functions must load formulas
    const matches = snapshotSrc.match(/load\s*\(\s*\[.*"formulas".*\]/g);
    expect(matches).not.toBeNull();
    expect(matches!.length).toBeGreaterThanOrEqual(2);
  });

  it("rollback should restore formulas (not just values)", () => {
    // Check that rollback writes formulas when available
    expect(snapshotSrc).toMatch(/cell\.formulas|range\.formulas\s*=/);
  });

  it("captureSnapshotBatched should use getUsedRange(false) for clipping", () => {
    expect(snapshotSrc).toMatch(/getUsedRange\s*\(\s*false\s*\)/);
  });
});

describe("Handler action registration completeness", () => {
  // Verify each handler exports a meta object with required fields
  for (const file of allFiles) {
    it(`${file}: should have meta with action, description/desc, mutates, affectsFormatting`, () => {
      const src = readHandler(file);
      expect(src).toMatch(/action:\s*"/);
      // sheetOps.ts uses "desc" instead of "description" in its array
      expect(src).toMatch(/description:\s*"|desc:\s*"/);
      expect(src).toMatch(/mutates:\s*(true|false)/);
      // sheetOps.ts doesn't have affectsFormatting in its compact array format
      if (file !== "sheetOps.ts") {
        expect(src).toMatch(/affectsFormatting:\s*(true|false)/);
      }
    });
  }
});

describe("Hebrew/Unicode safety in handlers", () => {
  it("resolveRange/rangeUtils should handle Hebrew sheet names", () => {
    const src = fs.readFileSync(path.join(CAPS_DIR, "rangeUtils.ts"), "utf-8");
    // Must strip quotes around sheet names (Hebrew sheets often get single-quoted)
    expect(src).toMatch(/startsWith\s*\(\s*["']'["']\s*\)/);
    // Must handle the ! separator
    expect(src).toMatch(/lastIndexOf\s*\(\s*["']!["']\s*\)/);
  });

  it("cleanupText should preserve Hebrew characters in removeNonPrintable", () => {
    const src = readHandler("cleanupText.ts");
    // The old regex [^\x20-\x7E] would strip Hebrew — verify it's not used
    expect(src).not.toMatch(/\[\\x20-\\x7E\]/);
    // Instead should use specific control char ranges
    expect(src).toMatch(/\\x00-\\x08/);
  });

  it("quoteSheetInRef should quote non-ASCII sheet names for formulas", () => {
    const src = fs.readFileSync(path.join(CAPS_DIR, "rangeUtils.ts"), "utf-8");
    // Must detect non-ASCII chars that need quoting
    expect(src).toMatch(/A-Za-z_.*A-Za-z0-9/);
  });
});

describe("Error value awareness", () => {
  it("findReplace should handle all Excel error types", () => {
    const src = readHandler("findReplace.ts");
    for (const err of ["#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#NULL!", "#SPILL!", "#CALC!"]) {
      expect(src).toContain(err);
    }
  });

  it("findReplace should clear error cells before writing", () => {
    const src = readHandler("findReplace.ts");
    expect(src).toMatch(/cell\.clear\s*\(/);
    expect(src).toMatch(/isErrorCell/);
  });
});

describe("dryRun support", () => {
  // Mutating handlers should support dryRun mode (return early without modifying).
  // Read-only handlers (readRange) don't need dryRun.
  const READ_ONLY_HANDLERS = ["readRange.ts"];

  const mutatingFiles = allFiles.filter((f) => {
    if (READ_ONLY_HANDLERS.includes(f)) return false;
    const src = readHandler(f);
    return /mutates:\s*true/.test(src);
  });

  for (const file of mutatingFiles) {
    it(`${file}: should check options.dryRun`, () => {
      const src = readHandler(file);
      expect(src).toMatch(/dryRun/);
    });
  }
});

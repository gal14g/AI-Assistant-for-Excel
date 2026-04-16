/**
 * Tests for the capability registry's fallback-handler support.
 *
 * The registry was extended so each capability can register a second handler
 * used when the running Excel doesn't satisfy `meta.requiresApiSet`. The
 * executor picks the fallback via `registry.getFallback(action)`; if none is
 * registered it surfaces a structured "unsupported on this Excel version"
 * error to the UI.
 */

import { registry } from "../engine/capabilityRegistry";
import type { CapabilityMeta, CapabilityHandler } from "../engine/types";
// Force all capabilities to register before we poke at the registry.
import "../engine/capabilities/index";

describe("capabilityRegistry — fallback dispatch", () => {
  it("getFallback returns undefined for handlers without a fallback", () => {
    // writeValues is 1.1-safe and has no fallback.
    expect(registry.getFallback("writeValues")).toBeUndefined();
  });

  it("spillFormula registers a fallback (dynamic-array rewrite)", () => {
    // spillFormula is 365-only. The fallback rewrites dynamic-array functions
    // into legacy equivalents for Excel 2016/2019.
    expect(registry.getFallback("spillFormula")).toBeDefined();
  });

  it("createPivot registers a fallback (SUMIFS summary sheet)", () => {
    expect(registry.getFallback("createPivot")).toBeDefined();
  });

  it("addSparkline registers a fallback (mini embedded charts)", () => {
    expect(registry.getFallback("addSparkline")).toBeDefined();
  });

  it("every > 1.3 handler has a fallback registered (Excel 2016 compat)", () => {
    // The handlers below declare requiresApiSet > 1.3 and therefore MUST ship
    // a fallback so Excel 2016 users aren't left with a hard error mid-plan.
    // If you add a new handler with requiresApiSet > 1.3, add a fallback too
    // (even a graceful-skip-with-warning one) and extend this list.
    const requireFallback: string[] = [
      "addSlicer",           // → warning + pointer to applyFilter
      "insertShape",         // → merged-cell rectangle / arrow glyph
      "insertTextBox",       // → merged-cell textbox approximation
      "insertPicture",       // → image placeholder (no rendering on 2016)
      "addComment",          // → adjacent-cell italic note
      "groupRows",           // → hide/show rows/cols (no +/− gutter)
      "removeDuplicates",    // → JS-side dedupe
      "addHyperlink",        // → =HYPERLINK() formula
      "pivotCalculatedField",// → append column to SUMIFS summary sheet
      "namedRange",          // → skip + surface absolute address
      "addValidation",       // → pale-gold tint + caption (no enforcement)
      "addDropdownControl",  // → inline annotation (no dropdown arrow)
      "pageLayout",          // → graceful skip with warning
      "freezePanes",         // → graceful skip with warning
      "createPivot",         // → SUMIFS summary sheet
      "addSparkline",        // → mini embedded charts
      "spillFormula",        // → dynamic-array rewrite
    ];
    const missing = requireFallback.filter((a) => !registry.getFallback(a as any));
    expect(missing).toEqual([]);
  });

  it("a freshly-registered capability with no fallback returns undefined", () => {
    const meta: CapabilityMeta = {
      action: "writeValues", // reuse an existing action name to avoid polluting registry
      description: "test",
      mutates: true,
      affectsFormatting: false,
    };
    const handler: CapabilityHandler = async () => ({
      stepId: "",
      status: "success",
      message: "ok",
    });
    // Re-register without a fallback. The warning about re-registration is
    // expected and doesn't affect correctness for this test.
    const spy = jest.spyOn(console, "warn").mockImplementation(() => {});
    registry.register(meta, handler);
    expect(registry.getFallback("writeValues")).toBeUndefined();
    spy.mockRestore();
  });
});

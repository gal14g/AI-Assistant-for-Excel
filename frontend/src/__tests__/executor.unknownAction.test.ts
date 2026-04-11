/**
 * Tests for buildUnknownActionError — the executor's fallback when a step
 * action isn't found in the capability registry.
 *
 * Reaching this branch means: the validator accepted the action (so its name
 * is in the schema) but no handler is wired up. That's almost always a
 * bundle/version mismatch where the backend planner emits an action the
 * frontend hasn't shipped a handler for yet. The error message should:
 *   1. Say so explicitly (so users don't think it's a broken plan)
 *   2. Suggest a reload (the actual fix)
 *   3. Offer fuzzy-match suggestions when there's a likely typo
 */

// Import capability index so the registry has its full action list before
// buildUnknownActionError calls registry.listActions().
import "../engine/capabilities/index";
import { buildUnknownActionError } from "../engine/executor";

describe("buildUnknownActionError", () => {
  it("includes the unknown action name", () => {
    const msg = buildUnknownActionError("totallyMadeUp");
    expect(msg).toContain("totallyMadeUp");
  });

  it("explains the likely cause (stale add-in bundle)", () => {
    const msg = buildUnknownActionError("totallyMadeUp");
    expect(msg).toMatch(/out of date|stale|bundle|reload/i);
  });

  it("suggests reloading the add-in", () => {
    const msg = buildUnknownActionError("totallyMadeUp");
    expect(msg).toMatch(/reload|Ctrl.*F5/i);
  });

  it("offers fuzzy-match suggestions for typos", () => {
    // "writeValue" (singular) is one char off from "writeValues"
    const msg = buildUnknownActionError("writeValue");
    expect(msg).toContain("writeValues");
  });

  it("suggests close matches for case-mangled action names", () => {
    // "createchart" (lowercase) should still match "createChart"
    const msg = buildUnknownActionError("createchart");
    expect(msg).toContain("createChart");
  });

  it("does not invent suggestions for wildly unrelated names", () => {
    // A nonsense string with no near neighbors should produce no "Did you mean"
    const msg = buildUnknownActionError("xyzqwerty12345");
    expect(msg).not.toMatch(/Did you mean/);
  });

  it("handles empty action gracefully", () => {
    const msg = buildUnknownActionError("");
    // Should still produce a meaningful error, not crash
    expect(msg.length).toBeGreaterThan(0);
    expect(msg).toMatch(/handler|action/i);
  });
});

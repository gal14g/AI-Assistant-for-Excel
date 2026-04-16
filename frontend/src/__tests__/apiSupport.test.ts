/**
 * Tests for the Excel API-support detection utility.
 *
 * These mock `Office.context.requirements.isSetSupported` to verify:
 *   1. `isSetSupported` returns true only when Office reports support
 *   2. Results are cached per (setName, version) pair
 *   3. `meetsRequirement` parses "ExcelApi 1.9" correctly
 *   4. An empty / missing requirement is treated as trivially satisfied
 *   5. `requireApiSet` throws a friendly error when the runtime is too old
 */

import {
  isSetSupported,
  meetsRequirement,
  parseApiRequirement,
  requireApiSet,
  _resetApiSupportCache,
} from "../engine/apiSupport";

describe("apiSupport", () => {
  beforeEach(() => {
    _resetApiSupportCache();
  });

  afterEach(() => {
    // Strip the mock so it doesn't leak across tests.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const g = globalThis as any;
    if (g.Office?.context) delete g.Office.context;
  });

  function installOfficeMock(impl: (set: string, ver: string) => boolean) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).Office = {
      ...(globalThis as any).Office,
      context: {
        requirements: {
          isSetSupported: jest.fn(impl),
        },
      },
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (globalThis as any).Office.context.requirements.isSetSupported as jest.Mock;
  }

  it("returns false when Office isn't present (e.g. unit test or browser mode)", () => {
    expect(isSetSupported("1.9")).toBe(false);
  });

  it("delegates to Office.context.requirements.isSetSupported when available", () => {
    installOfficeMock(() => true);
    expect(isSetSupported("1.9")).toBe(true);
  });

  it("caches results per (setName, version)", () => {
    const spy = installOfficeMock(() => true);
    isSetSupported("1.9");
    isSetSupported("1.9");
    isSetSupported("1.10");
    // Each unique version should trigger exactly one Office call.
    expect(spy).toHaveBeenCalledTimes(2);
  });

  it("returns false when Office throws (defensive catch)", () => {
    installOfficeMock(() => {
      throw new Error("Boom");
    });
    expect(isSetSupported("1.9")).toBe(false);
  });

  describe("parseApiRequirement", () => {
    it("parses 'ExcelApi 1.9' → { setName: 'ExcelApi', version: '1.9' }", () => {
      expect(parseApiRequirement("ExcelApi 1.9")).toEqual({
        setName: "ExcelApi",
        version: "1.9",
      });
    });

    it("accepts bare version strings and defaults set name to ExcelApi", () => {
      expect(parseApiRequirement("1.11")).toEqual({
        setName: "ExcelApi",
        version: "1.11",
      });
    });

    it("returns null for empty / undefined / null input", () => {
      expect(parseApiRequirement("")).toBeNull();
      expect(parseApiRequirement(undefined)).toBeNull();
      expect(parseApiRequirement(null)).toBeNull();
    });
  });

  describe("meetsRequirement", () => {
    it("returns true when the runtime reports support", () => {
      installOfficeMock((set, ver) => set === "ExcelApi" && ver === "1.9");
      expect(meetsRequirement("ExcelApi 1.9")).toBe(true);
    });

    it("returns false when the runtime doesn't support the version", () => {
      installOfficeMock((set, ver) => set === "ExcelApi" && ver === "1.3");
      expect(meetsRequirement("ExcelApi 1.11")).toBe(false);
    });

    it("returns true when no requirement is specified (trivially satisfied)", () => {
      // No Office mock installed — but there's no requirement so it doesn't matter.
      expect(meetsRequirement(undefined)).toBe(true);
      expect(meetsRequirement("")).toBe(true);
      expect(meetsRequirement(null)).toBe(true);
    });
  });

  describe("requireApiSet", () => {
    it("throws a friendly error naming the missing API set", () => {
      installOfficeMock(() => false);
      expect(() => requireApiSet("1.9", "sparklines")).toThrow(/ExcelApi 1\.9/);
      expect(() => requireApiSet("1.9", "sparklines")).toThrow(/sparklines/);
    });

    it("doesn't throw when the API set is supported", () => {
      installOfficeMock(() => true);
      expect(() => requireApiSet("1.9", "sparklines")).not.toThrow();
    });
  });
});

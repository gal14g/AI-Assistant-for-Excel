/**
 * Excel API-set detection helpers.
 *
 * Office.js ships different Excel capabilities in different "ExcelApi" set
 * versions:
 *   1.1 → Excel 2016 baseline
 *   1.3 → Excel 2016 (with batch sync improvements)
 *   1.7 → Excel 2019
 *   1.8 → PivotTables (fluent API), tables, named ranges, validation
 *   1.9 → Sparklines, pictures, shapes, textboxes, hyperlinks
 *  1.10 → Slicers, comments, cross-workbook range events
 *  1.11 → Dynamic arrays (SPILL behaviour)
 *  1.12+→ Excel 2021 and Microsoft 365
 *
 * We check support at *runtime* via Office.context.requirements.isSetSupported
 * so a single bundle can target multiple Excel versions and gracefully fall
 * back when a capability isn't available (e.g. Excel 2016 seeing a plan that
 * wants to build a PivotTable).
 */

/** Cache so we don't hit the Office.js bridge repeatedly. */
const _cache = new Map<string, boolean>();

/**
 * Returns true if the current Excel runtime supports the given API set at the
 * specified minimum version. Defaults to `ExcelApi` since that is the set we
 * care about almost everywhere in this codebase.
 *
 * If Office.js isn't available (e.g. in unit tests or browser mode), this
 * conservatively returns `false` so handlers fall back to their legacy paths.
 */
export function isSetSupported(version: string, setName: string = "ExcelApi"): boolean {
  const key = `${setName}@${version}`;
  const cached = _cache.get(key);
  if (cached !== undefined) return cached;

  let supported = false;
  try {
    // Guard: Office may be undefined in tests or non-Office contexts.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const g = globalThis as any;
    if (g?.Office?.context?.requirements?.isSetSupported) {
      supported = Boolean(g.Office.context.requirements.isSetSupported(setName, version));
    }
  } catch {
    supported = false;
  }

  _cache.set(key, supported);
  return supported;
}

/**
 * Reset the cache — test-only. Production code should treat API availability
 * as immutable for the lifetime of the process.
 */
export function _resetApiSupportCache(): void {
  _cache.clear();
}

/**
 * Parse a `CapabilityMeta.requiresApiSet` string like `"ExcelApi 1.9"` into
 * its parts. Accepts loose variants (`"1.9"` alone → {set:"ExcelApi", ver:"1.9"}).
 * Returns `null` when the input is empty / undefined (meaning "no requirement").
 */
export function parseApiRequirement(
  requirement: string | undefined | null,
): { setName: string; version: string } | null {
  if (!requirement) return null;
  const trimmed = requirement.trim();
  if (!trimmed) return null;

  // "ExcelApi 1.9" → ["ExcelApi", "1.9"]
  const parts = trimmed.split(/\s+/);
  if (parts.length === 1) {
    // Bare version string — default to ExcelApi.
    return { setName: "ExcelApi", version: parts[0] };
  }
  return { setName: parts[0], version: parts[1] };
}

/**
 * Convenience: given the `requiresApiSet` meta string (e.g. "ExcelApi 1.9"),
 * answer whether the current runtime satisfies it. Treats an empty/missing
 * requirement as trivially satisfied (baseline 1.1).
 */
export function meetsRequirement(requirement: string | undefined | null): boolean {
  const parsed = parseApiRequirement(requirement);
  if (!parsed) return true;
  return isSetSupported(parsed.version, parsed.setName);
}

/**
 * Throw a friendly error if the running Excel is older than `minVersion`.
 * Handlers can call this at the top to fail fast with an actionable message
 * when no sensible fallback exists for the legacy runtime (e.g. slicers).
 */
export function requireApiSet(minVersion: string, featureName?: string): void {
  if (!isSetSupported(minVersion)) {
    const feature = featureName ? ` (${featureName})` : "";
    throw new Error(
      `This action requires Excel with ExcelApi ${minVersion}+${feature}. ` +
      `Your Excel version is older — upgrade to Excel 2019, 2021, or Microsoft 365, ` +
      `or ask the assistant for an alternative approach.`,
    );
  }
}

/**
 * Shape of a reasonable "this feature is unavailable here" error that handlers
 * can return when there's no fallback. Keeps the shape consistent in the UI.
 */
export interface UnsupportedFeatureOptions {
  stepId?: string;
  feature: string;
  requiredVersion: string;
  suggestion?: string;
}

export function buildUnsupportedFeatureMessage(opts: UnsupportedFeatureOptions): string {
  const suggestion = opts.suggestion ? ` Suggestion: ${opts.suggestion}` : "";
  return (
    `${opts.feature} requires Excel with ExcelApi ${opts.requiredVersion}+ ` +
    `(Excel 2019/2021/Microsoft 365). Your Excel version is older.${suggestion}`
  );
}

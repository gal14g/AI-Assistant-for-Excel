/**
 * Merged-cell safety helpers.
 *
 * Handlers that read-then-write (sortRange, removeDuplicates, reorderRows,
 * lateralSpreadDuplicates, extractMatchedToNewRow, groupSum, matchRecords,
 * deduplicateAdvanced, fillBlanks, etc.) all break badly on merged ranges:
 *   - `range.values` returns nulls in non-anchor cells of a merge
 *   - Writing values to a merged range throws in some Office.js versions
 *   - `removeDuplicates` / `autoFilter.apply` error out entirely
 *
 * These helpers let handlers detect merges at the entry point and either
 *   (a) auto-unmerge and proceed (preferred — user's intent usually survives), or
 *   (b) surface a clear error so the LLM doesn't silently produce wrong results.
 *
 * Usage pattern inside a handler:
 *
 *   const mergeReport = await ensureUnmerged(context, rng, {
 *     operation: "sortRange",
 *     policy: "unmergeWithWarning",
 *   });
 *   if (mergeReport.error) return mergeReport.error;
 *   // ...proceed
 *   // Append mergeReport.warning to the result message if present.
 */

export interface MergeCheckOptions {
  /** Used only for message wording. */
  operation: string;
  /** How to react when merges are detected. */
  policy: "unmergeWithWarning" | "refuseWithError";
}

export interface MergeCheckResult {
  /** True if at least one merged area was found. */
  hadMerges: boolean;
  /** Count of merged areas that were unmerged. */
  unmergedCount: number;
  /** Human-readable suffix to append to the handler's success message, if any. */
  warning?: string;
  /** If policy="refuseWithError" and merges were found, a StepResult-shaped error. */
  error?: { stepId: ""; status: "error"; message: string };
}

/**
 * Detect merged areas inside `range`. On `policy="unmergeWithWarning"`,
 * unmerge them in place before returning. On `policy="refuseWithError"`,
 * return a structured error the handler can short-circuit with.
 *
 * Safe on single-cell or empty ranges. Failures (merge API unavailable on
 * very old ExcelApi sets) degrade to hadMerges=false.
 */
export async function ensureUnmerged(
  context: Excel.RequestContext,
  range: Excel.Range,
  options: MergeCheckOptions,
): Promise<MergeCheckResult> {
  try {
    // Range.isEntireRow / isEntireColumn short-circuit the merge check — full
    // rows/columns are rarely merged as a whole and the API may not support
    // getMergedAreasOrNullObject on those forms.
    range.load(["address", "rowCount", "columnCount"]);
    await context.sync();
    if (range.rowCount <= 1 && range.columnCount <= 1) {
      return { hadMerges: false, unmergedCount: 0 };
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mergedAreas = (range as any).getMergedAreasOrNullObject?.();
    if (!mergedAreas) {
      return { hadMerges: false, unmergedCount: 0 };
    }
    mergedAreas.load(["isNullObject", "address", "areaCount"]);
    await context.sync();

    if (mergedAreas.isNullObject) {
      return { hadMerges: false, unmergedCount: 0 };
    }

    const count = mergedAreas.areaCount ?? 1;

    if (options.policy === "refuseWithError") {
      return {
        hadMerges: true,
        unmergedCount: 0,
        error: {
          stepId: "",
          status: "error",
          message:
            `${options.operation} cannot run on ${range.address} — the range contains ` +
            `${count} merged area(s). Merged cells silently destroy row-axis data when read. ` +
            `Say "unmerge cells in ${range.address}" as a separate command first, then re-run ` +
            `the original request. (Home > Merge & Center also works if you prefer the UI.)`,
        },
      };
    }

    // Unmerge in place. Range.unmerge unmerges every merge inside the range.
    try {
      range.unmerge();
      await context.sync();
    } catch {
      // Best-effort — if the unmerge itself fails, report the issue so the
      // handler can decide whether to proceed.
      return {
        hadMerges: true,
        unmergedCount: 0,
        error: {
          stepId: "",
          status: "error",
          message:
            `${options.operation} could not unmerge ${range.address} automatically. ` +
            `Unmerge the cells manually (Home > Merge & Center) and retry.`,
        },
      };
    }

    return {
      hadMerges: true,
      unmergedCount: count,
      warning:
        ` (auto-unmerged ${count} merged area${count === 1 ? "" : "s"} first — merged ranges ` +
        `break most data-manipulation handlers; the content of non-anchor cells was lost by ` +
        `the original merge operation, not by this step.)`,
    };
  } catch {
    // If merge detection throws entirely (API variance across Excel versions),
    // degrade gracefully — the handler will proceed and either succeed or
    // surface its own error on the write.
    return { hadMerges: false, unmergedCount: 0 };
  }
}

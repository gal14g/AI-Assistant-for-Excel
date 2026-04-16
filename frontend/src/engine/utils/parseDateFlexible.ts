/**
 * parseDateFlexible — tolerant date parser.
 *
 * Accepts every shape that realistic user data can produce:
 *   - ISO 8601:   2026-03-15, 2026-03-15T14:30:00, 2026/03/15
 *   - dd/mm/yyyy: 15/03/2026 (default for most locales)
 *   - mm/dd/yyyy: 03/15/2026 (when localeHint='mdy')
 *   - dd-mm-yyyy: 15-03-2026
 *   - Excel serial date (number 1..2_958_466 — valid Excel date range)
 *   - JavaScript Date object
 *   - Month-name forms:  "15 Mar 2026", "March 15, 2026", "Mar 15 2026"
 *   - 2-digit years with reasonable cutoff (< 50 → 20xx, >= 50 → 19xx)
 *
 * Unambiguous dates (where one number exceeds 12) are parsed correctly
 * regardless of `localeHint`. Ambiguous dates (both ≤ 12) respect
 * `localeHint`; default is `"dmy"` since that is the majority-world format.
 *
 * Returns a UTC `Date` on success, `null` on failure.
 */

export type DateLocaleHint = "dmy" | "mdy" | "auto";

const MONTH_NAMES: Record<string, number> = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
};

function normalizeYear(y: number): number {
  if (y >= 100) return y;
  return y < 50 ? 2000 + y : 1900 + y;
}

/** Validate that (y, m, d) is a real calendar date (rejects Feb 30 etc.). */
function makeDateIfValid(y: number, m: number, d: number): Date | null {
  if (m < 0 || m > 11) return null;
  if (d < 1 || d > 31) return null;
  const date = new Date(Date.UTC(y, m, d));
  if (
    date.getUTCFullYear() !== y ||
    date.getUTCMonth() !== m ||
    date.getUTCDate() !== d
  ) {
    return null;
  }
  return date;
}

/** Excel serial → Date (days since 1899-12-30, accounting for Excel's 1900 leap-year bug). */
function excelSerialToDate(serial: number): Date | null {
  // Excel valid range is ~1900-01-01 (serial 1) through 9999-12-31 (~2_958_465).
  if (!Number.isFinite(serial) || serial < 1 || serial > 2_958_465) return null;
  const excelEpoch = Date.UTC(1899, 11, 30);
  const ms = excelEpoch + serial * 86_400_000;
  const d = new Date(ms);
  return Number.isNaN(d.getTime()) ? null : d;
}

export function parseDateFlexible(
  value: unknown,
  localeHint: DateLocaleHint = "dmy",
): Date | null {
  // Already a Date?
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  // Excel serial (number)?
  if (typeof value === "number") {
    return excelSerialToDate(value);
  }

  if (typeof value !== "string") return null;

  const s = value.trim();
  if (!s) return null;

  // ── ISO: yyyy-mm-dd or yyyy/mm/dd (optionally with Thh:mm:ss — we strip the time part) ──
  // The year-leading form is always unambiguous.
  {
    const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ].*)?$/);
    if (m) return makeDateIfValid(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  // ── Month-name forms ──
  // "15 Mar 2026", "15 March 2026", "March 15 2026", "March 15, 2026"
  {
    const dayFirst = s.match(/^(\d{1,2})[-/ ]+([A-Za-z]+)[-/ ,]+(\d{2,4})$/);
    if (dayFirst) {
      const mi = MONTH_NAMES[dayFirst[2].toLowerCase()];
      if (mi !== undefined) {
        return makeDateIfValid(normalizeYear(Number(dayFirst[3])), mi, Number(dayFirst[1]));
      }
    }
    const monthFirst = s.match(/^([A-Za-z]+)[-/ ]+(\d{1,2})[-/ ,]+(\d{2,4})$/);
    if (monthFirst) {
      const mi = MONTH_NAMES[monthFirst[1].toLowerCase()];
      if (mi !== undefined) {
        return makeDateIfValid(normalizeYear(Number(monthFirst[3])), mi, Number(monthFirst[2]));
      }
    }
  }

  // ── Numeric d-m-y or m-d-y with / or - or . separators ──
  {
    const m = s.match(/^(\d{1,4})[-/.](\d{1,2})[-/.](\d{1,4})$/);
    if (m) {
      const a = Number(m[1]);
      const b = Number(m[2]);
      const c = Number(m[3]);

      // Year-leading form like 2026-3-15 should already have matched above.
      // If we reach here with a large first number, treat it defensively as year.
      if (a >= 1000) {
        return makeDateIfValid(a, b - 1, c);
      }

      // Disambiguate: if one number is clearly a day-of-month (>12) we know
      // which slot is the month regardless of locale.
      let day: number, month: number, year: number;
      if (a > 12 && b <= 12) {
        // a is day, b is month → dd/mm/yyyy layout
        day = a; month = b; year = c;
      } else if (b > 12 && a <= 12) {
        // a is month, b is day → mm/dd/yyyy layout
        day = b; month = a; year = c;
      } else {
        // Ambiguous — use locale hint. mdy → a is month; dmy/auto → a is day.
        if (localeHint === "mdy") {
          month = a; day = b; year = c;
        } else {
          day = a; month = b; year = c;
        }
      }
      return makeDateIfValid(normalizeYear(year), month - 1, day);
    }
  }

  // ── Last-ditch: JS Date.parse for oddball formats ──
  // We only accept if the result is a finite Date, to avoid silently returning
  // "Invalid Date".
  const fallback = Date.parse(s);
  if (!Number.isNaN(fallback)) {
    const d = new Date(fallback);
    // Convert to UTC by extracting y/m/d in local time and rebuilding in UTC.
    // Date.parse() interprets yyyy-mm-dd as UTC but "Mar 15 2026" as local.
    // Normalize to a UTC calendar date (time zeroed).
    return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }

  return null;
}

/** Format a UTC Date back to `dd/mm/yyyy` — inverse convention for the dmy locale. */
export function formatDateDMY(d: Date): string {
  const day = String(d.getUTCDate()).padStart(2, "0");
  const mo = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yr = d.getUTCFullYear();
  return `${day}/${mo}/${yr}`;
}

/** Format a UTC Date back to `mm/dd/yyyy`. */
export function formatDateMDY(d: Date): string {
  const day = String(d.getUTCDate()).padStart(2, "0");
  const mo = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yr = d.getUTCFullYear();
  return `${mo}/${day}/${yr}`;
}

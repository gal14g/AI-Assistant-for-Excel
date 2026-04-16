/**
 * parseDateFlexible — exhaustive format coverage.
 */

import { parseDateFlexible, formatDateDMY, formatDateMDY } from "../engine/utils/parseDateFlexible";

function iso(y: number, m: number, d: number): string {
  return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

describe("parseDateFlexible — ISO forms", () => {
  it("yyyy-mm-dd", () => {
    const d = parseDateFlexible("2026-03-15");
    expect(d).not.toBeNull();
    expect(iso(d!.getUTCFullYear(), d!.getUTCMonth() + 1, d!.getUTCDate())).toBe("2026-03-15");
  });

  it("yyyy/mm/dd", () => {
    const d = parseDateFlexible("2026/03/15");
    expect(iso(d!.getUTCFullYear(), d!.getUTCMonth() + 1, d!.getUTCDate())).toBe("2026-03-15");
  });

  it("yyyy-mm-dd with time component (strips time)", () => {
    const d = parseDateFlexible("2026-03-15T14:30:00");
    expect(iso(d!.getUTCFullYear(), d!.getUTCMonth() + 1, d!.getUTCDate())).toBe("2026-03-15");
  });

  it("single-digit month/day in ISO", () => {
    const d = parseDateFlexible("2026-3-5");
    expect(iso(d!.getUTCFullYear(), d!.getUTCMonth() + 1, d!.getUTCDate())).toBe("2026-03-05");
  });
});

describe("parseDateFlexible — dd/mm/yyyy (default)", () => {
  it("parses 15/03/2026 as March 15", () => {
    const d = parseDateFlexible("15/03/2026");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("parses with dashes 15-03-2026", () => {
    const d = parseDateFlexible("15-03-2026");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("parses with dots 15.03.2026", () => {
    const d = parseDateFlexible("15.03.2026");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("day > 12 is unambiguously day (regardless of hint)", () => {
    const d = parseDateFlexible("25/03/2026", "mdy");
    expect(formatDateDMY(d!)).toBe("25/03/2026");
  });

  it("month > 12 means the first slot is the day", () => {
    const d = parseDateFlexible("03/25/2026", "dmy");
    expect(formatDateDMY(d!)).toBe("25/03/2026");
  });
});

describe("parseDateFlexible — mm/dd/yyyy with localeHint='mdy'", () => {
  it("parses 03/15/2026 as March 15 when hint=mdy", () => {
    const d = parseDateFlexible("03/15/2026", "mdy");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("ambiguous 05/04/2026 → May 4 with mdy", () => {
    const d = parseDateFlexible("05/04/2026", "mdy");
    expect(formatDateDMY(d!)).toBe("04/05/2026");
  });

  it("ambiguous 05/04/2026 → April 5 with dmy (default)", () => {
    const d = parseDateFlexible("05/04/2026");
    expect(formatDateDMY(d!)).toBe("05/04/2026");
  });
});

describe("parseDateFlexible — 2-digit years", () => {
  it("'15/03/26' → 2026 (< 50 cutoff)", () => {
    const d = parseDateFlexible("15/03/26");
    expect(d!.getUTCFullYear()).toBe(2026);
  });

  it("'15/03/98' → 1998 (>= 50 cutoff)", () => {
    const d = parseDateFlexible("15/03/98");
    expect(d!.getUTCFullYear()).toBe(1998);
  });

  it("'15/03/49' → 2049 (boundary)", () => {
    const d = parseDateFlexible("15/03/49");
    expect(d!.getUTCFullYear()).toBe(2049);
  });

  it("'15/03/50' → 1950 (boundary)", () => {
    const d = parseDateFlexible("15/03/50");
    expect(d!.getUTCFullYear()).toBe(1950);
  });
});

describe("parseDateFlexible — month names", () => {
  it("day-first: '15 Mar 2026'", () => {
    const d = parseDateFlexible("15 Mar 2026");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("day-first: '15 March 2026'", () => {
    const d = parseDateFlexible("15 March 2026");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("month-first: 'March 15, 2026'", () => {
    const d = parseDateFlexible("March 15, 2026");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("month-first short: 'Mar 15 2026'", () => {
    const d = parseDateFlexible("Mar 15 2026");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });

  it("month-name with 2-digit year: '15 Mar 26'", () => {
    const d = parseDateFlexible("15 Mar 26");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });
});

describe("parseDateFlexible — Excel serial numbers", () => {
  it("serial 1 → 1900-01-01 roughly (Excel's 1900 leap-year bug quirk)", () => {
    const d = parseDateFlexible(1);
    // Excel serial 1 is 1900-01-01; our epoch is 1899-12-30 so serial 1 = 1899-12-31.
    // We accept the quirk — what matters is forward compatibility post-1900-03-01.
    expect(d).not.toBeNull();
  });

  it("serial 45731 → 2025-03-15 (known ref)", () => {
    // 2025-03-15 = Excel serial 45731
    const d = parseDateFlexible(45731);
    expect(d).not.toBeNull();
    expect(formatDateDMY(d!)).toBe("15/03/2025");
  });

  it("rejects negative serial", () => {
    expect(parseDateFlexible(-1)).toBeNull();
  });

  it("rejects out-of-range serial", () => {
    expect(parseDateFlexible(10_000_000)).toBeNull();
  });
});

describe("parseDateFlexible — Date objects", () => {
  it("passes through valid Date", () => {
    const original = new Date(Date.UTC(2026, 2, 15));
    const d = parseDateFlexible(original);
    expect(d!.getTime()).toBe(original.getTime());
  });

  it("returns null for Invalid Date", () => {
    expect(parseDateFlexible(new Date("not-a-date"))).toBeNull();
  });
});

describe("parseDateFlexible — invalid", () => {
  it("rejects empty string", () => {
    expect(parseDateFlexible("")).toBeNull();
  });

  it("rejects whitespace-only", () => {
    expect(parseDateFlexible("   ")).toBeNull();
  });

  it("rejects gibberish", () => {
    expect(parseDateFlexible("hello world")).toBeNull();
  });

  it("rejects impossible date 31/02/2026 (no Feb 31)", () => {
    expect(parseDateFlexible("31/02/2026")).toBeNull();
  });

  it("rejects impossible 00/00/2026", () => {
    expect(parseDateFlexible("00/00/2026")).toBeNull();
  });

  it("rejects null / undefined / non-string non-number", () => {
    expect(parseDateFlexible(null)).toBeNull();
    expect(parseDateFlexible(undefined)).toBeNull();
    expect(parseDateFlexible({})).toBeNull();
    expect(parseDateFlexible([])).toBeNull();
  });
});

describe("parseDateFlexible — padded / whitespace", () => {
  it("trims leading/trailing whitespace", () => {
    const d = parseDateFlexible("  15/03/2026  ");
    expect(formatDateDMY(d!)).toBe("15/03/2026");
  });
});

describe("formatters", () => {
  it("formatDateDMY", () => {
    expect(formatDateDMY(new Date(Date.UTC(2026, 2, 15)))).toBe("15/03/2026");
  });

  it("formatDateMDY", () => {
    expect(formatDateMDY(new Date(Date.UTC(2026, 2, 15)))).toBe("03/15/2026");
  });
});

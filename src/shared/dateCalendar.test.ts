import {
  calendarGrid,
  displayIso,
  isoOf,
  isoToday,
  monthLabel,
  parseIso,
  shiftMonth,
  WEEKDAY_HEADERS,
} from "./dateCalendar";

describe("isoOf", () => {
  it("zero-pads month and day", () => {
    expect(isoOf(2026, 9, 3)).toBe("2026-09-03");
  });

  it("leaves a two-digit month/day alone", () => {
    expect(isoOf(2026, 12, 25)).toBe("2026-12-25");
  });
});

describe("parseIso", () => {
  it("parses a strict YYYY-MM-DD value", () => {
    expect(parseIso("2026-09-13")).toEqual({ year: 2026, month: 9, day: 13 });
  });

  it("trims surrounding whitespace", () => {
    expect(parseIso("  2026-09-13  ")).toEqual({ year: 2026, month: 9, day: 13 });
  });

  it("rejects a blank value", () => {
    expect(parseIso("")).toBeUndefined();
  });

  it("rejects a half-typed value", () => {
    expect(parseIso("2026-09-")).toBeUndefined();
  });

  it("rejects a month out of range", () => {
    expect(parseIso("2026-13-01")).toBeUndefined();
  });

  // The regression this exists for: `new Date(2026, 1, 30)` rolls FORWARD to March 2 instead of
  // throwing — silently accepting this would put a date on screen nobody typed.
  it("rejects a day that does not exist in that month", () => {
    expect(parseIso("2026-02-30")).toBeUndefined();
  });

  it("accepts a real leap-day date", () => {
    expect(parseIso("2028-02-29")).toEqual({ year: 2028, month: 2, day: 29 });
  });

  it("rejects a leap day in a non-leap year", () => {
    expect(parseIso("2026-02-29")).toBeUndefined();
  });
});

describe("isoToday", () => {
  it("formats the given Date in local time", () => {
    expect(isoToday(new Date(2026, 8, 3))).toBe("2026-09-03");
  });
});

describe("displayIso", () => {
  it("formats as DD/MMM/YYYY", () => {
    expect(displayIso("2026-09-13")).toBe("13/Sep/2026");
  });

  it("falls back to the raw value for something unparseable, never inventing a date", () => {
    expect(displayIso("not-a-date")).toBe("not-a-date");
  });
});

describe("monthLabel", () => {
  it("names the month and year", () => {
    expect(monthLabel(2026, 9)).toBe("September 2026");
  });
});

describe("shiftMonth", () => {
  it("steps forward within a year", () => {
    expect(shiftMonth(2026, 9, 1)).toEqual({ year: 2026, month: 10 });
  });

  it("steps backward within a year", () => {
    expect(shiftMonth(2026, 9, -1)).toEqual({ year: 2026, month: 8 });
  });

  it("wraps forward across a year boundary", () => {
    expect(shiftMonth(2026, 12, 1)).toEqual({ year: 2027, month: 1 });
  });

  it("wraps backward across a year boundary", () => {
    expect(shiftMonth(2026, 1, -1)).toEqual({ year: 2025, month: 12 });
  });

  it("wraps backward by more than one year's worth of months", () => {
    expect(shiftMonth(2026, 1, -13)).toEqual({ year: 2024, month: 12 });
  });
});

describe("calendarGrid", () => {
  it("is always exactly 42 cells, whatever the month's length", () => {
    // February 2026 (28 days) and October 2026 (31 days) — the panel must not change height.
    expect(calendarGrid(2026, 2)).toHaveLength(42);
    expect(calendarGrid(2026, 10)).toHaveLength(42);
  });

  it("marks exactly the real days of the month as inMonth", () => {
    const cells = calendarGrid(2026, 9); // September has 30 days
    expect(cells.filter((c) => c.inMonth)).toHaveLength(30);
  });

  it("starts the grid on a Sunday", () => {
    // September 1 2026 is a Tuesday, so the first two cells are the trailing days of August.
    const cells = calendarGrid(2026, 9);
    expect(cells[0].iso).toBe("2026-08-30");
    expect(cells[0].inMonth).toBe(false);
    expect(cells[2].iso).toBe("2026-09-01");
    expect(cells[2].inMonth).toBe(true);
  });

  it("fills the tail with the start of the next month", () => {
    const cells = calendarGrid(2026, 9);
    const last = cells[cells.length - 1];
    expect(last.inMonth).toBe(false);
    expect(last.iso.startsWith("2026-10")).toBe(true);
  });

  it("every cell parses back to a real calendar date", () => {
    for (const c of calendarGrid(2026, 2)) {
      expect(parseIso(c.iso)).toBeDefined();
    }
  });
});

describe("WEEKDAY_HEADERS", () => {
  it("has seven entries starting on Sunday", () => {
    expect(WEEKDAY_HEADERS).toEqual(["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]);
  });
});

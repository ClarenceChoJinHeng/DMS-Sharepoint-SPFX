/**
 * Tests for `dd-mm-yyyy` typed dates.
 *
 * The dangerous cases are the ones that produce a PLAUSIBLE wrong date rather than an error: a
 * day/month swap, a rolled-forward 31-02, a guessed century, and the UTC shift that makes an ISO date
 * render as the day before. Those get the coverage.
 */

import { autoFormatDmy, daysInMonth, isFuture, isoToDmy, parseDmy } from "./dateInput";

describe("parseDmy", () => {
  it("reads day, month, year in that order — never guessed", () => {
    // 02-08 is 2 August. Read the other way round it silently becomes 8 February.
    expect(parseDmy("02-08-2026")).toEqual({ state: "ok", iso: "2026-08-02" });
  });

  it("accepts single-digit day and month", () => {
    expect(parseDmy("2-8-2026")).toEqual({ state: "ok", iso: "2026-08-02" });
  });

  it("accepts slashes and dots, because people paste from Excel and email", () => {
    expect(parseDmy("02/08/2026")).toEqual({ state: "ok", iso: "2026-08-02" });
    expect(parseDmy("02.08.2026")).toEqual({ state: "ok", iso: "2026-08-02" });
  });

  it("treats blank as empty, not invalid — a field not yet filled is not an error", () => {
    expect(parseDmy("").state).toBe("empty");
    expect(parseDmy("   ").state).toBe("empty");
  });

  it("rejects a day the month does not have, and names the limit", () => {
    const r = parseDmy("31-02-2026");
    expect(r.state).toBe("invalid");
    if (r.state === "invalid") expect(r.reason).toContain("28 days");
  });

  it("rejects 31 April", () => {
    expect(parseDmy("31-04-2026").state).toBe("invalid");
  });

  it("accepts 29 February in a leap year and rejects it otherwise", () => {
    expect(parseDmy("29-02-2024")).toEqual({ state: "ok", iso: "2024-02-29" });
    expect(parseDmy("29-02-2026").state).toBe("invalid");
  });

  it("handles the century leap rules", () => {
    expect(parseDmy("29-02-2000")).toEqual({ state: "ok", iso: "2000-02-29" });
    expect(parseDmy("29-02-1900").state).toBe("invalid");
  });

  it("rejects month 00 and 13", () => {
    expect(parseDmy("01-00-2026").state).toBe("invalid");
    expect(parseDmy("01-13-2026").state).toBe("invalid");
  });

  it("rejects day 00", () => {
    expect(parseDmy("00-01-2026").state).toBe("invalid");
  });

  it("REJECTS a two-digit year rather than guessing the century", () => {
    // "02-08-26" could be 1926 or 2026. A guess here files a document a hundred years out.
    expect(parseDmy("02-08-26").state).toBe("invalid");
  });

  it("rejects text and partial input", () => {
    expect(parseDmy("2 August 2026").state).toBe("invalid");
    expect(parseDmy("02-08").state).toBe("invalid");
    expect(parseDmy("02-08-").state).toBe("invalid");
  });

  it("pads to a valid ISO string", () => {
    expect(parseDmy("5-1-2026")).toEqual({ state: "ok", iso: "2026-01-05" });
  });
});

describe("isoToDmy", () => {
  it("round-trips with parseDmy", () => {
    expect(isoToDmy("2026-08-02")).toBe("02-08-2026");
    expect(parseDmy(isoToDmy("2026-08-02"))).toEqual({ state: "ok", iso: "2026-08-02" });
  });

  it("returns empty for anything that is not a plain ISO date", () => {
    expect(isoToDmy("")).toBe("");
    expect(isoToDmy("2026-08-02T00:00:00Z")).toBe("");
    expect(isoToDmy("8/2/2026")).toBe("");
  });
});

describe("autoFormatDmy", () => {
  it("adds the dashes as digits are typed", () => {
    expect(autoFormatDmy("0", false)).toBe("0");
    expect(autoFormatDmy("02", false)).toBe("02-");
    expect(autoFormatDmy("02-0", false)).toBe("02-0");
    expect(autoFormatDmy("02-08", false)).toBe("02-08-");
    expect(autoFormatDmy("02-08-2026", false)).toBe("02-08-2026");
  });

  it("DOES NOT re-add a dash while deleting — otherwise backspace is impossible", () => {
    expect(autoFormatDmy("02-", true)).toBe("02-");
    expect(autoFormatDmy("02", true)).toBe("02");
  });

  it("drops anything that is not a digit or separator", () => {
    expect(autoFormatDmy("0a2", false)).toBe("02-");
  });

  it("normalises pasted slashes and dots", () => {
    expect(autoFormatDmy("02/08/2026", false)).toBe("02-08-2026");
    expect(autoFormatDmy("02.08.2026", false)).toBe("02-08-2026");
  });

  it("stops at eight digits, so extra keystrokes cannot extend the year", () => {
    expect(autoFormatDmy("020820261234", false)).toBe("02-08-2026");
  });

  it("returns empty for an emptied field", () => {
    expect(autoFormatDmy("", false)).toBe("");
  });
});

describe("isFuture", () => {
  const today = new Date(2026, 7, 15); // 15 Aug 2026, built locally — no UTC parsing anywhere.

  it("is false for today", () => {
    expect(isFuture("2026-08-15", today)).toBe(false);
  });

  it("is false for the past", () => {
    expect(isFuture("2026-08-14", today)).toBe(false);
    expect(isFuture("2025-12-31", today)).toBe(false);
  });

  it("is true for tomorrow and beyond", () => {
    expect(isFuture("2026-08-16", today)).toBe(true);
    expect(isFuture("2027-01-01", today)).toBe(true);
  });

  it("is false for a value it cannot read, rather than blocking on nonsense", () => {
    expect(isFuture("", today)).toBe(false);
    expect(isFuture("nonsense", today)).toBe(false);
  });

  it("does not shift a date across a timezone boundary", () => {
    // The bug this guards: new Date("2026-08-15") is UTC midnight, which is 14 Aug locally in the
    // Americas — so "today" would read as the future and a correct date would be refused.
    expect(isFuture("2026-08-15", new Date(2026, 7, 15, 23, 59, 59))).toBe(false);
    expect(isFuture("2026-08-15", new Date(2026, 7, 15, 0, 0, 0))).toBe(false);
  });
});

describe("daysInMonth", () => {
  it("knows the short months", () => {
    expect(daysInMonth(4, 2026)).toBe(30);
    expect(daysInMonth(6, 2026)).toBe(30);
    expect(daysInMonth(9, 2026)).toBe(30);
    expect(daysInMonth(11, 2026)).toBe(30);
  });

  it("knows the long months", () => {
    expect(daysInMonth(1, 2026)).toBe(31);
    expect(daysInMonth(12, 2026)).toBe(31);
  });

  it("knows February", () => {
    expect(daysInMonth(2, 2026)).toBe(28);
    expect(daysInMonth(2, 2024)).toBe(29);
  });
});

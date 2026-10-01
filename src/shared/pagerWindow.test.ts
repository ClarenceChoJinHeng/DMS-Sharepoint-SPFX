import { pagerWindow } from "./pagerWindow";

describe("pagerWindow", () => {
  it("shows everything when there are fewer pages than the window size", () => {
    expect(pagerWindow(1, 3, 7)).toEqual({ start: 0, end: 2 });
  });

  it("shows everything when the page count exactly matches the window size", () => {
    expect(pagerWindow(3, 7, 7)).toEqual({ start: 0, end: 6 });
  });

  it("centres the window on the current page, away from either edge", () => {
    // count=20, size=7, current=10 -> a symmetric window of 7 around it.
    expect(pagerWindow(10, 20, 7)).toEqual({ start: 7, end: 13 });
  });

  it("clamps to the START when the current page is near the beginning", () => {
    expect(pagerWindow(0, 20, 7)).toEqual({ start: 0, end: 6 });
  });

  it("clamps to the END when the current page is the last one reached", () => {
    // The exact shape reported live: pressing Next repeatedly leaves `current` at the newest page.
    expect(pagerWindow(19, 20, 7)).toEqual({ start: 13, end: 19 });
  });

  // The screenshot this was built to fix: 10 pages known, sitting on page index 8 (the 9th).
  it("regression: page 9 of 10 known, window 7 — clamped near the end, still exactly 7 wide", () => {
    const w = pagerWindow(8, 10, 7);
    expect(w).toEqual({ start: 3, end: 9 });
    expect(w.end - w.start + 1).toBe(7);
  });

  it("the window is always exactly `size` wide once count >= size, whatever the position", () => {
    for (let current = 0; current < 30; current++) {
      const w = pagerWindow(current, 30, 7);
      expect(w.end - w.start + 1).toBe(7);
    }
  });

  it("renders nothing for a zero page count", () => {
    expect(pagerWindow(0, 0, 7)).toEqual({ start: 0, end: -1 });
  });

  it("renders nothing for a zero window size", () => {
    expect(pagerWindow(0, 10, 0)).toEqual({ start: 0, end: -1 });
  });

  it("clamps a negative current to the first page", () => {
    expect(pagerWindow(-5, 20, 7)).toEqual(pagerWindow(0, 20, 7));
  });

  it("clamps a current beyond the known range to the last page", () => {
    expect(pagerWindow(500, 20, 7)).toEqual(pagerWindow(19, 20, 7));
  });

  it("never returns a start below zero", () => {
    for (const size of [1, 2, 3, 4, 7, 8, 20]) {
      for (let count = 0; count < 12; count++) {
        for (let current = -2; current < count + 2; current++) {
          expect(pagerWindow(current, count, size).start).toBeGreaterThanOrEqual(0);
        }
      }
    }
  });

  it("never returns an end beyond the last known page", () => {
    for (const size of [1, 2, 3, 4, 7, 8, 20]) {
      for (let count = 1; count < 12; count++) {
        for (let current = -2; current < count + 2; current++) {
          expect(pagerWindow(current, count, size).end).toBeLessThanOrEqual(count - 1);
        }
      }
    }
  });
});

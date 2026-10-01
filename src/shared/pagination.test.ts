import { paginate, pageWindow } from "./pagination";

describe("paginate", () => {
  const rows = [1, 2, 3, 4, 5, 6, 7];

  it("slices and reports the human range", () => {
    const p = paginate(rows, 0, 3);
    expect(p.slice).toEqual([1, 2, 3]);
    expect(p.from).toBe(1);
    expect(p.to).toBe(3);
    expect(p.total).toBe(7);
    expect(p.pageCount).toBe(3);
  });

  it("does not overrun the last page", () => {
    const p = paginate(rows, 2, 3);
    expect(p.slice).toEqual([7]);
    expect(p.to).toBe(7);
  });

  /* ⚠ THE CLAMP IS THE POINT. A filter that shortens the list leaves the caller's page past the end,
     and an unclamped slice returns [] - which renders as "nothing matches" over a list that has
     plenty. */
  it("clamps a page past the end rather than returning nothing", () => {
    const p = paginate(rows, 99, 3);
    expect(p.page).toBe(2);
    expect(p.slice).toEqual([7]);
  });

  it("clamps a negative page, and survives an empty list", () => {
    expect(paginate(rows, -5, 3).page).toBe(0);
    const empty = paginate([], 0, 3);
    expect(empty.pageCount).toBe(1);
    expect(empty.from).toBe(0);
    expect(empty.total).toBe(0);
  });
});

describe("pageWindow", () => {
  /* The client's screenshot: 1 2 3 ... 47 */
  it("keeps the first and last page, eliding the middle", () => {
    expect(pageWindow(0, 47)).toEqual([0, 1, undefined, 46]);
  });

  it("windows around the current page", () => {
    expect(pageWindow(20, 47)).toEqual([0, undefined, 19, 20, 21, undefined, 46]);
  });

  /* A gap of ONE is filled: "1 ... 3" is longer than "1 2 3". */
  it("fills a single-page gap rather than eliding it", () => {
    expect(pageWindow(2, 5)).toEqual([0, 1, 2, 3, 4]);
  });

  it("returns every page when they all fit", () => {
    expect(pageWindow(0, 3)).toEqual([0, 1, 2]);
    expect(pageWindow(0, 1)).toEqual([0]);
  });

  /* ⚠ NO DUPLICATES when the window touches an end - the first/last are already in the set. */
  it("never repeats a page number", () => {
    for (let count = 1; count <= 12; count++) {
      for (let page = 0; page < count; page++) {
        const nums = pageWindow(page, count).filter((n) => n !== undefined);
        expect(nums.length).toBe(new Set(nums).size);
      }
    }
  });
});

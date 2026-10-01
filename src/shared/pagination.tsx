// In-memory pagination, for lists that are already loaded.
//
// Client, 2026-09-04: *"I think its best to apply for pagination for My Submission, File Request, no
// need to use auto scroll."* Those lists were capped at 60vh and scrolled; on a site with real
// history that is a small window onto a long list, and a scroll position is not something anybody can
// return to or describe to a colleague.
//
// ⚠ THIS IS NOT THE AUDIT LOG'S PAGER, and must not be merged with it. That one pages on the SERVER
// through `readAudit` continuation tokens, because that list can hold tens of thousands of rows and a
// token belongs to the query that produced it. These lists are already fully in memory — My
// Submissions reads every library up front, Requests reads the whole request list — so paging them is
// slicing, and a server pager here would be a rewrite of two working reads for no gain.
//
// ⚠ THE RULE THESE LISTS INHERIT: PAGING MUST NEVER HIDE OUTSTANDING WORK WITHOUT SAYING SO. On the
// Requests page a filter may not hide a Pending row silently, and a page boundary is the same hazard —
// the difference is that a pager ALWAYS states the total ("Showing 1 to 20 of 63"), which a scroll
// container never did. That count is the safety, not decoration.
import * as React from "react";

/** The default page size. Twenty rows is roughly a screen without scrolling on a laptop. */
export const PAGE_SIZE = 20;

export interface Page<T> {
  /** The rows to render for the current page. */
  slice: T[];
  /** Total pages, never below 1 — an empty list still has a first page to be on. */
  pageCount: number;
  /** The page actually used, clamped into range. */
  page: number;
  /** 1-based index of the first row shown, or 0 when there are none. */
  from: number;
  /** 1-based index of the last row shown. */
  to: number;
  /** Total rows across every page. */
  total: number;
}

/**
 * Slice `rows` for `page`.
 *
 * ⚠ THE PAGE IS CLAMPED HERE, NOT BY THE CALLER. Deleting the last row of the last page, or typing a
 * filter that shortens the list, leaves the caller's `page` state pointing past the end — and an
 * unclamped slice returns `[]`, which renders as "nothing matches" over a list that has plenty. The
 * caller keeps a plain number; this decides what it means.
 */
export function paginate<T>(rows: readonly T[], page: number, size: number = PAGE_SIZE): Page<T> {
  const all = rows ?? [];
  const total = all.length;
  const per = size > 0 ? size : PAGE_SIZE;
  const pageCount = Math.max(1, Math.ceil(total / per));
  const clamped = Math.min(Math.max(0, Math.floor(page)), pageCount - 1);
  const start = clamped * per;
  return {
    slice: all.slice(start, start + per),
    pageCount,
    page: clamped,
    from: total === 0 ? 0 : start + 1,
    to: Math.min(start + per, total),
    total,
  };
}

const s: Record<string, React.CSSProperties> = {
  bar: {
    display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12,
    marginTop: 14, flexWrap: "wrap",
  },
  count: { fontSize: 12, color: "#605e5c" },
  btns: { display: "flex", alignItems: "center", gap: 4 },
  /* Taken from the Audit Log's own pager, which IS the design in the client's screenshot
     (2026-09-04) - same 30px square, same border, same green for the current page. Copied rather
     than imported because that one is bound to a token-based server pager in a web part; what is
     shared here is the look, and the two must not drift, so this note is the link between them. */
  btn: {
    minWidth: 30, height: 30, padding: "0 8px", fontSize: 12.5, border: "1px solid #d6d6d6",
    borderRadius: 4, background: "#fff", cursor: "pointer", color: "#1b1b1b",
    display: "inline-flex", alignItems: "center", justifyContent: "center",
  },
  btnOff: {
    minWidth: 30, height: 30, padding: "0 8px", fontSize: 12.5, border: "1px solid #ececec",
    borderRadius: 4, background: "#fff", cursor: "not-allowed", color: "#c4c4c4",
    display: "inline-flex", alignItems: "center", justifyContent: "center",
  },
  here: {
    minWidth: 30, height: 30, padding: "0 8px", fontSize: 12.5, border: "1px solid #0f6c3f",
    borderRadius: 4, background: "#fff", color: "#0f6c3f", fontWeight: 700,
    display: "inline-flex", alignItems: "center", justifyContent: "center",
  },
  gap: { minWidth: 20, textAlign: "center", fontSize: 12.5, color: "#8a8886" },
};

/**
 * Which page numbers to show, with `undefined` marking a gap.
 *
 * ⚠ FIRST AND LAST ARE ALWAYS PRESENT. The screenshot's `1 2 3 … 47` is not decoration: without the
 * last page there is no way to reach the end of a 47-page list except by pressing Next 46 times, and
 * a reader cannot see how much there is.
 *
 * Exported and pure so the windowing can be reasoned about on its own — it is the only part of this
 * file with a rule rather than a shape.
 */
export function pageWindow(page: number, pageCount: number, span: number = 1): (number | undefined)[] {
  const want: { [n: number]: true } = { 0: true, [pageCount - 1]: true };
  for (let i = page - span; i <= page + span; i++) {
    if (i >= 0 && i < pageCount) want[i] = true;
  }
  const nums = Object.keys(want).map(Number).sort((a, b) => a - b);
  const out: (number | undefined)[] = [];
  let prev = -1;
  for (const n of nums) {
    // A gap of exactly one page is filled rather than elided: "1 … 3" is longer than "1 2 3".
    if (prev >= 0 && n - prev === 2) out.push(prev + 1);
    else if (prev >= 0 && n - prev > 2) out.push(undefined);
    out.push(n);
    prev = n;
  }
  return out;
}

/**
 * Previous / page number / next, with the count beside it.
 *
 * ⚠ THE COUNT IS ALWAYS RENDERED, even on a single page. It is what tells a reader the list is
 * COMPLETE — the property a 60vh scroll box never had — and hiding it when the rows happen to fit is
 * how somebody comes to believe a short page is the whole list on the day it is not.
 *
 * Renders NOTHING when there is nothing at all: an empty list already has its own empty state, and a
 * pager reading "Showing 0 to 0 of 0" beneath it is noise.
 */
export function Pager({
  page,
  onPage,
  label = "items",
}: {
  page: Page<unknown>;
  onPage: (next: number) => void;
  /** What the rows ARE — "submissions", "requests". Plural; it only ever follows a number. */
  label?: string;
}): React.ReactElement | null {
  if (page.total === 0) return null;
  const first = page.page <= 0;
  const last = page.page >= page.pageCount - 1;
  return (
    <div style={s.bar}>
      <span style={s.count}>
        Showing {page.from} to {page.to} of {page.total} {label}
      </span>
      {page.pageCount > 1 && (
        <span style={s.btns}>
          <button
            type="button"
            style={first ? s.btnOff : s.btn}
            disabled={first}
            aria-label="Previous page"
            onClick={() => onPage(page.page - 1)}
          >
            &lsaquo;
          </button>
          {/* Numbered, with the current page boxed in green and gaps elided — the client's
              screenshot (2026-09-04). It replaced "1 of 47", which said where you were and gave no
              way to get anywhere except one step at a time. */}
          {pageWindow(page.page, page.pageCount).map((n, i) =>
            n === undefined ? (
              <span key={`gap${i}`} style={s.gap} aria-hidden="true">
                &hellip;
              </span>
            ) : (
              <button
                key={n}
                type="button"
                style={n === page.page ? s.here : s.btn}
                aria-current={n === page.page ? "page" : undefined}
                aria-label={`Page ${n + 1}`}
                onClick={() => onPage(n)}
              >
                {n + 1}
              </button>
            ),
          )}
          <button
            type="button"
            style={last ? s.btnOff : s.btn}
            disabled={last}
            aria-label="Next page"
            onClick={() => onPage(page.page + 1)}
          >
            &rsaquo;
          </button>
        </span>
      )}
    </div>
  );
}

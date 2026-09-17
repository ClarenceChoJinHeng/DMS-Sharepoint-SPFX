// A bounded, sliding window over pagination BUTTONS, for a pager whose pages are only reachable in
// order — through a server continuation cursor, never a jump to an unfetched page (see AuditLog's
// own note on this: "REST paging is a CONTINUATION TOKEN, so page 47 is only reachable by walking
// 1–46"). The set of pages the pager KNOWS how to jump straight back to only ever grows as more are
// visited, and rendering one button per known page, unbounded, is what floods the toolbar the moment
// anyone pages far enough — reported live, 2026-09-14: nine presses of Next produced ten buttons.
//
// ⚠ THIS IS NOT `shared/pagination.tsx`'s `paginate()`. That one SLICES an already-loaded array for
// an in-memory list (My Submissions, Requests); this only decides which of the ALREADY-KNOWN page
// NUMBERS get a button, and changes nothing about how the audit log fetches a page.

export interface PagerWindow {
  /** First index to render, inclusive. */
  start: number;
  /** Last index to render, inclusive. `end < start` means render nothing (an empty pager). */
  end: number;
}

/**
 * @param current the page on screen (0-based).
 * @param count   how many pages have been reached so far — `tokens.length` in the caller.
 * @param size    how many numbered buttons to show at once.
 *
 * Centres on `current` and clamps to `[0, count-1]` at either edge, so the window is always exactly
 * `size` wide once there are at least `size` known pages (never fewer buttons than fit just because
 * `current` sits near one end) and simply shows everything known when there are not.
 */
export function pagerWindow(current: number, count: number, size: number): PagerWindow {
  if (count <= 0 || size <= 0) return { start: 0, end: -1 };
  const last = count - 1;
  const clamped = Math.min(Math.max(0, current), last);

  let start = clamped - Math.floor((size - 1) / 2);
  let end = start + size - 1;

  if (start < 0) {
    end -= start; // shift the whole window right by the overhang
    start = 0;
  }
  if (end > last) {
    start -= end - last; // shift left
    end = last;
  }
  // A second clamp on `start`: the right-shift above can only ever push it further from 0 (still
  // safe), but the left-shift for the `end > last` case can push it negative when `count < size` —
  // already handled by taking everything in that case, this is the floor that guarantees it.
  start = Math.max(0, start);

  return { start, end };
}

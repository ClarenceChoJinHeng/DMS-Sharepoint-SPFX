import { CSSProperties } from "react";

/**
 * Shared responsive style fragments.
 *
 * Client, 2026-09-07: *"ensure all webpart is mobile friendly. Do not change the function, or copies
 * or anything just ensure its responsive."* So everything here is ADDITIVE CSS — no state, no
 * re-renders, no logic, no wording. Spreading one of these into an existing style object changes how
 * it behaves at narrow widths and nothing else.
 *
 * ⚠ WHY THERE ARE NO MEDIA QUERIES IN HERE, WHICH IS THE WHOLE CONSTRAINT.
 * 33 of the 35 web-part components style themselves with inline objects
 * (`Record<string, CSSProperties>`), and **inline styles cannot carry `@media`** — the same
 * limitation that forced `GroupManager`'s spinner to be an SVG with `animateTransform` rather than a
 * CSS animation. Only `Form.tsx` and `BulkUpload.tsx` render a `<style>` block and can take real
 * breakpoints.
 *
 * So the approach for everything else is CSS that is responsive WITHOUT a breakpoint: wrap instead
 * of overflow, `auto-fit` instead of a fixed column count, and a scroll container around the one
 * thing that genuinely cannot reflow — a table. That is strictly better than a breakpoint anyway,
 * because a SharePoint web part's width is set by the page section it sits in, not by the viewport:
 * a two-column section on a desktop can be narrower than a phone in landscape, and a media query
 * would get that exactly backwards.
 *
 * ⚠ AND THAT IS ALSO WHY `vw` UNITS ARE ABSENT. They measure the viewport, which is not the web
 * part's container. `100%` and `minmax(0, …)` measure what actually constrains the content.
 */

/**
 * Wrap a `<table>` in this. The table keeps its own width; the wrapper takes the scrolling.
 *
 * ⚠ THE SCROLL MUST BE ON A WRAPPER, NOT ON THE TABLE. Putting `display: block; overflow-x: auto`
 * on the `<table>` itself is the popular shortcut and it breaks column alignment: a block-level
 * table no longer establishes a table formatting context, so every row sizes its cells
 * independently and the columns stop lining up. The fault then looks like a data bug, not a CSS one.
 *
 * `WebkitOverflowScrolling` gives iOS momentum scrolling; without it a narrow table feels stuck.
 */
export const SCROLL_X: CSSProperties = {
  overflowX: "auto",
  WebkitOverflowScrolling: "touch",
  maxWidth: "100%",
};

/**
 * A minimum width for a table inside `SCROLL_X`, so columns are not crushed to one character.
 *
 * ⚠ IT IS A FLOOR, NOT A WIDTH. Paired with the table's existing `width: 100%` it means "fill the
 * container, but never go below this" — so nothing changes on a desktop and the wrapper scrolls on a
 * phone. Setting a `width` here instead would shrink wide tables on large screens.
 *
 * 560 rather than something larger: it has to fit a name, a badge and a date without scrolling on a
 * phone held in landscape, since those are the columns people actually read.
 */
export const TABLE_MIN: CSSProperties = { minWidth: 560 };

/**
 * A flex row that wraps instead of squashing.
 *
 * ⚠ THE COMMONEST FAILURE IN THIS CODEBASE, by a wide margin: 164 flex containers and only 55 of
 * them declare `flexWrap`. Without it a row of controls does not overflow visibly — it COMPRESSES,
 * so a button ends up two characters wide and a label becomes one letter per line. Nothing errors
 * and nothing scrolls; it just becomes unusable, which is why it survives review.
 */
export const WRAP_ROW: CSSProperties = { flexWrap: "wrap" };

/**
 * A grid that reflows from N columns to one as its container narrows.
 *
 * ⚠ `minmax(min(<px>, 100%), 1fr)` RATHER THAN `minmax(<px>, 1fr)`. A bare pixel floor is itself a
 * minimum width, so on a container narrower than that floor the grid overflows instead of
 * collapsing — the exact failure this is meant to remove. The inner `min()` caps the floor at the
 * container's own width.
 *
 * @param min the width below which a column drops to the next row.
 */
export function autoGrid(min: number, gap = 14): CSSProperties {
  return {
    display: "grid",
    gridTemplateColumns: `repeat(auto-fit, minmax(min(${min}px, 100%), 1fr))`,
    gap,
  };
}

/**
 * Stop a long unbroken string forcing its container wider than the screen.
 *
 * The strings that do it here are real and everywhere: server-relative paths, term-set GUIDs, group
 * names like `GHO_GCA_EG_UPLOADER_HIGHLY_CONFIDENTIAL`. `overflowWrap: anywhere` breaks them only
 * when there is no other option, so ordinary prose still wraps at spaces.
 */
export const BREAK_LONG: CSSProperties = {
  overflowWrap: "anywhere",
  wordBreak: "break-word",
};

/* ⚠ THERE IS DELIBERATELY NO `FIT` / PAGE-SHELL HELPER HERE, AND IT WAS TRIED AND REMOVED THE
   SAME DAY (2026-09-07). It was `{ maxWidth: "100%", boxSizing: "border-box" }`, on the theory
   that a pixel `maxWidth` on a page shell overflows a narrow container.

   THAT THEORY IS WRONG. `max-width` means "no wider than", so a shell capped at 1100px already
   fits a phone — there was nothing to fix. What the helper DID do was add
   `box-sizing: border-box`, which moves a shell's horizontal padding INSIDE the cap and so
   narrows its content column ON A DESKTOP: 1100 + 2x24 becomes 1100 total, 1052 of content.

   A responsive pass must not move a desktop layout. Deleted rather than parked, because an
   exported const nobody uses is one lint cannot flag and something eventually imports again —
   the same reasoning that fully removed `inviteToGroup` from `spGroups.ts` rather than
   leaving it dormant. */

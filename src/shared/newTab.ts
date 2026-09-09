import * as React from "react";

/**
 * The one definition of "this link opens in a new tab, and actually does".
 *
 * ⚠⚠ `target="_blank"` ALONE IS NOT ENOUGH INSIDE A SHAREPOINT MODERN PAGE, WHICH IS THE WHOLE
 * POINT OF THIS FILE. The page shell intercepts anchor clicks to route between modern pages without
 * a full reload (`_interceptAnchorClick` -> `_onIntercept` -> `push`) - the same soft navigation this
 * project already has written up as the reason a stale bundle fails on some navigations and not
 * others. A SAME-ORIGIN href is exactly what that router claims, and it can swallow the `target`
 * while doing so, so the link navigates in place.
 *
 * Reported live on the Term Store links (2026-09-09): *"I notice the Open the term stoer management
 * link is not opening a new tab"*. The markup was correct in the source AND in the shipped bundle -
 * every anchor carried `target="_blank" rel="noopener noreferrer"` - so the cause was never the
 * markup, and no amount of re-checking it would have found anything.
 *
 * ⚠ ON THOSE SCREENS THIS IS NOT COSMETIC. `StructureManager` holds an unsaved-changes guard, so
 * navigating in the SAME tab either prompts the admin or loses a half-filled Add form - the level
 * they were part-way through describing when they went to look up a term.
 *
 * `window.open` is not an anchor click, so nothing at the anchor level can intercept it. That is why
 * this is robust to being wrong about WHICH runtime mechanism is doing the swallowing.
 *
 * ⚠ IT READS THE URL OFF THE ANCHOR ITSELF, and takes no argument for it deliberately. The first
 * version was `openInNewTab(url)`, which put the address in the markup TWICE per link - so editing
 * the `href` and not the handler would have shown one page and opened another. `currentTarget.href`
 * cannot disagree with the link it is attached to.
 *
 * ⚠ THE ANCHOR THEREFORE MUST KEEP ITS `href`, and should keep `target`/`rel` too: they are what
 * make middle-click, Ctrl/Cmd-click and the context menu's own "Open link in new tab" behave
 * normally, and what happens if this handler never runs.
 */
export function openInNewTab(e: React.MouseEvent<HTMLAnchorElement>): void {
  /* Leave every MODIFIED click to the browser. Ctrl/Cmd/Shift-click and middle-click already mean
     "open this somewhere else", and handling them here would open two tabs or steal the window the
     admin asked for. */
  if (e.defaultPrevented) return;
  if (e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) return;

  // Resolved absolute form, whatever the href was written as.
  const url = e.currentTarget.href;
  if (!url) return;

  const opened = window.open(url, "_blank", "noopener,noreferrer");

  /* ⚠ ONLY SUPPRESS THE ANCHOR IF THE TAB ACTUALLY OPENED. A pop-up blocker answers `null`, and
     calling `preventDefault` regardless would turn a link that navigates in the wrong place into a
     link that does nothing at all - strictly worse, and the dead-control defect this project has
     now shipped three times. Falling through lets the anchor's own `target` have its go. */
  if (opened) e.preventDefault();
}

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
  // ⚠⚠⚠ TEMPORARY DIAGNOSTIC (2026-09-17), for the fourth live test of this same bug — REMOVE once
  // the mechanism is actually confirmed from a real console, not guessed at again. Three prior
  // "fixes" all reasoned about a mechanism nobody had actually observed running; this settles
  // whether the handler is invoked at all, before touching the mechanism a fourth time.
  console.info("[newTab-diag] handler invoked", { defaultPrevented: e.defaultPrevented, button: e.button });

  /* Leave every MODIFIED click to the browser. Ctrl/Cmd/Shift-click and middle-click already mean
     "open this somewhere else", and handling them here would open two tabs or steal the window the
     admin asked for. */
  if (e.defaultPrevented) return;
  if (e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) return;

  // Resolved absolute form, whatever the href was written as.
  const url = e.currentTarget.href;
  if (!url) {
    console.info("[newTab-diag] no href, bailing");
    return;
  }

  // ⚠⚠⚠ `stopPropagation` WAS NOT ENOUGH EITHER, AND THIS IS THE THIRD ATTEMPT AT THE SAME BUG
  // (2026-09-17). Live diagnosis this time: middle-click and the browser's own "Open link in new
  // tab" context-menu item — both of which bypass this handler entirely — worked correctly. A
  // plain click did not, threw no console error, and left the address bar showing the term store
  // URL while the PAGE CONTENT never actually swapped — i.e. SharePoint's router DID act on the
  // click, unchanged by anything `stopPropagation` did.
  //
  // The reason: `stopPropagation` only stops an event reaching OTHER NODES (ancestors, in bubble
  // phase). It does NOT stop OTHER LISTENERS ATTACHED TO THE SAME NODE from also firing — and
  // React's synthetic `stopPropagation` only calls the native event's `stopPropagation`, never
  // its `stopImmediatePropagation`. If SharePoint's interceptor is registered as its OWN native
  // listener directly on matching anchors (rather than only via delegation at some ancestor), a
  // sibling listener on that SAME node runs regardless of what an earlier listener on it already
  // called `stopPropagation()` — this is a genuine, well-known gap in React's event model, not a
  // guess about DOM structure. `stopImmediatePropagation()` on the underlying NATIVE event is the
  // one call that stops every other listener on this exact node, in either direction, whatever it
  // turns out to be attached to.
  //
  // Crucially, this does NOT touch the browser's own NATIVE handling of `target="_blank"` on the
  // anchor — that is governed by `preventDefault`, never by propagation — so the pop-up-blocked
  // case below still falls through to a real, unhijacked new-tab open via the browser itself, with
  // every other listener on this anchor simply never given the chance to run at all.
  e.stopPropagation();
  e.nativeEvent.stopImmediatePropagation();

  const opened = window.open(url, "_blank", "noopener,noreferrer");

  console.info("[newTab-diag] window.open result", { openedTruthy: !!opened, url });

  /* ⚠ ONLY SUPPRESS THE ANCHOR'S OWN DEFAULT IF THE TAB ACTUALLY OPENED VIA `window.open`. A
     pop-up blocker answers `null`, and calling `preventDefault` regardless would turn a link that
     navigates in the wrong place into a link that does nothing at all — strictly worse, and the
     dead-control defect this project has now shipped three times. Falling through lets the
     anchor's own `target` have its go, now that `stopPropagation` above has taken the one thing
     that was hijacking it out of the running. */
  if (opened) {
    e.preventDefault();
    console.info("[newTab-diag] preventDefault called");
  } else {
    console.info("[newTab-diag] NOT calling preventDefault — falling through to native target");
  }
}

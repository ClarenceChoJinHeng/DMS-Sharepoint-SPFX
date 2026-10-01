/**
 * Keeping a request list current without the reader pressing refresh.
 *
 * The problem, reported on site 2026-08-21: a requester watches My Submissions while an approver
 * decides in their own session, and nothing changes until the page is reloaded by hand. The same
 * applies in reverse — a Head of Unit's queue does not show a request that arrived a minute ago.
 *
 * ⚠ THERE IS NO PUSH CHANNEL. SPFx runs in the browser with no server component, and SharePoint
 * webhooks notify a SERVICE, not a page. So "live" here means two cheap things:
 *   - **Refresh when the tab regains focus.** This is the one that actually solves it: the reader
 *     switches to another account or tab, something happens, they switch back — and the list is
 *     current before they have looked at it. No timer, and nothing at all while the tab is idle.
 *   - **A slow poll while the tab is visible**, as the backstop for someone who simply sits on the
 *     page. Deliberately slow: this is a queue people check, not a chat.
 *
 * WHAT IT MUST NEVER DO is re-render the list under someone who is mid-decision. A refresh that
 * lands while a confirm dialog is open, or while a write is in flight, can replace the row being
 * acted on — so both callers pass `blocked` and the refresh is simply skipped. A skipped refresh
 * costs nothing: the next focus or tick picks it up, and the write path reloads on its own anyway.
 */

/** How often to re-read while the tab is visible. */
export const POLL_MS = 45000;

export interface RefreshGate {
  /** Is the tab actually being looked at? A hidden tab polls nothing. */
  visible: boolean;
  /**
   * Is something in progress that a re-render would disturb — an open confirm dialog, or a write
   * that has not come back yet? Blocking is always safe; the next tick catches up.
   */
  blocked: boolean;
}

/**
 * Should a refresh run right now?
 *
 * Pure so the rule can be tested without a DOM, and so both callers share ONE definition of "not a
 * good moment" rather than each remembering which flags to check.
 */
export function shouldRefresh(gate: RefreshGate): boolean {
  if (!gate || gate.blocked) return false;
  return gate.visible === true;
}

/* ── The hook ────────────────────────────────────────────────────────────────
   React is imported here rather than the effect being written out in both callers: the listener
   set, the cleanup and the ref dance below are easy to get subtly wrong, and two copies would drift
   in exactly the way that leaves one page silently not refreshing. */
import { useEffect, useRef } from "react";

/**
 * Re-run `reload` when the tab regains focus, and slowly while it is visible.
 *
 * `reload` is held in a REF and never in the effect's dependencies. It is redefined on every render
 * (it closes over state), so depending on it would tear down and rebuild the interval several times
 * a second — which in practice means the timer never fires at all. The ref keeps one interval for
 * the component's life while still calling the CURRENT reload.
 *
 * `blocked` is read through a ref for the same reason: it changes whenever a dialog opens, and the
 * timer must not restart every time.
 *
 * Errors from `reload` are swallowed. This is a background convenience — a failed poll must never
 * put an error in front of someone who did not ask for anything.
 */
export function useLiveRefresh(reload: () => void | Promise<void>, blocked: boolean): void {
  const reloadRef = useRef(reload);
  const blockedRef = useRef(blocked);
  reloadRef.current = reload;
  blockedRef.current = blocked;

  useEffect(() => {
    const run = (): void => {
      const visible = typeof document === "undefined" || document.visibilityState !== "hidden";
      if (!shouldRefresh({ visible, blocked: blockedRef.current })) return;
      try {
        const r = reloadRef.current();
        if (r && typeof (r as Promise<void>).catch === "function") {
          (r as Promise<void>).catch(() => undefined);
        }
      } catch {
        /* background refresh — never surfaced */
      }
    };

    // The one that matters: the reader comes back from the approver's tab and the list is already
    // current. `visibilitychange` covers a tab switch; `focus` covers a window switch.
    window.addEventListener("focus", run);
    document.addEventListener("visibilitychange", run);
    const timer = setInterval(run, POLL_MS);
    return () => {
      window.removeEventListener("focus", run);
      document.removeEventListener("visibilitychange", run);
      clearInterval(timer);
    };
    // Empty deps ON PURPOSE — see the refs above. This effect must run once per mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
}

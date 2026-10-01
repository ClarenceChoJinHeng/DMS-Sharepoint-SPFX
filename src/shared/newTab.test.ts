import * as React from "react";
import { openInNewTab } from "./newTab";

/**
 * The two rules here are the ones whose failure this project has actually shipped: a dead control,
 * and a handler that fights the browser's own gestures. Neither is visible by reading the JSX.
 */

type Fake = {
  ev: React.MouseEvent<HTMLAnchorElement>;
  prevented: () => boolean;
  propagationStopped: () => boolean;
  immediatePropagationStopped: () => boolean;
};

function click(over: Partial<Record<string, unknown>> = {}, href = "https://x/_layouts/15/termstoremanager.aspx"): Fake {
  let prevented = false;
  let propagationStopped = false;
  let immediatePropagationStopped = false;
  const ev = {
    button: 0,
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    altKey: false,
    defaultPrevented: false,
    currentTarget: { href },
    preventDefault: () => {
      prevented = true;
    },
    stopPropagation: () => {
      propagationStopped = true;
    },
    // `stopImmediatePropagation` lives on the underlying NATIVE event, not the synthetic one —
    // `e.nativeEvent.stopImmediatePropagation()` is what the implementation actually calls, so the
    // fake must carry a `nativeEvent` with its own stub rather than one on the top-level event.
    nativeEvent: {
      stopImmediatePropagation: () => {
        immediatePropagationStopped = true;
      },
    },
    ...over,
  } as unknown as React.MouseEvent<HTMLAnchorElement>;
  return {
    ev,
    prevented: () => prevented,
    propagationStopped: () => propagationStopped,
    immediatePropagationStopped: () => immediatePropagationStopped,
  };
}

describe("openInNewTab", () => {
  const realOpen = window.open;
  afterEach(() => {
    window.open = realOpen;
  });

  it("opens the anchor's own href in a new tab and suppresses the anchor", () => {
    const calls: Array<[string, string]> = [];
    window.open = ((u: string, t: string) => {
      calls.push([u, t]);
      return {} as Window;
    }) as typeof window.open;

    const c = click();
    openInNewTab(c.ev);

    // The URL is READ, never passed in — so it can never disagree with the link it sits on.
    expect(calls).toEqual([["https://x/_layouts/15/termstoremanager.aspx", "_blank"]]);
    expect(c.prevented()).toBe(true);
    // ⚠⚠ THE FIX FOR THE SECOND LIVE REPORT (2026-09-17): `preventDefault` alone was not enough —
    // SharePoint's own click interceptor sits ABOVE this component and still saw the click bubble
    // past it. `stopPropagation` is what actually keeps that listener from running at all.
    expect(c.propagationStopped()).toBe(true);
  });

  /* ⚠ THE ONE THAT MATTERS MOST. A pop-up blocker answers `null`; calling `preventDefault` anyway
     would turn a link that goes to the wrong place into a link that does nothing at all, which is
     strictly worse and is the dead-control defect this project has shipped three times. */
  it("leaves the anchor alone when the tab could NOT be opened, but still stops propagation", () => {
    window.open = (() => null) as typeof window.open;
    const c = click();
    openInNewTab(c.ev);
    expect(c.prevented()).toBe(false);
    // `stopPropagation` runs BEFORE `window.open` is even attempted, unconditionally — it is what
    // stops SharePoint's interceptor from hijacking the click; the anchor's own NATIVE
    // `target="_blank"` handling is governed by `preventDefault`, not propagation, so it is left
    // free to fire on its own once that interceptor can no longer see the click.
    expect(c.propagationStopped()).toBe(true);
  });

  /* Ctrl/Cmd/Shift/Alt-click and middle-click already mean "open this somewhere else". Handling them
     would open two tabs, or steal the window the admin actually asked for. */
  it("passes every modified click straight through to the browser", () => {
    for (const over of [
      { ctrlKey: true },
      { metaKey: true },
      { shiftKey: true },
      { altKey: true },
      { button: 1 },
    ]) {
      let opened = false;
      window.open = (() => {
        opened = true;
        return {} as Window;
      }) as typeof window.open;

      const c = click(over);
      openInNewTab(c.ev);
      expect(opened).toBe(false);
      expect(c.prevented()).toBe(false);
      // A modified click must be left ENTIRELY alone, including propagation — this handler has no
      // business touching it at all, not even to stop it reaching anything else.
      expect(c.propagationStopped()).toBe(false);
    }
  });

  it("does nothing to a click something else has already handled", () => {
    let opened = false;
    window.open = (() => {
      opened = true;
      return {} as Window;
    }) as typeof window.open;

    const c = click({ defaultPrevented: true });
    openInNewTab(c.ev);
    expect(opened).toBe(false);
    expect(c.propagationStopped()).toBe(false);
  });

  it("does not open a blank address", () => {
    let opened = false;
    window.open = (() => {
      opened = true;
      return {} as Window;
    }) as typeof window.open;

    const c = click({}, "");
    openInNewTab(c.ev);
    expect(opened).toBe(false);
    expect(c.prevented()).toBe(false);
    expect(c.propagationStopped()).toBe(false);
  });
});

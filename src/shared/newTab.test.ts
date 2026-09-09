import * as React from "react";
import { openInNewTab } from "./newTab";

/**
 * The two rules here are the ones whose failure this project has actually shipped: a dead control,
 * and a handler that fights the browser's own gestures. Neither is visible by reading the JSX.
 */

type Fake = {
  ev: React.MouseEvent<HTMLAnchorElement>;
  prevented: () => boolean;
};

function click(over: Partial<Record<string, unknown>> = {}, href = "https://x/_layouts/15/termstoremanager.aspx"): Fake {
  let prevented = false;
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
    ...over,
  } as unknown as React.MouseEvent<HTMLAnchorElement>;
  return { ev, prevented: () => prevented };
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
  });

  /* ⚠ THE ONE THAT MATTERS MOST. A pop-up blocker answers `null`; calling `preventDefault` anyway
     would turn a link that goes to the wrong place into a link that does nothing at all, which is
     strictly worse and is the dead-control defect this project has shipped three times. */
  it("leaves the anchor alone when the tab could NOT be opened", () => {
    window.open = (() => null) as typeof window.open;
    const c = click();
    openInNewTab(c.ev);
    expect(c.prevented()).toBe(false);
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
    }
  });

  it("does nothing to a click something else has already handled", () => {
    let opened = false;
    window.open = (() => {
      opened = true;
      return {} as Window;
    }) as typeof window.open;

    openInNewTab(click({ defaultPrevented: true }).ev);
    expect(opened).toBe(false);
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
  });
});

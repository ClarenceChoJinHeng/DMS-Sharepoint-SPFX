import { POLL_MS, shouldRefresh } from "./liveRefresh";

describe("shouldRefresh", () => {
  it("refreshes a visible, idle page", () => {
    expect(shouldRefresh({ visible: true, blocked: false })).toBe(true);
  });

  it("never refreshes a hidden tab — a background page has nobody reading it", () => {
    expect(shouldRefresh({ visible: false, blocked: false })).toBe(false);
  });

  it("NEVER refreshes while something is in progress, even when visible", () => {
    // The case this exists for: a confirm dialog open, or a decision write in flight. Replacing the
    // row under someone mid-decision is worse than being a few seconds stale.
    expect(shouldRefresh({ visible: true, blocked: true })).toBe(false);
    expect(shouldRefresh({ visible: false, blocked: true })).toBe(false);
  });

  it("treats a missing gate as 'not now' rather than throwing", () => {
    expect(shouldRefresh(undefined as unknown as { visible: boolean; blocked: boolean })).toBe(false);
  });

  it("polls slowly — this is a queue people check, not a chat", () => {
    // Pinned so nobody drops it to a few seconds without meeting this comment: every open tab pays
    // one list read per interval, for a list that changes a handful of times a day.
    expect(POLL_MS).toBeGreaterThanOrEqual(30000);
  });
});

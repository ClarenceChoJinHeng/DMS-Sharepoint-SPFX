import {
  buildQueue,
  nextUndecidedIndex,
  statusToDecision,
  swapState,
  QueueEntry,
  QueueItemLike,
} from "./approvalQueue";

const doc = (ID: number, status = 2): QueueItemLike => ({ ID, OData__ModerationStatus: status });
const PENDING = 2, APPROVED = 0, REJECTED = 1;

describe("statusToDecision", () => {
  it("maps SharePoint's moderation integers", () => {
    expect(statusToDecision(APPROVED)).toBe("Approved");
    expect(statusToDecision(REJECTED)).toBe("Rejected");
    expect(statusToDecision(PENDING)).toBe("Pending");
  });

  it("treats anything else as Pending rather than guessing", () => {
    // 3 is Draft. Reading it as Approved would show a decision that was never made.
    expect(statusToDecision(3)).toBe("Pending");
    expect(statusToDecision(-1)).toBe("Pending");
  });
});

describe("buildQueue", () => {
  it("finds the current document's position in the pending set", () => {
    const { queue, index } = buildQueue([doc(10), doc(11), doc(12)], doc(11));
    expect(queue.length).toBe(3);
    expect(index).toBe(1);
    expect(queue.every((e) => e.decided === null)).toBe(true);
  });

  it("keeps the server's order — oldest first is decided by $orderby, not here", () => {
    const { queue } = buildQueue([doc(10), doc(11), doc(12)], doc(10));
    expect(queue.map((e) => e.item.ID)).toEqual([10, 11, 12]);
  });

  // Opening an already-decided document by direct link is how a past decision gets reviewed.
  // It is not in a Pending query, so it is prepended rather than handled as a separate state.
  it("prepends an already-approved document at position 0 with its real decision", () => {
    const { queue, index } = buildQueue([doc(10), doc(11)], doc(99, APPROVED));
    expect(index).toBe(0);
    expect(queue.length).toBe(3);
    expect(queue[0].item.ID).toBe(99);
    expect(queue[0].decided).toBe("Approved");
  });

  it("prepends an already-rejected document the same way", () => {
    const { queue } = buildQueue([doc(10)], doc(99, REJECTED));
    expect(queue[0].decided).toBe("Rejected");
  });

  it("is safe with an empty or missing pending set", () => {
    const { queue, index } = buildQueue([], doc(7));
    expect(index).toBe(0);
    expect(queue.length).toBe(1);
    expect(buildQueue(undefined as unknown as QueueItemLike[], doc(7)).queue.length).toBe(1);
  });
});

describe("nextUndecidedIndex", () => {
  const q = (...decided: Array<"Approved" | "Rejected" | null>): Array<QueueEntry<QueueItemLike>> =>
    decided.map((d, i) => ({ item: doc(i + 1), decided: d }));

  it("moves forward to the next undecided entry", () => {
    expect(nextUndecidedIndex(q(null, null, null), 0)).toBe(1);
    expect(nextUndecidedIndex(q("Approved", null, null), 0)).toBe(1);
  });

  it("skips entries already decided", () => {
    expect(nextUndecidedIndex(q(null, "Approved", "Rejected", null), 0)).toBe(3);
  });

  // An approver who pressed Prev to re-check an earlier document would otherwise be told the
  // queue was finished while undecided work sat behind them.
  it("wraps backwards rather than reporting the queue finished", () => {
    expect(nextUndecidedIndex(q(null, "Approved", "Approved"), 2)).toBe(0);
  });

  it("returns -1 only when nothing anywhere is undecided", () => {
    expect(nextUndecidedIndex(q("Approved", "Approved"), 0)).toBe(-1);
    expect(nextUndecidedIndex(q("Approved"), 0)).toBe(-1);
    expect(nextUndecidedIndex([], 0)).toBe(-1);
  });

  it("never returns the position it was given", () => {
    expect(nextUndecidedIndex(q(null, "Approved"), 0)).toBe(-1);
  });
});

describe("swapState", () => {
  // THE data-integrity rule. Carrying a comment to the next document attaches one approver's
  // reasoning to a different document's audit trail.
  it("always clears the comment", () => {
    expect(swapState({ item: doc(1), decided: null }).comments).toBe("");
    expect(swapState({ item: doc(1, APPROVED), decided: "Approved" }).comments).toBe("");
  });

  it("clears the error and dismisses the success popup", () => {
    const st = swapState({ item: doc(1), decided: null });
    expect(st.submitError).toBe("");
    // null, not undefined: the component's `submitted` state is Decision | null and the popup
    // renders on any truthy value, so the reset has to produce the same shape it started with.
    expect(st.submitted).toBeNull();
    expect(st.fieldText).toEqual({});
  });

  it("takes the decision from the new document, not the previous one", () => {
    expect(swapState({ item: doc(1, PENDING), decided: null }).decision).toBe("Pending");
    expect(swapState({ item: doc(1, APPROVED), decided: null }).decision).toBe("Approved");
  });

  // The queue is never re-queried, so a locally decided entry's server status still reads
  // Pending. Its own flag has to win or the badge and the disabled Ok button disagree.
  it("prefers the local decision over a stale server status", () => {
    expect(swapState({ item: doc(1, PENDING), decided: "Approved" }).decision).toBe("Approved");
    expect(swapState({ item: doc(1, PENDING), decided: "Rejected" }).decision).toBe("Rejected");
  });
});

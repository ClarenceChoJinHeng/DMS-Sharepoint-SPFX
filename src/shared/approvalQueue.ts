// Approval queue logic — position, ordering, and what a swap resets.
//
// Pure and SPFx-free so the rules that are easy to get wrong are testable without a tenant:
// where the current document lands when it is not pending, which position "Approve another file"
// goes to, and — the one that is a data-integrity bug rather than a cosmetic one — that a
// comment never survives a move to another document.
//
// See docs/superpowers/specs/2026-07-31-approval-queue-navigation-design.md

export type Decision = "Approved" | "Rejected" | "Pending";

/** The only fields the queue logic needs. The component's richer item type satisfies it. */
export interface QueueItemLike {
  ID: number;
  OData__ModerationStatus: number;
}

/**
 * One position in the queue.
 *
 * `decided` is LOCAL. The queue is captured once and never re-queried, so the server copy of a
 * decided item still reads Pending. This flag is what keeps a decided position navigable — an
 * approver can step back to confirm what they did — while blocking a second submit.
 */
export interface QueueEntry<T extends QueueItemLike> {
  item: T;
  decided: Decision | null;
}

/** SharePoint's moderation status as a decision. 0 Approved, 1 Rejected, 2 Pending. */
export function statusToDecision(status: number): Decision {
  if (status === 0) return "Approved";
  if (status === 1) return "Rejected";
  return "Pending";
}

/**
 * Build the queue and the starting position.
 *
 * When the current document is NOT in the pending set — which is how an already-decided document
 * gets reviewed by direct link — it is prepended at position 0 rather than handled as a separate
 * state. The counter then reads "1 of 13", Prev is simply disabled, and no navigation control
 * needs a special case for "you are not in the queue".
 *
 * A prepended item carries its real recorded decision, so the badge and the disabled Ok button
 * behave the same as for one decided in this session.
 */
export function buildQueue<T extends QueueItemLike>(
  pending: T[],
  current: T,
): { queue: Array<QueueEntry<T>>; index: number } {
  const queue: Array<QueueEntry<T>> = (pending ?? []).map((item) => ({ item, decided: null }));
  const found = queue.findIndex((e) => e.item.ID === current.ID);
  if (found !== -1) return { queue, index: found };
  queue.unshift({ item: current, decided: statusToDecision(current.OData__ModerationStatus) });
  return { queue, index: 0 };
}

/**
 * The position "Approve another file" should go to: the next entry with no decision.
 *
 * Searches forward from `pos` first, then wraps. The wrap matters because an approver who used
 * Prev to check an earlier document would otherwise be told the queue was finished while
 * undecided work sat behind them.
 *
 * Returns -1 when nothing is left, which is what suppresses the button entirely — offering a
 * Next that leads nowhere is worse than not offering one.
 */
export function nextUndecidedIndex<T extends QueueItemLike>(
  queue: Array<QueueEntry<T>>,
  pos: number,
): number {
  for (let i = pos + 1; i < queue.length; i++) if (queue[i].decided === null) return i;
  for (let i = 0; i < queue.length; i++) if (i !== pos && queue[i].decided === null) return i;
  return -1;
}

/** Per-document state, as it must be after moving to another document. */
export interface SwapState {
  decision: Decision;
  comments: string;
  submitError: string;
  submitted: Decision | null;
  fieldText: Record<string, string>;
}

/**
 * What every navigation must reset.
 *
 * Extracted so `comments: ""` is asserted by a test rather than trusted to a list of setter
 * calls. Carrying a comment onto the next document would attach one approver's reasoning to a
 * different document's audit trail — the only item here that is a data-integrity bug rather
 * than a cosmetic one.
 *
 * A locally decided entry's own flag wins over its server status, which is still Pending because
 * the queue is never re-queried.
 */
export function swapState<T extends QueueItemLike>(entry: QueueEntry<T>): SwapState {
  return {
    decision: entry.decided ?? statusToDecision(entry.item.OData__ModerationStatus),
    comments: "",
    submitError: "",
    submitted: null,
    fieldText: {},
  };
}

/**
 * The decision a `?decision=` link asks the page to pre-select.
 *
 * Crystal's email templates 2 and 3 (spec `2026-08-28-email-bundling-and-templates-design.md` §5)
 * carry `[Approve]` and `[Reject]` links. A link that ACTUALLY approves needs a bearer trigger URL
 * inside the message — rejected for HC, and worse on the reminder, which goes to up to ten
 * mailboxes. So the links open this page with the radio already ticked and the approver still
 * presses the button.
 *
 * ⚠ PRE-SELECTING IS NOT DECIDING, and nothing here may ever become an auto-submit. Everything the
 * page does at the moment of submitting — the destination-folder guard, the name-clash check, the
 * `ApproveItems` probe — runs on that button press. A link that skipped it would skip all three.
 *
 * ⚠ ANYTHING UNRECOGNISED IS `undefined`, never a guess. A mangled parameter must leave the page
 * exactly as it behaves with no parameter at all: nothing ticked, the button disabled until the
 * approver chooses. Guessing would put a decision in front of them that nobody asked for.
 *
 * `Pending` is deliberately unreachable from a link — it is a system state, not something an
 * approver can pick, and the submit button is disabled while it is selected.
 */
/* `URLSearchParams.get` answers `null`, converted to `undefined` at the call site rather than
   widened here — the house lint rule bans `null` in new signatures. The `?? ""` below still
   absorbs one at runtime, because a defensive parameter reader should not depend on that. */
export function decisionFromLink(raw: string | undefined): Decision | undefined {
  const v = (raw ?? "").trim().toLowerCase();
  if (v === "approve" || v === "approved") return "Approved";
  if (v === "reject" || v === "rejected") return "Rejected";
  return undefined;
}

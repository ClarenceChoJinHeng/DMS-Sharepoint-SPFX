# Approval Queue Navigation — Design

**Date:** 2026-07-31
**Status:** Agreed, implementing
**Touches:** `ApprovalDocument.tsx`
**Amends:** `2026-07-27-upload-approval-popup-polish.md` (the success popup's button row)

---

## Problem

The Approval Document web part reviews exactly one file. It reads `?itemId=` from the query
string, loads that item, and every exit — the back link, `Cancel`, and the success popup's
button — navigates back to the Staging library.

So approving five documents means five round trips through the library grid: open a file,
decide, dismiss the popup, land in Staging, find the next file, open it. The library grid is
the only way to move between documents, and it is not a review queue — it is a file list that
happens to contain the pending items somewhere among the approved ones.

The approver's actual job is "clear what is waiting for me". Nothing in the current UI knows
that a queue exists, so nothing can tell them how many are left or take them to the next one.

## Shape of the change

Give the web part a **queue** — the pending documents in the approver's scope — and let it
move within that queue without leaving the page. Three entry points to the same movement:

1. **Prev / Next** above the preview, with a position counter, for moving without deciding.
2. **"Approve another file"** as the primary button on the existing success popup, for moving
   straight after deciding.
3. The **URL stays in sync**, so a refresh or a copied link still resolves to the document on
   screen.

Everything else about the page — the metadata table, the destination guard, the decision REST
calls, the popup artwork — is unchanged.

## The queue

### Query

One query at mount, alongside the existing single-item load:

```
/_api/web/lists/getbytitle('Staging')/items
  ?$filter=OData__ModerationStatus eq 2
  &$expand=File,Author
  &$select=ID,FileLeafRef,OData__ModerationStatus,Created,Author/Title,
           File/Length,File/ServerRelativeUrl
  &$orderby=Created asc
  &$top=200
```

`OData__ModerationStatus eq 2` is Pending. Oldest first, so the queue drains in the order
things arrived. The `$select` is the same field set the single-item load already uses, so the
queue entry can populate the page header without a second request per document.

### Scope comes from permissions, not from a filter

SharePoint security-trims list queries. Inheritance is broken per unit folder in Staging, so
the caller only gets items in folders they can read. An approver holding Design on their unit
gets their unit's pending files and nothing else, with no filter logic on our side and nothing
to keep in sync when units are added.

Two consequences to be clear about, because both are properties of the trim rather than
choices we are making:

- **The trim is by folder access, not by role.** Anyone with read access to a unit folder —
  an uploader in that unit — sees the same queue if they open the page. They cannot act on it:
  the approve REST call 403s for Contribute. This does not widen access; an uploader can
  already open `?itemId=` for any file in their folder today. It makes existing access easier
  to walk, which is the point of the feature.
- **An approver covering several units gets all of them in one queue.** Intended. It is the
  same "what is waiting for me" set, just spanning more folders.

A Full Control account reads everywhere, so its queue is every pending file on the site. Hence
the `$top=200` cap. Not a correctness limit — a limit on how much a site admin accidentally
pulls while testing.

### Captured once

The queue is fetched at mount and **not re-queried after a decision**. Deciding on a document
removes it from the server-side pending set; re-querying would make items disappear from under
the approver's position and turn Next into an unpredictable jump. Instead each entry carries a
local `decided` flag, set when its decision succeeds. Positions stay fixed for the life of the
page; a reload gets a fresh queue.

### When the current item is not in the queue

Opening an already-approved or already-rejected document by direct link is legitimate — it is
how someone reviews a past decision. That item will not appear in a Pending query.

Handle it by **prepending the current item to the queue** when its ID is absent, at position 0.
The counter reads `1 of 13`, Prev is disabled, Next moves into the pending set. The page never
has to render a "you are not in the queue" state, and no navigation control has to be special
cased.

## Navigation

### Controls

A single row above the preview, replacing the bare `Preview` heading:

```
Preview                                    ‹ Prev     3 of 12     Next ›
```

- **Prev** disabled at position 0; **Next** disabled at the last position.
- The counter is the position in the whole queue, including decided items — a number that
  moves backwards as you work would be worse than no number.
- Decided entries stay navigable. Landing on one shows a badge beside the counter reading
  `✓ Approved` or `✕ Rejected`, and the decision panel's Ok button is disabled — the decision
  is already recorded and resubmitting it would fire the REST calls a second time. An approver
  can step back to confirm what they did, not to redo it.

### What a swap does

Moving to another document, by either route, must reset **all** per-document state:

| State | On swap |
|---|---|
| `item` | replaced from the cached queue entry — no refetch, the queue query already selected every field the header needs |
| `fieldText` | cleared, then refetched for the new item (`FieldValuesAsText` is per-item and cannot be batched into the queue query) |
| `comments` | **cleared** |
| `decision` | reset from the new item's own `OData__ModerationStatus` |
| `submitError` | cleared |
| `submitted` | cleared — dismisses the popup |
| URL `?itemId=` | rewritten via `history.replaceState` |

Carrying a comment from one document onto the next would attach one approver's reasoning to a
different document's audit trail. It is the one item on this list that is a data-integrity bug
rather than a cosmetic one, so it gets called out here and gets a test.

`history.replaceState` rather than `pushState`: Back should return the approver to the Staging
library they came from, not walk them backwards through their own review session one document
at a time.

### Loading during a swap

The existing full-page `Loading document…` blanks everything. During a swap it would flash the
whole layout on every Next press. Because the queue entry already holds the header and preview
data, only `fieldText` is in flight — so render the new document immediately and show the
metadata table in a muted placeholder state until its labels arrive. The preview iframe
reloads on its own as its `src` changes.

## The popup

Unchanged in artwork, title, and message. The button row becomes:

- **Primary — "Approve another file"** → advance to the next *undecided* entry.
- **Secondary — "Back to library"** → the existing `backUrl()` behaviour.

When no undecided entries remain, the primary button is not rendered and the message gains a
line: **"That was the last document waiting for your approval."** Offering a Next that leads
nowhere is worse than not offering one.

The primary button advances to the next *undecided* item, whereas the Prev/Next controls move
by *position*. The difference is deliberate: after deciding, the approver wants the next piece
of work; while browsing, they want the next document in the list.

## Error handling

- **Queue query fails.** The page still works as a single-document review. Navigation controls
  are not rendered, the popup keeps its current single button, and a console error is logged.
  The queue is a convenience layered on top of a page that already functions; a failure to
  build it must not block the decision the approver came to make.
- **A decision fails.** Unchanged — the existing `submitError` stripe shows, the popup does not
  appear, and the queue position does not move.
- **The destination guard blocks an approval.** Unchanged, and the approver stays on the
  document so they can retry after reconciliation.

## Testing

- Comment text does not survive a swap in either direction.
- The decision radio reflects the new document's own status, not the previous one's.
- Approving the last undecided item yields a popup with no "Approve another file" button.
- A direct link to an already-approved item renders at position 0 with Prev disabled.
- A failed queue query leaves a working single-document page.
- `history.replaceState` leaves exactly one entry, so Back reaches Staging in one press.

## Out of scope

- **A dropdown or filter for the queue.** Worth adding if real queues turn out to be long
  enough that Prev/Next is tedious. Cannot be judged before the client uses it.
- **Bulk approve.** Approving several documents from one screen is a different feature with a
  different risk profile — the destination guard runs per document, and a bulk action makes it
  easy to approve without looking at the preview at all.
- **Re-querying the queue as items change.** Deliberately excluded above.
- **Sorting or grouping by segment, unit, or uploader.** Oldest-first is the only order the
  feature needs to justify itself.

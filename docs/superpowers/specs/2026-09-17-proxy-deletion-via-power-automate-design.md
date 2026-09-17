# Deletion executed by proxy (`crs@sdguthrie.com` via Power Automate), never by the PIC/Approver/HoD themselves

**Client, 2026-09-17:** *"they cannot directly delete anymore but someone do it on their behalf via
power automate via admin account. Reason being they can delete folders which is dangerous."* And:
*"The flow process must stay the same its just that behind the scenes its done via power automate to
make it look like it is the same flow process but behind the scenes its not."*

**Status: NOT BUILT. Design only.**

---

## 1. What this reverses

Two things shipped very recently, both in the opposite direction:

- **2026-08-15 → present**: the whole deletion/share design rests on *"the approval executes in the
  approver's own browser session… no service account, no flow, nothing acting on anyone's behalf"* —
  deliberately, so the audit row names who actually did it and an approval can never exceed the
  approver's own rights.
- **2026-09-11 / 2026-09-15**: `DELS`/`DELSHC` were restored to `pic`/`pic_hc` (direct delete of their
  own pending/rejected file, no approval needed) and an approver/system-admin gained a direct
  delete/share button on an **approved** document, bypassing the request queue entirely.

Both are being undone here, on purpose, because the underlying grant (`CRS Delete` on the library)
also lets the holder delete a **folder** — which is the real danger, not the file-level action the UI
exposes. Removing the SharePoint-level grant is the only way to close that; a UI-level restriction
does nothing, since the person still holds the permission underneath it.

## 1.5. ADDENDUM (same day): HOD also removed from DECIDING a deletion request

Client: *"Can you remove HOD? HOD no need deletion as well, and Approver cannot delete on Documents
Library anymore."* Confirmed: **Approver (`hou`) is unaffected — still decides both pending- and
approved-stage deletion requests, same as today.** HOD is removed from deciding Documents-library
deletion requests at all (keeps deciding Share requests).

This is a SEPARATE mechanism from §2 below. Removing `DEL`/`DELHC` from `hod`'s SharePoint role only
stops HOD holding the grant to execute a delete personally — it says nothing about who may *approve*
a queued request in `Requests.tsx`, which is governed by `ViewerScope`/`inScope` in
`shared/requests.ts` and keyed on `DEPTVIEW` (HOD's decision scope), never on `DEL`. Built:
`inScope`'s `hodUnits` branch now requires `row.type !== "Deletion"` — HOD's `hodUnits` scope applies
to Share requests only.

## 2. Who loses what

`DEL` / `DELS` / `DELHC` / `DELSHC` come off every persona that currently carries them:

| Persona | Role removed |
|---|---|
| `pic` | `DELS` |
| `pic_hc` | `DELSHC` |
| `hou` | `DELS`, `DEL` |
| `hou_hc` | `DELSHC`, `DELHC` |
| `hod` | `DEL`, `DELHC` |

**Client, explicitly:** *"I think HOD also have to done by proxy."* HoD's department-wide direct
delete (2026-08-20) is included — nobody keeps a direct grant.

**Share is untouched.** *"Share is fine"* — `SHARE`/`SHAREHC` stay exactly as they are, executed in
the approver's own session via `SP.Web.ShareObject`, same as today.

⚠ **Reminder carried over, not part of this build:** the share-expiry auto-removal flow
(`2026-08-30-share-expiry-enforcement-flow-runbook.md`) is still designed and not built. Client
flagged it again as a known gap — *"a reminder we haven't implement the auto removing the person via
date after sharing"* — but did not ask for it in this batch. Keep it on the list; do not fold it into
this spec.

## 3. The mechanism

The **UI stays byte-for-byte the same**. A PIC still raises a request (or, for their own pending
file, still presses Delete with no approval step); a HoU/HoD still sees the same queue, the same
decision dialog, the same required-reason gate. What changes is which side actually calls
`GetFileById(...)/recycle()`.

### 3.1 The request-and-approve path (Documents/HC Documents, PIC → HoU/HoD)

Unchanged UI. In `applyDecision` (Requests.tsx), on **Approve**:

1. Resolve the file's **current live location** — reuse `resolveStamped`/the `SubmissionFileId`
   fallback already built for the 2026-09-10 "a request survives its document being routed" fix.
   This is the same resolution the client-side code does today right before calling `recycle()`
   itself; it simply keeps doing that resolution and stops there.
2. Write the resolved `ItemUniqueId` (and, if useful, the resolved library title) onto the
   `CRS Requests` row, alongside `Status=Approved`, `DecidedBy`, `DecisionNote` — exactly the fields
   it writes today. This is a write to `CRS Requests`, a list the approver already has `CRS Request`
   (Add+Edit) rights on; it needs no new SharePoint permission.
3. **Does not call `.recycle()`.** That is the whole change on the client side for this path.

### 3.2 The no-approval-needed path (PIC's own pending/rejected file; the 2026-09-15 direct-delete-on-approved-document shortcut for Approver/system admin)

**Client, explicitly:** *"Pending files as well, PIC will still delete but without approval so
crs@sdguthrie.com via power automate flow will just delete."*

Same mechanism, minus the human-decision step: the client writes a `CRS Requests` row with `Status`
**already** `Approved` at creation (never passing through `Pending`), `RequestedBy` = the person
pressing the button, and the resolved file location from step 1 above. Nothing here needs the
`Requests.tsx` decision dialog at all — it is the same write `submitRequest` already does, just with
`Status` pre-set instead of left at `Pending`.

This covers, uniformly:
- PIC deleting their own pending/rejected draft (2026-09-11's `withdrawDialog` on My Submissions).
- The 2026-09-15 direct-delete button for an approver/system admin on an approved document.
- HoD's direct delete, now routed the same way rather than removed outright.

One shape, one flow, one audit trail — whether or not a second person had to approve it.

### 3.3 The Power Automate flow

New flow, `CRS — Execute approved deletion` (working name). Trigger: `CRS Requests`, created or
modified. Condition: `RequestType eq 'Deletion' and Status eq 'Approved'`. Action, signed in as
`crs@sdguthrie.com`:

```
POST _api/web/GetFileById(guid'<the resolved ItemUniqueId column>')/recycle()
```

- **On success**: write the flow's own `Deleted` audit row (3.5). `Status` stays `Approved` — see 3.4
  for why nothing needs to change it.
- **On failure** (file genuinely gone, permission hiccup, etc.): set `Status = 'Failed'`, append the
  reason to `DecisionNote` — the exact behaviour `applyDecision` performs synchronously today, just
  moved into the flow's own error branch (a Scope + "Configure run after: has failed", the pattern
  every other flow here already uses).

⚠ Uses `GetFileById`, which is **web-scoped** — it resolves a document by GUID regardless of which
library it currently lives in. This is why step 3.1's client-side resolution only has to find the
CURRENT id once; the flow needs no per-library branching of its own.

### 3.4 No new dedupe field — reuse `Status`, rely on `recycle()`'s own idempotency

**Client, 2026-09-17, on adding a dedicated dedupe marker:** *"I guess we can reuse the current
one"* — no new column on `CRS Requests`. `Status eq 'Approved'` is the trigger condition on its own.

⚠ **This means a later, unrelated edit to an already-executed row (fixing a typo in `DecisionNote`,
for instance) re-satisfies the trigger and calls `recycle()` a second time.** Accepted deliberately:
`GetFileById(...)/recycle()` on a document that is already in the recycle bin simply errors — nothing
is double-deleted, no data is at risk, and a redundant, harmless error in the flow's own run history
is exactly the kind of thing this project already tolerates elsewhere (*"a duplicate row is something
a human reads past"*, the same fail-open stance the audit flows' own dedupe checks take when they
can't be certain). If this ever proves noisy in practice, the fix is adding the marker field back —
deliberately deferred rather than built defensively now.

### 3.5 Audit attribution — "for now real person, later maybe the service account"

**Client, explicitly:** *"As of now we record it as the current flow using the approver or uploader
but I think soon we might need to change to crs@sdguthrie.com but just keep in mind."*

So the flow's own `Deleted` audit row must read its actor off the **request row** —
`DecidedBy` for an approved-then-executed request, `RequestedBy` for a self-approved one — never the
flow's own identity. Recorded here as a design constraint for THIS build; the "keep in mind" note is
a future toggle, not something to build defensively now (no config flag, no branch — just don't hardcode
the wrong assumption in a way that's expensive to undo later).

⚠ **This is not the only writer of a `Deleted` audit row for this event.** The existing generic
delete-trigger flows (`Audit — approval deletions`, `Audit — Documents deletions`, and their HC
clones) fire unconditionally off the library's own delete event and read `DeletedByUserName` from the
recycle-bin metadata — which will now always read `crs@sdguthrie.com`, since that account is
literally what performs the delete. Left alone, **every proxy deletion logs twice**: once accurately
(the new flow, real actor) and once inaccurately (the existing flow, service account as actor).

**Fix, scoped to the existing flows, not rebuilt:** add one condition clause to each —
`DeletedByUserName is not equal to crs@sdguthrie.com` — before their own `Create item`. An ordinary
deletion nobody routed through this system (if any such route still exists) still logs correctly
through the existing mechanism; a proxy deletion is left to the new, more accurate flow alone.

⚠ This is an edit to FOUR existing, already-verified flows (the normal pair + the HC pair). Treat it
with the same care every HC-clone edit in this project has needed — read each exported definition
before touching it, confirm which library reference and which condition each one actually holds, and
change nothing else.

## 4. What the service account needs

Whichever account the flow's connection uses (`crs@sdguthrie.com`) needs real delete rights on all
six CRS libraries. Simplest route, consistent with how every other elevated flow in this project
works: add it to **`CRS Owners`**, which already holds Full Control on the web regardless of any
folder ACL — no new permission level, no new Group Map row.

## 5. Open questions, none of which block writing this spec but all of which decide the exact build

1. **Exact library reference to swap into the four existing audit flows** for the exclusion clause —
   confirm against each flow's live definition before editing (same caution this project's own HC-clone
   history repeatedly warns about).
2. **Does the flow need to distinguish Staging vs Documents at all?** `GetFileById` doesn't care, so
   probably not — the same flow covers a pending-file proxy-delete and an approved-document
   request-delete identically. Flagged here only so it isn't quietly re-split into two flows later
   for no reason.

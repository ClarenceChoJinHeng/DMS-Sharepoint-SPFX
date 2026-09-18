# Tag/approve by proxy — design

**Date:** 2026-09-18
**Status:** PIC's half approved by client and in progress. Approver's half is WITHDRAWN — see the
correction immediately below. Read that before anything else in this document.
**Companion doc:** `2026-09-18-tag-approve-proxy-reversion.md` — how to undo the PIC half if the lag
turns out to be unacceptable once it's live. Read that one first if this is ever being rolled back.

## ⚠⚠ CORRECTION, 2026-09-18 (same day) — THE APPROVER HALF IS OFF. `CRS Approve` KEEPS `Edit Items`.

Everything below still describes the ORIGINAL two-sided plan (PIC's tagging write AND Approver's
decision write, both moved to proxy) and is kept as the historical record of the reasoning, not as
what is being built. **Only the PIC half is going ahead.**

**Why:** the whole point of moving the Approver's decision write to proxy was to let `Edit Items`
come off `CRS Approve` at final cutover, the same as `CRS Upload`. That is not possible. `ApproveItems`
has a hard SharePoint dependency on `EditListItems` — confirmed live, on this tenant's `CRS Approve`
level, by unticking each independently and watching the other come with it. There is no permission
level, and no combination of group memberships, that grants a person `ApproveItems` without also
granting `EditListItems` — SharePoint enforces the two together in the permission-level editor
itself. So "keep Approve Items (for Draft Item Security's unit-wide pending-file visibility), drop
Edit Items (to stop hand-editing)" — the entire premise of the Approver half of this design — cannot
be built as SharePoint's own permission model works.

**An alternative WAS explored — explicit per-item unique permissions, granted by the proxy account
at upload time, replacing Draft Item Security's native exemption entirely** (break inheritance per
pending/rejected file; visibility becomes "do you hold Read on this specific item" rather than
"do you hold `ApproveItems`"). Worked out mathematically viable at modest volume, and NOT viable at
realistic worst-case volume: ten segments each holding a few thousand simultaneously-open pending
plus rejected files can consume the bulk of a single approval library's 50,000-unique-permission-
scope ceiling on rejected files alone, before a single pending file or folder is counted — and
rejected files are the one population nothing in that scheme naturally shrinks, since nobody is
forced to clean them up. Parked, not built. Worth reopening only with the client's *actual* expected
concurrent pending+rejected volume in hand, not a round stress-test number — see the session that
produced this correction for the full arithmetic if it comes back up.

**The decision, for now:** `CRS Upload` (PIC) loses `Edit Items` — unaffected by any of this, since
`CRS Upload` never held `ApproveItems` and so never had the coupling problem. `CRS Approve`
(Approver) **keeps** `Edit Items`, indefinitely — approvers can still hand-edit a file's tags/name
directly in the library, which is the one piece of the original client complaint this rollout does
NOT close. `ApprovalDocument.tsx`'s decision write (Approve/Reject) has been **reverted** to its
original direct `ApprovedBy`-MERGE-then-`approve()` / raw `OData__ModerationStatus` MERGE sequence —
the `writePendingDecision`/`CRS Pending Decisions` code path built for it is unused by any caller as
of this correction. `BulkApprovePanel.tsx` was never converted in the first place, so it needed no
reverting.

**What stays, unaffected:** everything in the Scope section below **except** "Approver's decision
write" and "Self-approve." PIC's tagging write (`Form.tsx`, `BulkUpload.tsx`, both normal and HC),
and PIC's staging-replace write, proceed exactly as designed. `CRS Pending Decisions`'s schema and
`spPendingDecisions.ts` stay in place, unused, per this project's own habit of parking rather than
deleting unused mechanisms — see the reversion doc's "What to do with the new schema" section, same
reasoning applies here.

**Self-approve** (`autoApproveOwnUpload`) was never converted to the decision-proxy in the first
place — it still does the original direct `OData__ModerationStatus` MERGE in `Form.tsx`, currently
disabled behind `SELF_APPROVE_DISABLED_PENDING_TAG_CONFIRMATION` for an *unrelated* reason (Task 9:
it must not fire until the async tag write is confirmed applied, since PIC's tagging is proxied
regardless of what happens on the Approver side). That gate is real and still needed; nothing here
changes it.

## Problem

Client: an uploader (PIC) or an approver can open a file directly in the SharePoint library view
and rename it or change its metadata columns by hand — bypassing the app's own naming convention
(`[Project] - [Vendor] - [Document Name] - [Date]`) and the tag values chosen in the upload form.
This breaks the standard protocol and the filename.

SharePoint has no way to grant "Edit rights, but only when our app's own code is making the
write" — a REST call from `Form.tsx` running in the uploader's browser and the same person typing
a new value into the library's native "Edit properties" panel are the exact same permission,
exercised the exact same way. The only way to stop the second without also breaking the first is
to take the underlying `Edit`/`Approve` right away from the person entirely, and have the write
performed by something else on their behalf.

This mirrors the delete-by-proxy work already shipped (2026-09-17) — same reasoning, same general
shape, extended to the two writes that make up "tag a file" and "approve/reject a file."

## Scope

**⚠ As of the 2026-09-18 correction above, only the first and last bullets are being built.**
"Approver's decision write" and "Self-approve" are struck through below and kept as the record of
the original, wider scope — read the correction, not these two bullets, for what is actually true.

- **PIC's tagging write** on upload, in both `Form.tsx` and `BulkUpload.tsx`, both normal and HC.
- ~~**Approver's decision write** (Approve/Reject) in `ApprovalDocument.tsx` (the Preview page) and
  `BulkApprovePanel.tsx` (the command-bar sidebar), both normal and HC.~~ WITHDRAWN — `CRS Approve`
  cannot drop `Edit Items` without also dropping `ApproveItems`, so there is nothing left to proxy
  this write in order to achieve. `ApprovalDocument.tsx` has been reverted to its direct write.
- ~~**Self-approve** (`autoApproveOwnUpload`) in `Form.tsx` — a Head of Unit's own upload can
  auto-approve today with no separate approver; that write also moves to proxy.~~ WITHDRAWN for the
  same reason — never actually converted in code either; still the original direct MERGE, currently
  gated off for the unrelated Task 9 reason described in the correction above.
- **PIC's staging-replace write** (`Form.tsx`'s "yes, replace it" consent on a name clash) — see
  its own section below. Found during review, not in the original scope: it's a third Edit-shaped
  write that the tagging/decision writes don't cover, and it would otherwise 403 the moment
  `Edit Items` comes off `CRS Upload`.

## Non-goals

- **Share is untouched.** `CRS Share` (Manage Permissions) is a different concern — granting
  access, not changing a file's name or metadata — and the client's stated problem is specifically
  about renaming/retagging. Do not fold Share into this.
- **Deletion is already covered** by the 2026-09-17 proxy-deletion work. Not revisited here.
- **Auto-route's own routing logic does not change.** It already reacts to the moderation status
  changing, regardless of who or what changed it — nothing here needs to touch it.
- **The client-side safety checks stay exactly where they are.** The destination-folder guard, the
  name/approved-side clash checks, the `ApproveItems` eligibility probe — all of these are *reads*,
  not writes, so none of them are affected by removing Edit/Approve rights. They keep running in
  the user's own browser session, unchanged, before anything gets written anywhere.

## The user-facing requirement, stated plainly

**The UI/UX must be identical to today.** Same screens, same buttons, same number of clicks. The
proxy mechanism is invisible. The one accepted change is timing: a write that used to be instant
now takes roughly a poll cycle (about a minute) to actually land, and the existing screens (My
Submissions, the approval queue) already poll and refresh themselves, so a file just shows
"Pending" a little longer before settling — the same way it would if a human took a minute to
notice and click. **The client has explicitly accepted this lag** and asked for a reversion plan
to be ready in case they change their mind once it's live.

## Architecture

Two independent proxy mechanisms, both the same shape as deletion: a browser write to a list the
user already has rights to, and a flow that watches it and performs the actual privileged write as
`crs@sdguthrie.com`.

```
PIC uploads (Files/Add, unchanged — Add rights only)
        |
        v
Browser writes tag values into the CRS Submissions row
(already written on every upload today; extended with a new payload field)
        |
        v
CRS - Apply pending tags  (flow, polls ~1 min, runs as crs@sdguthrie.com)
        |
        v
Real file gets its metadata columns MERGEd on
```

```
Approver clicks Approve/Reject (client-side checks run first, unchanged)
        |
        v
Browser writes a decision row to CRS Pending Decisions (new list)
        |
        v
CRS - Apply pending decisions  (flow, polls ~1 min, runs as crs@sdguthrie.com)
        |
        v
Real item's OData__ModerationStatus gets flipped
        |
        v
Auto-route picks it up exactly as it does today (unchanged)
```

**One structural win:** because both flows resolve the target file by its unique ID
(`GetFileById(guid'...')`, the same web-scoped call the deletion flow already uses), **one flow
covers both normal and HC libraries** — no separate HC clone needed for either of these two flows,
unlike almost everything else in this project. Worth stating explicitly since forgetting the HC
clone is this codebase's single most common recurring defect.

**A side benefit:** once PIC/Approver no longer hold Edit/Approve on the document libraries at
all, SharePoint's *native* Approve/Reject command and the native "Edit properties" panel stop
working for them too, not just our own screens. This closes a gap this project has documented
repeatedly as unfixable by the app alone (a web part cannot intercept a SharePoint-native
control) — it gets closed as a side effect of this change, for free.

## Schema

### `CRS Submissions` — extended, not replaced

**Do not repurpose the existing `MetadataSnapshot` field for this.** It already exists, is
already written on every upload, and is checked directly (`Form.tsx` ~3329): it stores tier and
metadata *labels only* (`snapshot[l.column] = l.label`), it does **not** include `Document Type`,
`Year`, `Keyword`, or `Business Segment`, and it's built purely for **display** — My Submissions'
detail panel reads it to show what a deleted/archived/replaced record used to carry. Its
correctness requirement is "good enough for a person to read," which is a different bar from "the
exact payload a flow uses to write real managed-metadata fields." Reusing it would entangle two
concerns that have different tolerances for being slightly wrong, and risks silently degrading a
feature (deleted-record display) that already works and is well-tested.

**New fields instead:**
- `TagPayload` (Note) — every `{FieldName, FieldValue}` pair `Form.tsx`/`BulkUpload.tsx` currently
  send to `validateUpdateListItem` for this file, serialised as JSON. This is the *exact* write
  payload — taxonomy fields already in `Label|GUID` format (via the existing `toTaxValue` /
  `buildLevelFormValues` helpers, unchanged), not just labels. The full field list is whatever
  `formValues` in `Form.tsx`/`BulkUpload.tsx` currently builds (tier values, confidentiality,
  Document Type, Year, Remark, LegallyPrivileged, ProjectName, Vendor/CustomerName, Keyword,
  DocumentDate, Business Segment) — reference the code at implementation time rather than
  re-deriving the list by hand here, since it's whatever's actually being sent today.
- `TagStatus` (Text) — blank/`Pending` -> `Tagged` (success) or `Failed` (with `TagError` holding
  the reason). Same three-state shape this codebase uses everywhere: blank/unknown is not the same
  as a definite failure.
- `TagError` (Note, optional) — populated only on `Failed`.

### `CRS Pending Decisions` — new list

One row per approve/reject decision, written instead of the direct `OData__ModerationStatus` MERGE.

| Column | Type | Notes |
|---|---|---|
| `ItemUniqueId` | Text | The document being decided on. Resolved the same way requests already are - `ItemUniqueId` first, falling back to the `SubmissionFileId` stamp on a 404, reusing `resolveStamped`/`findByStamp` from `shared/requests.ts` rather than a third copy of that logic. |
| `Decision` | Text | `Approved` or `Rejected` |
| `DecidedBy` | Text | Email of whoever clicked - the approver, or the uploader themselves for a self-approve. |
| `DecidedAt` | DateTime | ISO. |
| `RejectionComment` | Note | Only for `Rejected`. |
| `IsSelfApprove` | Boolean | True when this came from `autoApproveOwnUpload`, not a human decision. Exists so the applying flow (and any future reporting) can tell the two apart without guessing from `DecidedBy` matching the file's own author. |
| `Status` | Text | `Pending` -> `Applied` or `Failed`, same shape as `TagStatus`. |
| `StatusError` | Note | Optional, on `Failed`. |

Deliberately a **new, purpose-built list**, not an extension of `CRS Requests`. That list's entire
design - `RequestType: "Deletion" | "Share"`, a requester and a *separate* decider, the whole
`ViewerScope`/`inScope` visibility model - is built around *someone else decides on your behalf*.
"The approver decides their own upload" (self-approve) and "an approver decides someone else's
upload" (the ordinary case) are a different shape: one action, not a request-then-decision pair,
and bolting a third `RequestType` onto `CRS Requests`' already heavily-worked machinery risks
breaking something that currently works, for no real benefit.

## Data flow

### Upload + tag (PIC, HoU, any uploader)

1. `Files/Add` - unchanged. Needs `AddListItems` only, which nothing in this change touches.
2. Browser builds the same `formValues` array it builds today, but instead of POSTing it to
   `validateUpdateListItem`, writes it (JSON-encoded) into the `TagPayload` field of the
   `CRS Submissions` row for this file, alongside the existing `SubmissionId`/`BatchId`/
   `SubmissionFileId`/`MetadataSnapshot` fields (all unchanged). `TagStatus` starts blank.
3. `CRS - Apply pending tags` (new flow, polls `CRS Submissions` every ~1 minute) finds rows with
   `TagStatus` blank, parses `TagPayload`, resolves the file via `ItemUniqueId` (or the
   `SubmissionFileId` stamp as a fallback, same pattern as everywhere else this project resolves a
   file), and POSTs the SAME `formValues` array to `validateUpdateListItem` - as
   `crs@sdguthrie.com`, not the uploader.
4. On success: `TagStatus = Tagged`. On failure: `TagStatus = Failed`, `TagError` holds
   SharePoint's own response (never swallowed - same reason this project always logs the raw
   `HasException` detail rather than a generic "tagging failed").

**What the UI shows meanwhile:** nothing new. My Submissions already polls itself. A file with
`TagStatus` still `Pending`/blank simply shows blank metadata columns for a short window, same as
it always has for a file uploaded moments ago before this change - the only difference is the
window is now measured in up to a minute rather than being instant.

### Approve / Reject (a human approver)

1. Every existing client-side check runs exactly as today - destination-folder guard,
   `documentsFileClash`, whatever else `ApprovalDocument.tsx`/`BulkApprovePanel.tsx` already do
   before deciding whether to allow the click through. None of this changes; all of it is reads.
2. If the checks pass, the browser writes a row to `CRS Pending Decisions` instead of MERGEing
   `OData__ModerationStatus` directly.
3. `CRS - Apply pending decisions` (new flow, polls ~1 minute) finds `Status = Pending` rows,
   resolves the file, and applies the decision **as two separate writes, not one MERGE** —
   matching exactly what `ApprovalDocument.tsx`'s current `approve()`/`reject()` calls do today,
   because SharePoint rejects a single MERGE that combines `OData__ModerationStatus` with any
   other field (a real bug this project already paid for once, 1.0.348.0: that MERGE 500s on
   every attempt, silently, because the write actually needed is illegal). For a **Reject**, this
   is a non-issue — `OData__ModerationStatus` + `OData__ModerationComments` together in one MERGE
   is fine, since both are moderation fields. For an **Approve**, it must be:
   - First, a plain MERGE setting `ApprovedBy` (and `ApprovalComment` if there's a typed note) —
     an ordinary field edit, no moderation status touched, so it doesn't hit the restriction.
   - Then the moderation status flip — `File.approve()` (a dedicated method, not a field MERGE,
     so it also doesn't hit the restriction) or the flow-equivalent MERGE of
     `OData__ModerationStatus` alone. Any edit to a moderated item reverts its status to Pending
     unless the *same* write re-asserts it, so this step must come after the `ApprovedBy` write,
     never combined with it.

   Getting this wrong doesn't error loudly — it 500s on the second call, or silently leaves the
   item back at Pending after a "successful" first write, which is exactly how this bug went
   unnoticed in `ApprovalDocument.tsx` for a while before it was caught, all as `crs@sdguthrie.com`.
4. Auto-route reacts to the status change exactly as it does today - untouched.
5. `Status = Applied` or `Failed` on the decision row.

### Self-approve - the one sequencing detail that matters

Self-approve happens in the same browser session as the upload, immediately after tagging, and
must not let a document route while its metadata is still blank - the exact defect this project
already hit once (2026-08-26, SDG, a document filed a tier shallower than its siblings with no
error anywhere).

**The eligibility check stays exactly where it is** - `probeFolderApproveAccess`
(`EffectiveBasePermissions`, a read) still runs client-side, unchanged, since reading whether the
uploader *would* hold `ApproveItems` on that folder needs no Edit/Approve right of their own.

**The decision write is deferred until tagging is confirmed complete.** After writing the
`TagPayload` row, the browser polls the `CRS Submissions` row (the same kind of short-interval
polling My Submissions and the Requests page already do) until `TagStatus` reads `Tagged` - then,
and only then, writes the `CRS Pending Decisions` row with `IsSelfApprove = true`. If `TagStatus`
comes back `Failed`, no decision row is written at all; the document stays Pending, tagging failed,
and that failure is visible the same way any other `Failed` `CRS Submissions` row would be.

**Net effect on timing:** self-approve takes roughly two poll cycles rather than one (tag flow,
then decision flow) - call it up to two minutes rather than one, worst case. The user does nothing
differently; they see "Pending" for a little longer before it settles to "Approved" on its own.

### Replacing a pending draft (`Form.tsx`'s "yes, replace it" consent)

Found during review of this design, not in the original brief. `Form.tsx` has one write this
document's original scope missed entirely: the 2026-08-28 "replace a pending draft" feature. On a
name clash, if the uploader consents, the browser calls `Files/Add(overwrite=true)` directly
against the *existing* item — including, deliberately, a colleague's draft the uploader can't
even see (Draft Item Security hides it from them; the overwrite call doesn't need to read it
first). This is a real content write, separate from the `validateUpdateListItem` tagging call,
and SharePoint requires `EditListItems` — not just `AddListItems` — to overwrite an *existing*
item's content this way. Uploading a brand-new file only needs Add; replacing one is, under the
hood, an edit of that list item. So this call 403s the moment `Edit Items` comes off `CRS Upload`,
with nothing else in this design covering it.

**It needs no new mechanism.** The 2026-09-17 proxy-deletion work already exists precisely to
recycle a file by `ItemUniqueId` as `crs@sdguthrie.com`, with no regard for who authored it or who
requested it — which is exactly the shape this needs, since "replace" today is already a
unilateral, no-approval action taken by whoever clicked Yes, whether or not the file is theirs.
Re-express the replace as: delete the clashing draft via the existing proxy path, then let the
browser perform an ordinary, Add-only upload of the new content under the same name once the old
item is gone.

1. On consent, the browser writes a self-approved deletion request — same shape as
   `MySubmissions.tsx`'s `writeApprovedDeletionRequest`, `Status` pre-set to `Approved`, targeting
   the clashing draft's `ItemUniqueId` — no new list, no new flow. The already-built
   `CRS - Execute approved deletion` picks it up and recycles the file as `crs@sdguthrie.com`,
   exactly as it does for every other proxied deletion.
2. The browser polls (same short-interval pattern as self-approve's `TagStatus` poll above) until
   the old item stops resolving — `GetFileById(guid'<oldId>')` starts 404ing — confirming the
   recycle has actually happened.
3. Only then does the browser call the ordinary `Files/Add` — **no `overwrite` flag, because
   there is nothing left to overwrite** — uploading the new content under the same name. This is
   an ordinary new-file upload, needing `AddListItems` only, unaffected by removing `Edit Items`.
4. Tagging then proceeds exactly as every other upload's tagging does (`TagPayload` row, applied
   by `CRS - Apply pending tags`).

**On recoverability, since this changes *how* it's achieved, not whether it exists:** the current
code comment ties this feature's safety to version history being a "hard prerequisite" for
recovering an overwritten draft. Routing it through the deletion-by-proxy path instead means
recovery is via the **recycle bin** (93 days, same guarantee the deletion feature already relies
on) rather than version history — a different mechanism, not a weaker one; nothing about
recoverability actually regresses.

**On timing:** this adds a full poll cycle to what is today an instant overwrite — worse than the
"up to a minute" already accepted for tagging/decisions, since it's poll-then-Files/Add rather
than fire-and-forget, and it happens synchronously inside the upload button's click handler. The
UI needs some kind of "still working" state while this plays out (a spinner/disabled state on the
consent dialog's own button, most likely — not a new page-level affordance), which isn't detailed
further here; worth deciding at implementation time, same as the general "where does Failed
surface" open question below.

## Permission-level changes - the highest-risk step in this whole rollout

**⚠ CORRECTED, 2026-09-18 — read the box at the top of this document first.** Everything in this
section originally covered BOTH `CRS Upload` and `CRS Approve`. Only `CRS Upload` is actually being
changed. `CRS Approve` keeps `Edit Items` — SharePoint will not allow it to be removed without also
removing `ApproveItems`, which was confirmed live and is the entire reason the Approver half of this
design was withdrawn. Do not untick `Edit Items` on `CRS Approve` on the strength of the steps below;
they are corrected to say so, but the original wording elsewhere in this document (Scope, Data flow)
still describes the withdrawn plan and should not be followed for the Approver side.

Today, `UPL`/`UPLHC` map to the `CRS Upload` SharePoint permission level, and `APR`/`APRHC` map to
`CRS Approve`. `CRS Upload` currently grants `Edit Items` (needed for the direct write this change
removes) alongside what it still needs to keep (`Add Items`). `CRS Approve` keeps `Edit Items`
**permanently** — it is not part of this cutover.

**Once the tagging code and flow are built and proven, `Edit Items` is unticked from `CRS Upload`
only** in Site Settings -> Permission Levels, on each site. This is a manual, site-level SharePoint
change - not something the app package can do.

**This step must be LAST, and only after everything else is deployed and confirmed working with
the OLD permissions still in place.** If `Edit Items` is removed before the new code/flow is
live, every upload on the site breaks immediately - the client-side write attempts still exist in
the old code and will simply 403. The correct order is:

1. Build and deploy the new `TagPayload` schema and `CRS — Apply pending tags`.
2. Build and deploy the code changes (`Form.tsx`, `BulkUpload.tsx`) to write to `CRS Submissions`
   instead of writing directly. `ApprovalDocument.tsx` and `BulkApprovePanel.tsx` are NOT touched -
   they keep their original direct decision writes, since `CRS Approve` is not changing.
3. Test end to end **with the old permission levels still in place** - the new mechanism works
   alongside the old rights at this point; nothing is broken by testing here, because the direct
   tagging write path is already gone from the code, so the new proxy path is what's actually being
   exercised regardless of what rights the user still holds.
4. Only once that's confirmed working: remove `Edit Items` from `CRS Upload` on each site, as the
   final cutover. `CRS Approve` is untouched at every step.

**Migration for already-provisioned groups:** same as every other role change in this project -
existing groups keep granting whatever the permission level currently allows until the level
itself is edited (step 4 above), at which point the change applies to every holder of that level
at once, with no group-by-group action needed (this is a level-level change, not a per-group role
change, so the usual "delete and recreate the group" migration pattern does not apply here).

## HC parity

The tagging flow is library-agnostic by construction (resolves by unique ID), so no separate HC
clone is needed for it. The permission-level change applies identically to `CRS Upload`/`CRS
UploadHC` mappings across both the normal and HC pairs.

## Bulk Upload

`BulkUpload.tsx` is in scope (it's open to every uploader, not admin-only, per the 2026-08-22
change) and gets the identical treatment on its tagging write. It has no approval step of its own
(writes straight to the approved side), so only the tagging half ever applied to it - unaffected by
today's correction either way.

## Error handling

- A failed tag write: `TagStatus = Failed`, reason preserved. Never silently dropped.
- **Where does a `Failed` state actually surface to a person?** Not fully specified here - worth
  deciding at implementation time whether this needs a new UI affordance (an admin-visible list of
  failed tags) or whether it's sufficient to be visible via the existing screens reading this new
  field. Flagging as an open question rather than guessing.

## Testing

- Normal upload -> confirm tags land within a poll cycle, `TagStatus = Tagged`.
- HC upload -> same, confirm the one flow (no HC clone) handles it correctly.
- Ordinary approve/reject via the Preview page and via the bulk sidebar -> confirm nothing changed
  from today's behaviour, since neither screen's decision write is touched by this change any more.
- Self-approve -> confirm it stays disabled behind `SELF_APPROVE_DISABLED_PENDING_TAG_CONFIRMATION`
  until Task 9 replaces the flag with a real poll-based gate; not otherwise affected by today's
  correction.
- A deliberately malformed tag payload -> confirm `TagStatus = Failed` with a real reason, not a
  swallowed error.
- Replace a pending draft -> confirm the old file is recycled, the new content lands under the
  same name, tags are applied, and the whole sequence completes with no extra click. Do this once
  for the uploader's own draft and once for a `hidden` clash (a colleague's draft they can't see)
  - the two go through the same proxy delete either way, but the second is the one most likely to
  reveal a mistaken assumption about needing to read the target item first.
- Confirm a PIC still sees only their own pending file in staging after the cutover (Draft Item
  Security's author exemption is unrelated to `Edit Items`, so this should be unaffected - but
  it's cheap to verify given how much rides on it).
- **The permission-level cutover itself** - the step most likely to go wrong. Test on a throwaway
  site or a quiet window, not live, and confirm every existing screen still works for a PIC/approver
  test account immediately after `Edit Items` is removed.

## Open questions

1. Where does a `Failed` tag or decision actually get surfaced to a human? (noted above)
2. Does removing `Edit Items` from `CRS Approve` affect anything else that role currently relies on
   besides the moderation-status write and the read-only eligibility probe? Worth a grep of every
   place `APR`/`APRHC` is checked before the cutover step, to be sure nothing else silently assumed
   Edit rights.
3. Timing for the actual cutover on SDG's live tenant - this should not happen on the same day as
   any other major deploy, given how much is riding on the ordering being right.

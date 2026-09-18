# Tag/approve by proxy — reversion

**Date:** 2026-09-18
**Companion to:** `2026-09-18-tag-approve-proxy-design.md`

> ⚠⚠ **CORRECTED, SAME DAY — THIS DOC NOW COVERS PIC ONLY.** The design's Approver half (the
> `ApprovalDocument.tsx`/`BulkApprovePanel.tsx` decision write, `CRS — Apply pending decisions`, and
> removing `Edit Items` from `CRS Approve`) was withdrawn before it ever went live — see the design
> doc's own correction box. There is nothing to revert on the Approver side because nothing was ever
> cut over: `ApprovalDocument.tsx` has already been reverted to its original direct write,
> `BulkApprovePanel.tsx` was never touched, `CRS — Apply pending decisions` was never built, and
> `CRS Approve` never had `Edit Items` removed. Everything below that mentions "Approver" or the
> decision-write flow describes the ORIGINAL, wider plan and no longer applies — follow it for the
> PIC/tagging half only.

Prepared alongside the forward design, at the client's request, in case the lag introduced by
proxying the tagging write turns out to be unacceptable once it's live. This is **not** a plan to
be executed now — it exists so a reversion can happen quickly and safely later, without having to
reconstruct the reasoning under pressure.

## What "revert" means here

Give PIC back their direct `Edit Items` right on the document libraries, and restore the
client-side code that writes tags directly, instead of via `crs@sdguthrie.com`. This undoes the
PIC-side 2026-09-18 change; it does **not** touch the separate 2026-09-17 deletion-by-proxy work,
which stays as it is regardless. It never had an Approver side to undo — see the correction above.

## Order matters — do these in sequence, not in parallel

Doing this out of order risks either losing a file's tags entirely (if a request row is sitting
mid-flight when the code is reverted) or a flow silently overwriting a manual correction someone
made after direct writes resumed. Follow this order:

### 1. Turn off both flows first

`CRS — Apply pending tags` and `CRS — Apply pending decisions`. This stops them picking up any new
rows. Nothing breaks by doing this alone — uploads/approvals made via the (still-current) proxy
code will simply sit `Pending` and not get applied, same as if SharePoint were down for a moment.

Leave `CRS — Execute approved deletion` running. It's the 2026-09-17 proxy-deletion flow, not
introduced by this feature, and the staging-replace-draft path added here reuses it rather than
adding a third flow — turning it off would also stop ordinary deletion requests working, which is
outside the scope of this reversion.

### 2. Drain what's already in flight

Before reverting the code, check both lists for rows still `Pending`:
- `CRS Submissions` with `TagStatus` blank or `Pending`.
- `CRS Pending Decisions` with `Status` = `Pending`.

These represent uploads/approvals that happened in the last cycle before the flows were turned
off. Turn the flows back **on** briefly, let them clear the backlog, confirm both lists show no
`Pending` rows, then turn them off again. Skipping this step means whichever files are mid-flight
at the moment of reversion never get tagged or approved at all — they'd need to be fixed by hand.

Also check `CRS Requests` for a self-approved deletion row still `Approved` but not yet actually
recycled — the staging-replace-draft path writes one of these and then polls for the recycle to
land before doing the follow-up upload. If the code is reverted mid-poll, that in-flight replace
never completes its second half (the fresh `Files/Add`); let `CRS — Execute approved deletion`
finish recycling first (it's staying on regardless, per the note above), then confirm nobody is
sitting on a stuck "replacing…" state in the browser before reverting.

### 3. Revert the code

`git revert` the commits that introduced this feature — restoring the original direct writes in
`Form.tsx`, `BulkUpload.tsx` (the `validateUpdateListItem` POST), `ApprovalDocument.tsx`, and
`BulkApprovePanel.tsx` (the `OData__ModerationStatus` MERGE). Rebuild and redeploy the package.
Since the flows are already off (step 1) and drained (step 2), there's no window where both the
old and new write paths are live at once.

### 4. Restore `Edit Items`

Site Settings → Permission Levels → re-tick `Edit Items` on `CRS Upload` and `CRS Approve`, on
each site this was applied to. This is the mirror of the original cutover's last step, and — same
as then — it should happen last, once the code writing directly is already back in place and
confirmed working. Doing this before step 3 gives PIC/Approver back the ability to edit directly
while the app itself is still writing through the (now-disabled) proxy path, which just means
nothing tags/approves for anyone until step 3 catches up — not dangerous, just confusing if left
that way for long.

### 5. Test

Same list as the forward design's own testing section: a normal upload, an HC upload, an ordinary
approve/reject through both screens, and self-approve — confirm all four are instant again, the
way they were before 2026-09-18.

## What to do with the new schema

**Leave it in place.** `TagPayload`/`TagStatus`/`TagError` on `CRS Submissions`, and the whole
`CRS Pending Decisions` list, are harmless once nothing is writing to or reading them — matches
this project's own established habit of parking unused mechanisms rather than deleting them,
since a deleted column takes its data with it and a later "maybe we should try this again"
conversation would otherwise mean rebuilding it from scratch. Only delete `CRS Pending Decisions`
outright if the client is certain this direction is permanently abandoned, not just paused.

## What NOT to do

- **Do not restore `Edit Items` before the code revert is deployed and confirmed.** That's the
  exact wrong-order risk the forward design's own cutover section warns about, just in reverse.
- **Do not skip step 2.** A handful of files silently never getting tagged is a small enough
  problem that it's tempting to skip draining and just fix stragglers by hand afterward — but
  "silently never tagged" is exactly the kind of defect this codebase has repeatedly been bitten
  by discovering weeks later, not the day it happened.

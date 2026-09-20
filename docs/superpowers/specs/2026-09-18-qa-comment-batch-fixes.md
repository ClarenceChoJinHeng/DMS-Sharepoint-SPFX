# Client QA comment batch — 2026-09-18 (comments 1, 4-10, 12-17)

Client sent a spreadsheet-style batch of numbered QA comments as screenshots (comment numbers 4-17,
plus two direct list items). This spec records what each one asked for, whether it was fixed in code,
and — for the ones that cannot be fixed from this repository — exactly what needs to happen instead.

## Built and verified (`tsc --noEmit` clean, `npx heft test --clean` → 0 failures, lint at the

documented pre-existing 43-warning baseline, no new categories)

### Comment 4 — native browser confirm on "Delete File"

The Document Upload form's "Delete File" button (removing a staged, not-yet-uploaded file) was
firing a NATIVE `window.confirm()` — `sdguthrie.sharepoint.com says: ...` — on top of the app's own
styled UI. This project has removed every other native confirm it ever shipped (see CLAUDE.md's
2026-08-15 and 2026-09-02 "the replace-clash popups are gone" sections); this was one of four that
had survived that pass.

**Fixed**: built one reusable custom confirm dialog (`confirmDialog` state + a single render block,
reusing the existing `.dms-popup-overlay`/`.dms-popup` markup the clash dialog already uses) and
routed all FOUR native confirms in `Form.tsx` through it — not just the one reported, for
consistency: "Delete File" on the open editor, "Remove set N", "Discard/Clear this set", and the
top-level "Cancel" button. `grep window.confirm src/webparts/form/components/Form.tsx` now returns
only comments, no live calls.

### Comment 5 — Confidential Level dropdown order

The dropdown showed the term store's alphabetical order (Confidential, Highly Confidential,
Restricted). Client wants: Highly Confidential, Confidential, Restricted.

**Fixed**: new shared module `src/shared/confidentialityOrder.ts`
(`sortByConfidentialityOrder`/`confidentialityRank`, pure, tested) applied to both the upload form's
and Bulk Upload's Confidential Level `<option>` lists. Deliberately kept OUT of `hcRouting.ts` —
that module's whole reason for existing is "every rule here has a wrong version that looks identical
on screen" (a SAFETY module); mixing a cosmetic sort into it would blur that. Unrecognised labels (a
segment with a differently-named level) sort past the three known ones, in their original order —
stable sort, so nothing is silently reordered by guesswork.

### Comment 7 — My Submissions legend order

The status legend at the bottom of My Submissions listed Pending, Rejected, Approved, Deleted,
Cancelled, Replaced, Archived, Share Requested, Delete Requested. Client's own numbering (1-9 on
their sheet) wants: Pending, Approved, Rejected, Cancelled, Replaced, Deleted, Archived, Share
Requested, Delete Requested.

**Fixed**: pure reorder of the `<li>` elements. Nothing about which seven items are hidden on the
Permission tab changed.

### Comments 1 & 8 — "Remark for Approval" → "Remark for Approver"

**Fixed** in both places this label exists: the upload form's own field (`Form.tsx`) and the
approver's detail panel on `ApprovalDocument.tsx` (which shows the SAME remark, read back for the
person it says it's addressed to). The underlying SharePoint column is still internally named
`Remark` — untouched, per gotcha #4 (one unknown field name fails the whole write).

**Not changed, deliberately**: `MySubmissions.tsx`'s shared `documentDetails.ts` module shows this
field as plain "Remark" (no "for Approval/Approver" suffix) and was not touched by the original
2026-09-10 rename either — not in scope for this comment, which only pointed at the upload form's
field. Bulk Upload's own Remark field also still says plain "Remark" — not reported, left as-is.

### Comment 15 — Status text should be capitalised

My Submissions' "Submissions" tab showed lowercase status text ("1 approved", "1 archived").

**Fixed**: `RECORD_STATE_LABEL` in `shared/submissionRecords.ts` (deleted/replaced/archived/
cancelled/pending/not checked) and the separate `.toLowerCase()` call on the Approved/Pending/
Rejected tally in `MySubmissions.tsx` were BOTH lowercase and are now Title Case. Type-level values
(`deleted`, `cancelled`, `archived`, `withdrawn`, `settling`, `unknown`) are untouched — display only.
Updated the four hardcoded-string test expectations in `submissionRecords.test.ts` to match.

### Comment 17 — Reject dialog wording

"Please provide your reasoning below." → "Please provide your reject reason below.", on the Requests
page's decision dialog (reject branch only — the approve branch's "Add a note below if needed." is
unchanged, not part of this ask).

### Comments 6 & 10 — Share request card showing two emails

The Requests page's decided-request card showed both the REQUESTER's email (top line) and the
RECIPIENT's email (share line) for a Share request — client: "no need show share by / Remove sharer,
only show recipients email... Just put Share to: XXX@email.com."

**Fixed, Share cards only**: the requester's email is dropped from the top meta line for
`r.type === "Share"` (kept for Deletion cards — deciding whether to delete something still needs to
know who's asking, a different question from deciding whether to grant access). The recipient line
now reads "Share to: &lt;email&gt;" so the one remaining email is unambiguous.

⚠ **Trade-off worth flagging, not silently accepted**: removing the requester's identity from a
Share card removes context an approver might reasonably want (who is asking, from which unit) before
granting access to a document. Built exactly as asked; raise it again if approvers report missing
that context in practice.

**Not changed**: `Documents with Shared Access`'s per-recipient "asked by" annotation (a live ACL
listing, a different feature — each recipient chip already only shows one email plus who requested
_that specific_ grant, which is not the same redundancy).

### Comment 10 (partial) — "Decided by" wording

**Not changed** — explicitly deferred to Wei Chien's own decision per the client's note ("Wei Chien
to decide if the word 'Decided by' requires change"). Not actioned either way.

## Confirmed already correct, no change needed

### Comment 9 — Keyword field wording, Bulk Upload & Document Upload

Both forms already show the identical placeholder — `"Enter words, phrases, or names related to
this file to make it easier to find in search."` — confirmed by direct source read. The screenshot's
wording ("Words to help find this document later") is from an earlier, already-superseded build; the
underlying concern (consistency between the two forms) is already satisfied.

### Comment 12 — Duplicated "system administrator" banner

Source has exactly ONE occurrence of this banner's JSX (`Requests.tsx` line ~3144) — confirmed by
grep across the whole file. A component cannot render its own single JSX block twice within one
pass, so the stacked-duplicate screenshot is almost certainly a stale browser tab/cached bundle —
this project has hit that exact false-alarm shape repeatedly (see CLAUDE.md's several "stale tab"
entries). **Action for the client: hard-refresh (clear cache) and re-check after deploying today's
other fixes** before treating this as a live defect.

## Needs a live-site check, not a code change

### Comment 14 — Approver's own upload routed for approval instead of auto-approving

`shouldAttemptSelfApprove` (`shared/selfApprove.ts`) gates the whole feature on ONE thing: the
`CRS Config` list's `autoApproveOwnUpload` row must read exactly `"yes"` (case-insensitive). Absent,
blank, or anything else means OFF, by design (fails CLOSED — a wrong "on" would auto-publish a
document nobody reviewed, which this codebase treats as the more expensive failure than a form
staying manual for a minute).

The rest of the mechanism (`probeFolderApproveAccess` — a live ACL read for `ApproveItems` on the
exact destination folder — and `pollForTagStatus`, confirming `CRS — Apply pending tags` has landed
before self-approving) is fully built and was already verified working earlier this session for a
different account. Reading the code found nothing broken.

**Most likely explanations, cheapest first — please check in this order:**

1. The `autoApproveOwnUpload` config row is not actually set to `"yes"` on the live `CRS Config`
   list (easy to have never been turned on, or to have been reset).
2. Afiq's account genuinely does not hold `ApproveItems` on the SPECIFIC folder he uploaded into
   (e.g. he approves a different unit than the one he filed into that day).
3. Tagging took longer than the ~90-second poll window (a documented, known-slow area — "bulk
   upload tags noticeably slower" was flagged and never investigated) and self-approve correctly
   backed off, leaving the document Pending for a human — not a bug, but worth knowing if it
   recurs often.

## Cannot be fixed in this codebase at all

### Comment 13 — SharePoint's own share-invite email names the approver, not the requester

"The file was shared by Afiq and approved by Kheng Wei but... the email says Kheng Wei invited you."
This is correct, unavoidable behaviour: `performShare` calls `SP.Web.ShareObject` — SharePoint's own
native sharing API — and it runs in the **decider's** browser session (the approver), because that
is who is actually granting access at that moment. SharePoint's own invite email is generated from
whoever made that API call; there is no parameter to make it say a different person's name. Not
fixable without abandoning the native share mechanism entirely (which this project deliberately
never did — see CLAUDE.md's 2026-07-23 "share guard retirement" — a web part cannot intercept or
customise SharePoint's own native Share control).

## Needs a Power Automate change, not code (runbook-shaped, none of this built)

These four items (comment 16, and the client's own three-item list: system admin visibility on
notification emails, suppressing admin CC on plain deletions, and stamping/naming attribution) are
all flow-side. This repository has no access to Power Automate; each needs to be built directly in
the flow designer, the same way the ETag guard was earlier this session.

### Comment 16 — No rejection email sent to the requester

"File sharing - rejected by Kheng Wei but Afiq didn't receive the Reject email notification." This
is `CRS — Notify request activity`'s own condition logic — needs opening in the designer to check
whether its trigger condition actually covers a `Rejected` outcome for Share requests, or only
`Approved`. Cannot be diagnosed further without seeing the flow's exported definition.

### Item 2 — System admin should receive everything an approver would receive

Needs a CC/BCC (or a parallel send) added to whichever notification flow(s) currently email only the
approver, addressed to the system admin identity. Needs deciding: which admin, and by what identity
(an individual's email, or a role-based distribution). Not `gdc@sdguthrie.com` itself — CC'ing the
account that IS the sender would be unusual; more likely a named human admin.

### Item 3 — System admin should NOT get every deletion email, only the approver's own confirmation

The opposite of item 2, scoped specifically to deletion notifications: suppress the general CC for
deletion-related emails, keep it only for the one email an approver would themselves receive as
confirmation. Needs the same flow(s) as item 2, with an extra condition specifically for
`RequestType eq 'Deletion'`.

### Item 4 — Attribution fixes

Three separate asks, likely three separate flow edits:

- **"Every Modified By on Approved files should be gdc@sdguthrie.com"** — Auto-route's/HC Auto
  Route's own stamp action currently restamps `Editor`/`Modified By` to the individual approver on
  approval; needs changing to stamp the proxy account instead.
- **"Every Audit Log Who column should show the name not email — the request is showing email"** —
  the request-related `Create_item` actions (in `CRS — Execute approved deletion` and
  `CRS — Notify request activity`) derive `ActorName` by taking the text before `@` in an email
  address (`first(split(email, '@'))`) rather than a real display name — this needs an added lookup
  action (e.g. a SharePoint user-profile or site-users lookup by email) before writing `ActorName`.
- **"For the who section it should show 3 people — uploader, approver, system admin gdc"** — needs
  deciding exactly which audit rows this applies to and what the three-name format should look like,
  before it can be built. Not started.

## Deliberately not investigated in this batch

Given the volume, nothing here was deployed or tested live — the usual next step (build → deploy →
client re-tests) applies to every item above marked "Fixed".

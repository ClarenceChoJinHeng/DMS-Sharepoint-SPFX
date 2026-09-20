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

## ✅ CODE-SIDE ADDENDUM, 2026-09-20 — "show names, not emails" was broader than the Audit Log alone

Client returned to this exact area with fresh screenshots (comment 16's "Reject email not received",
plus the Who column's inconsistency) and asked: (a) investigate the missing rejection email, (b) make
the "Who" column name each decider (PIC/Approver/system admin), (c) show a NAME rather than an email
on every column that names a person, and (d) stamp `Modified By` as gdc on approved library items.

**(c) turned out to be partly buildable in code after all** — the raw-email displays were not
confined to the Audit Log's Who column; `Requests.tsx` (Requested by / Decided by / Revoked by /
"asked by" / "approved by") and `MySubmissions.tsx` (Approved By, the Requests tab's "by …" line, the
detail view's decision line) all showed the same raw addresses. **Fixed**: `nameFromEmail` — the
function `AuditLog.tsx` already had for exactly this — moved to a new shared module,
`src/shared/displayName.ts` (pure, tested), and applied at every one of those sites. DISPLAY ONLY:
the stored email fields, search/filter comparisons, `writeAudit` calls and the CSV export are
untouched everywhere — only the on-screen label changed.

**A second, related defect fell out of the same screenshot**: the proxy account rendered three
different ways across adjacent Audit Log rows — `Guthrie Document Centre` (correct), the raw
`gdc@sdguthrie.com` (stored literally as `ActorName` by one flow action), and a bare `gdc` (stored
the same way by another). **Fixed with `canonicalServiceAccountName`** in the same module — it
recognises `gdc`/`crs` in any of those shapes (bare alias, full address, mixed case) and returns the
real display name, checked BEFORE `ActorName` in the Audit Log's Who cell and before the email-guess
everywhere else. This closes the inconsistency without touching a single flow, and cannot be
re-broken by a fourth action storing the value a fourth way tomorrow. `crs@sdguthrie.com` is kept in
the map (not deleted) for the retired proxy's historical rows, which this append-only log will carry
for as long as it exists.

**This narrows, but does not close, item 4's middle bullet below** — the client no longer needs a
flow-side directory lookup just to stop the Who column showing a bare address; a real person's email
now reads as a reasonable guessed name (`goh.kheng.wei@sdguthrie.com` → "Goh Kheng Wei") and the two
service accounts always read correctly. A genuine directory lookup (Office 365 Users / user-profile
by email) would still be MORE accurate for a real person than the dot-splitting guess — worth
keeping as a "nice to have", not the blocker it was before this fix.

**Verified**: `tsc --noEmit` clean, `npx heft test --clean` → 1991/1991 passing (10 new tests on the
new module), lint at the documented pre-existing 43-warning baseline, no new categories on any of the
four touched files.

## Needs a Power Automate change, not code (runbook-shaped, none of this built)

These items (comment 16, and the client's own three-item list: system admin visibility on
notification emails, suppressing admin CC on plain deletions, and stamping/naming attribution) are
all flow-side. This repository has no access to Power Automate; each needs to be built directly in
the flow designer, the same way the ETag guard was earlier this session.

### Comment 16 — No rejection email sent to the requester

"File sharing - rejected by Kheng Wei but Afiq didn't receive the Reject email notification."

⚠ **THIS IS NOT A BUG — IT IS A CONFIRMED, DELIBERATE GAP.** Per
`docs/superpowers/specs/2026-09-09-request-notification-flow-runbook.md` §1: only FOUR of the six
request-lifecycle emails were built (`ShareRequest`, `ShareApproved`, `DeleteRequest`,
`DeleteApproved`). The two REJECTED templates (8 and 11) were explicitly scoped OUT on 2026-09-09
("lets build the four emails now" — deferred, not cancelled), and `EventKind`'s Compose sends
**both** `ShareRejected` and `DeleteRejected` straight to `'Skip'` on purpose — the same rule §3 of
that runbook states outright: *"A rejected requester is told nothing at all... say this to the
client rather than letting them discover it."* Afiq not receiving anything is the flow doing exactly
what it was built to do; it was never built to notify a rejection.

**To build it** (§9 of the runbook already has the two-edit shape for the sibling `Failed` case —
apply the same pattern here):
1. In `EventKind`'s Compose, change the innermost branch from `'Skip'` to
   `if(equals(triggerOutputs()?['body/RequestType'],'Share'),'ShareRejected','DeleteRejected')`.
2. Add two branches to the email switch (or two new Conditions off `NeedsApprover`'s False side,
   alongside the two Approved emails already there), each addressed to `RequestedBy`, using
   templates 8 and 11 verbatim from
   `docs/superpowers/specs/2026-08-28-email-bundling-and-templates-design.md`:

   **Template 8 (Share rejected):**
   ```
   Subject: Rejected File Sharing Request: [File Name]

   Dear [PIC Name],

   Your request to share [File Name] to [Recipient Email] has been rejected.
   Document Name: [File Name]
   Status: Denied
   Reason/Comments: [Approver Reason]

   Thank you,
   [PIC Name]
   ```

   **Template 11 (Deletion rejected):**
   ```
   Subject: Rejected File Deletion Request: [File Name]

   Dear [PIC Name],

   Your request to delete [File Name] has been rejected.
   The file will remain in its current location.
   Document Name: [File Name]
   Status: Denied
   Reason/Comments: [Approver Reason]

   Thank you,
   [Approver Name]
   ```
   ⚠ Template 11's own sign-off says `[Approver Name]` while addressed to the PIC — flagged in the
   design doc §11.4 as "almost certainly a slip"; confirm with the client before building it as-is.

   Token sources are unchanged from §11.3 of the runbook: `[File Name]` = `ItemName`,
   `[PIC Name]` = `first(split(RequestedBy,'@'))`, `[Recipient Email]` = `ShareWith`,
   `[Approver Reason]` = `DecisionNote`. **No new Compose is needed** — `EventKind`'s existing branch
   structure already routes anything landing here to the same `RequestedBy`-only path the two
   Approved emails already use, with no approver lookup required.
3. `Skip` still catches `Revoked` and `Cancelled` — do not remove the fall-through entirely, only
   the two `Rejected` leaves.

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

Three separate asks — the middle one is now NARROWED by the code-side addendum above, the other two
are unchanged and still flow-only:

- **"Every Modified By on Approved files should be gdc@sdguthrie.com"** — Auto-route's/HC Auto
  Route's own stamp action currently restamps `Editor`/`Modified By` to the individual approver on
  approval (`triggerOutputs()?['body/Author/…']`, per CLAUDE.md's "THE `Approved` AUDIT ROW IS
  WRITTEN BY AUTO-ROUTE" and the `ApprovedBy` stamp sections); needs changing that ONE value to the
  proxy account's claims instead. ⚠ Do **not** touch `Author`/`Created By` on the same action — that
  is deliberately kept as the ORIGINAL UPLOADER (`dms-uploader-column-is-author` in project memory:
  "Uploader column must be Author, never Editor"), and changing it would erase who actually filed the
  document. Needs the same edit in BOTH `Auto-route` and `HC Auto Route` — every HC clone in this
  project has shipped with exactly one library/field reference nobody swapped.
- ~~**"Every Audit Log Who column should show the name not email"**~~ — **narrowed, see the addendum
  above.** The code-side fix already makes every screen show a reasonable name for any address,
  and always shows the CORRECT canonical name for the two known service accounts, regardless of what
  a flow writes. A directory lookup (Office 365 Users' "Get user profile (V2)" by email, written into
  `ActorName` before the `Create item` in `CRS — Execute approved deletion` and
  `CRS — Notify request activity`) would still be a genuine improvement for REAL people — the guessed
  name is a good approximation of `first.last@domain`, not a directory-verified one — but it is no
  longer blocking anything.
- **"For the who section it should show 3 people — uploader, approver, system admin gdc"** — needs
  deciding exactly which audit rows this applies to and what the three-name format should look like,
  before it can be built. Not started; unaffected by the code-side addendum, since that only fixes
  display of a SINGLE stored name per row, not adding a second and third name to one row.

## Deliberately not investigated in this batch

Given the volume, nothing here was deployed or tested live — the usual next step (build → deploy →
client re-tests) applies to every item above marked "Fixed".

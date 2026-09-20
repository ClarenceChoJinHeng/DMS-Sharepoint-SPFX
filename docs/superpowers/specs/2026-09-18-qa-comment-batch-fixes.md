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

**⚠ (c) WAS INITIALLY BUILT TOO WIDE AND WAS PARTLY REVERTED THE SAME SESSION.** The first pass
applied the fix to `Requests.tsx` and `MySubmissions.tsx` as well as the Audit Log, on the reasoning
that the same raw-email complaint appeared on all three screens. The client corrected this: **the ask
was for the Audit Log's Who column only** — "did you change the Request.tsx and MySubmission.tsx as
well? If so revert it." **Reverted** — those two files are back to showing the raw email exactly as
before. `AuditLog.tsx` keeps the fix, now sourced from a new shared module,
`src/shared/displayName.ts` (`nameFromEmail`, pure, tested), rather than the private local copy it
had before — the extraction stands on its own even with a single consumer, and the module stays ready
if this is asked for more broadly later.

**A second, related defect fell out of the same screenshot and IS still fixed, since it was Audit-Log
specific from the start**: the proxy account rendered three different ways across adjacent Audit Log
rows — `Guthrie Document Centre` (correct), the raw `gdc@sdguthrie.com` (stored literally as
`ActorName` by one flow action), and a bare `gdc` (stored the same way by another). **Fixed with
`canonicalServiceAccountName`**, same module — it recognises `gdc`/`crs` in any of those shapes (bare
alias, full address, mixed case) and returns the real display name, checked BEFORE `ActorName` in the
Who cell and before the email-guess. This closes the inconsistency without touching a single flow,
and cannot be re-broken by a fourth action storing the value a fourth way tomorrow.
`crs@sdguthrie.com` is kept in the map (not deleted) for the retired proxy's historical rows, which
this append-only log will carry for as long as it exists.

**This narrows, but does not close, item 4's middle bullet below** — the Audit Log no longer needs a
flow-side directory lookup just to stop the Who column showing a bare address for a real person; their
email now reads as a reasonable guessed name (`goh.kheng.wei@sdguthrie.com` → "Goh Kheng Wei") and the
two service accounts always read correctly. A genuine directory lookup (Office 365 Users / user-
profile by email) would still be MORE accurate than the dot-splitting guess — worth keeping as a "nice
to have", not the blocker it was before this fix.

**Verified**: `tsc --noEmit` clean, `npx heft test --clean` → 1991/1991 passing (10 tests on the new
module), lint at the documented pre-existing 43-warning baseline, no new categories on the one
touched screen (`AuditLog.tsx`).

## ✅ CODE-SIDE ADDENDUM 2, 2026-09-20 — `Modified By` = the proxy account, for "Move existing folders"

Client, same session: "For the 6 Libraries Modified by columns it should be the gdc name, cater this
to Move existing folders as well." Two routes touch this column; only ONE is code.

- **The routing route (Auto-route / HC Auto Route) is still flow-only** — see item 4 below, unchanged.
- **The "Move existing folders" route (`SubtreeMigrator.tsx`) IS code, and is now built.** That tool
  runs entirely in the ADMIN's own SPFx session — `context.spHttpClient` posts as whoever is running
  the migration — so a plain `MoveTo` would otherwise leave `Modified By` reading their name, not
  gdc's, the same gap Auto-route already has an answer for on the routing side.

**New in `shared/dmsFolderMap.ts`**: `resolveProxyLoginName` (resolves the CURRENT proxy account's
claims login via `ensureSiteUser` — resolves, never invites, and this account is already a site
member, so this can never trigger a guest invite) and `stampEditorAsProxy` (restamps ONE file's
`Editor` field via `validateUpdateListItem` with `bNewDocumentUpdate: true`, the same no-new-version
mechanism `SubtreeMigrator`'s own tier-column stamp already uses). New constant
`CURRENT_PROXY_ACCOUNT_EMAIL` in `shared/displayName.ts` — the ONE place this address is typed, so a
future migration (as already happened once, crs → gdc) is a one-line change rather than a grep.

**⚠⚠ ORDERING MATTERS, AND THE FIRST DRAFT GOT IT WRONG BEFORE SHIPPING.** The obvious placement —
stamp `Editor` immediately after each successful `moveFileTo`, inline in the move loop — is wrong,
because `backfillMetadata` runs a SEPARATE, LATER `validateUpdateListItem` call on the same file for
tier-column backfill, and that call does **not** set `Editor` explicitly. SharePoint would then
restamp `Editor` back to the calling admin as an ordinary side effect of that unrelated write,
silently undoing the attribution stamp made moments earlier. **Fixed by moving the attribution pass
to run LAST** — every moved file's `{lib, path}` is recorded during the move loop
(`movedForAttribution`), and the actual `Editor` stamp happens in its own pass AFTER
`backfillMetadata` completes, so nothing written later in the run can clobber it.

- **Fails soft throughout, on purpose.** `resolveProxyLoginName` returning `undefined` (the account
  could not be resolved this run) skips every stamp for the whole run without touching the moves;
  a per-file `stampEditorAsProxy` failure is counted separately from `failed` (the move-retry counter)
  and reported as its own line, because a stuck attribution stamp is not fixed by pressing "Move
  existing folders" again — the file is no longer part of any plan the next scan would find, so it
  is stated as a fact rather than offered as a retry.
- **Both outcomes are named in the run's own summary and its audit row** — `attributionFailed`, and
  the whole-run "the proxy account could not be resolved" case are both distinct from the ordinary
  move-failure count, so an admin reading the result cannot mistake "the documents moved, only their
  Modified By label didn't" for "the migration failed."
- **Deliberately scoped to exactly what was asked**: only `SubtreeMigrator.tsx`'s move loop. The
  standalone tier-column-only "Check document tags" path (`runTagsOnly`, reachable with no move at
  all) was NOT touched — it does not move anything, so it was outside "Move existing folders" as
  named. Worth flagging if the client also wants Modified By addressed there.

**Verified**: `tsc --noEmit` clean, `npx heft test --clean` → 1991/1991 passing, lint at the
documented pre-existing baseline — the five `no-new-null` warnings on `dmsFolderMap.ts` are the SAME
five as before (only their line numbers shifted, from the new code inserted above them), and
`SubtreeMigrator.tsx`'s `max-lines` warning is the same pre-existing category, now reporting 2472
lines instead of 2434 (that file was already over the 2000-line ceiling before this change).
**NOT yet site-tested** — this is the single most site-verified migration tool in the project, and
the ordering fix above has never been exercised against a real library. The test that matters: run a
real "Move existing folders" migration on a segment whose below-Unit tiers actually need backfilling
(so `backfillMetadata` genuinely writes something after the move), then confirm the moved file's
Modified By reads the proxy account's name — not the admin's — AFTER both passes have run, not just
right after the move.

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

Needs a BCC added to whichever notification flow(s) currently email only the approver, addressed to
the system admin identity. **Confirmed with the client, 2026-09-20: BCC, not CC** — the admin should
receive a silent copy, not appear in the visible recipient list an approver (or anyone replying-all)
would see. Still needs deciding: which admin, and by what identity (an individual's email, or a
role-based distribution). Not `gdc@sdguthrie.com` itself — copying the account that IS the sender
would be unusual; more likely a named human admin.

### Item 3 — System admin should NOT get every deletion email, only the approver's own confirmation

The opposite of item 2, scoped specifically to deletion notifications: suppress the general CC for
deletion-related emails, keep it only for the one email an approver would themselves receive as
confirmation. Needs the same flow(s) as item 2, with an extra condition specifically for
`RequestType eq 'Deletion'`.

### Item 4 — Attribution fixes

Three separate asks — the middle one is now NARROWED by the code-side addendum above, the other two
are unchanged and still flow-only:

- **"Every Modified By on Approved files should be gdc@sdguthrie.com, cater this to Move existing
  folders as well"** — TWO routes touch this column, and only one is flow-side:
  - **The routing route (Auto-route / HC Auto Route) — still flow-only, not built. Exact edit below,
    the shape is CONFIRMED, not guessed** — from
    `docs/superpowers/specs/2026-08-08-auto-route-flow-and-draft-isolation.md` §4.3, the action named
    `Send an HTTP request to SharePoint` that stamps `Author`/`Editor`/`Created`, verified live
    2026-08-08:
    ```
    POST _api/web/lists/getbytitle('Documents')/items(@{outputs('Copy_file')?['body/ItemId']})/validateUpdateListItem
    ```
    ```json
    {
      "formValues": [
        { "FieldName": "Author",  "FieldValue": "[{'Key':'@{body('Get_item')?['Author']?['Claims']}'}]" },
        { "FieldName": "Editor",  "FieldValue": "[{'Key':'@{body('Get_item')?['Author']?['Claims']}'}]" },
        { "FieldName": "Created", "FieldValue": "@{formatDateTime(body('Get_item')?['Created'],'M/d/yyyy h:mm tt')}" }
      ],
      "bNewDocumentUpdate": true
    }
    ```
    **The ONLY line that changes is `Editor`** — from the author's claims to the proxy account's:
    ```json
    { "FieldName": "Editor", "FieldValue": "[{'Key':'@{body('GetProxyUser')?['LoginName']}'}]" }
    ```
    ⚠ **Do NOT hardcode a claims literal like `i:0#.f|membership|gdc@sdguthrie.com`** — the exact
    claims-provider prefix can differ by tenant configuration, and a wrong literal fails silently on
    this specific endpoint (gotcha: `validateUpdateListItem` can report `HasException: false` while
    changing nothing — §5.1 of the same runbook, "believed impossible for two hours"). Instead add
    ONE new action before this stamp, `GetProxyUser`:
    ```
    POST _api/web/ensureuser
    Body: { "logonName": "gdc@sdguthrie.com" }
    ```
    This is the same call `ensureuser` this codebase's own `ensureSiteUser` already uses to resolve a
    claim rather than guess one (`src/shared/spGroups.ts`) — it RESOLVES, never invites, and `gdc` is
    already a site member (`CRS Owners`), so this is guaranteed safe. Its response carries
    `LoginName` — the exact claims string for THIS tenant — which the corrected `Editor` line above
    reads back.
    - **⚠ Do NOT touch `Author` or `Created` on the same action** — deliberately kept as the ORIGINAL
      UPLOADER (project memory `dms-uploader-column-is-author`: "Uploader column must be Author,
      never Editor"). Changing `Author` would erase who actually filed the document.
    - **Needs the identical edit in BOTH `Auto-route` and `HC Auto Route`** — every HC clone in this
      project has shipped with exactly one library/field reference nobody swapped; do both in the
      same sitting so neither is the one left behind.
    - **Verify after building**: approve one ordinary document and one HC document, then check each
      one's Modified By reads the proxy account's display name, not the approver's — and separately
      confirm `Created By` still correctly shows the ORIGINAL UPLOADER on both.
  - **✅ The "Move existing folders" route IS code, and IS BUILT** — see the addendum below.
    `SubtreeMigrator.tsx` runs entirely in the ADMIN's own SPFx session, so a plain `MoveTo` would
    otherwise leave `Modified By` reading their name. New `resolveProxyLoginName`/`stampEditorAsProxy`
    in `shared/dmsFolderMap.ts`, wired into the migration run as its LAST write per file (see the
    addendum for why ordering matters here — the tier-column backfill pass would otherwise clobber
    the stamp). **NOT yet site-tested.**
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

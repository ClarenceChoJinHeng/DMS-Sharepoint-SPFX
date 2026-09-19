# Session handoff — 2026-09-20

Written mid-session, per the client's standing instruction from 2026-09-19: keep this current AS
work happens, not only at the end. The session is large enough now that losing track mid-fix is the
exact risk being guarded against — read this before doing anything else if picking up cold.

## What happened, in order

1. Picked up from `docs/2026-09-19-session-handoff.md`. That session's bulk-upload tagging fix
   (`buildLevelFormValues` call added to `BulkUpload.tsx`) was already done and verified by `tsc`/the
   test suite, but explicitly **not yet confirmed live** — this session's job was watching a real
   bulk upload go all the way through `CRS — Apply pending tags` and confirming.
2. Client uploaded a 5-file bulk batch ("Slide Deck") and confirmed the CORE FIX WORKS: Business
   Segment/Department/Unit landed correctly on the live documents, not just in the tag payload.
3. Client then found a NEW bug while checking: My Submissions' Submissions tab showed the SAME 5
   files TWICE — once correctly tagged (`[Bulk]`, a real `SUB-...` reference), once as a phantom
   "Grouped by folder and date — uploaded before submissions were recorded" entry.
4. **Diagnosed and FIXED** — `src/shared/submissionRecords.ts`. Root cause: since the 2026-09-18
   tag-by-proxy migration, `SubmissionFileId`/`SubmissionId` are written onto the live document by
   `CRS — Apply pending tags` on a POLL, not at upload time. So for the settling window between
   upload and the flow completing, the live document row (blank `SubmissionId`) and a synthetic
   "settling" placeholder (built from the `CRS Submissions` record, which already has the real
   reference) both exist independently — one lands in `groupSubmissions`' real-reference group, the
   other falls into its "no reference" fallback. Two rows, one physical batch.
   - Fix: added `indexLiveByUniqueId` + a SECONDARY join in `mergeRecords`, tried only when the
     primary `SubmissionFileId` join misses. A document's `UniqueId` exists from the moment of
     creation (unlike the custom stamp), so it closes exactly this gap — and it cannot reopen the
     danger the original "never join on UniqueId" rule exists to prevent (that risk is about a
     ROUTED/approved file, where `UniqueId` changes and the fallback naturally stops matching at
     exactly that point).
   - Also fills `submissionId`/`batchId` on the merged row from the record when the live row's own
     columns are still blank, so `groupSubmissions` groups it correctly the moment either join hits.
   - 4 new tests, one existing test's premise updated (it asserted "never resolves via UniqueId,
     full stop" — now narrower and correctly scoped). Full suite **1968/1968 passing**, no new lint
     warnings.
5. **CONFIRMED LIVE** by the client on a fresh batch ("Licenses", 5 files, built/deployed by the
   client themselves) — no duplicate entry. Fix works.
6. Client found a SECOND, related bug while checking that same batch: the per-file/batch DETAIL
   VIEW (Confidentiality, Legally Privileged, Document Type, Year, etc.) showed stale/blank values
   even for files that WERE correctly tagged live, and pressing the on-page "↻ Refresh" button did
   NOT fix it — only leaving the page and coming back (a full remount) did.
7. **Diagnosed and FIXED** — `src/webparts/mySubmissions/components/MySubmissions.tsx`. The batch
   detail view's re-fetch guard (`sig`, in the `useEffect` around line ~1710, keyed off `mergedKey`
   = `library#itemId`) never changes while a file sits Pending during the settling window — only an
   APPROVAL (which moves the file to a new library/id) would change it. So `onRefresh` correctly
   reloaded the main `rows` list, but the guard saw an unchanged signature and skipped re-fetching
   the per-file `FieldValuesAsText` data, leaving the detail view stuck on its very first (pre-tag)
   snapshot until a full remount reset the guard.
   - Fix: folded `submissionId` and `record?.tagStatus` into the signature — both flip from blank to
     real values the instant tagging finishes, since `TagPayload` is applied as one atomic write.
   - `tsc --noEmit` clean, full suite still **1968/1968**, no new lint warnings (this file's
     pre-existing `max-lines` warning grew only from the added comment).
8. Client asked about hiding Document Date/Vendor-Customer/Remark on the detail panels for
   Bulk-sourced documents, since Bulk Upload collects no input for them. Flagged that this would
   reverse a DELIBERATE 2026-09-11 client decision ("always show all ten fixed fields, dash if
   blank, for consistency" — see CLAUDE.md, `documentDetails.ts`'s `FILE_FIXED_FIELDS`/
   `fixedRow`). **Client decided not to pursue this** — "client didn't complain so lets not go
   there." No code change made.
9. Client re-ran the bulk upload test as an HC upload (same "Slide Deck" path, HC library).
   **CONFIRMED LIVE, BOTH FIXES TOGETHER**: all 5 files tagged correctly this time (Business
   Segment/Department/Unit/Confidentiality=Highly Confidential/LegallyPrivileged=Yes/Document
   Type/Year/Keyword all correct, no untagged stragglers), My Submissions showed exactly ONE entry
   for the batch (no duplicate), and pressing "↻ Refresh" alone — no page reload needed — correctly
   picked up the tagged values.
10. **Open, unexplained observation** from the client: bulk upload tagging appears to take
    noticeably longer than a normal Form.tsx upload's tagging. NOT investigated. Not currently
    blocking anything; worth digging into only if the client raises it again or it becomes a real
    problem (e.g. `CRS — Apply pending tags` processing multiple items sequentially vs. Form.tsx's
    single-item path, or simply more files per run).
11. Client uploaded via the NORMAL upload form (`Form.tsx`), a genuinely complex case: ONE
    submission, TWO sets, spanning an ordinary Business Segment (`GHO › COSEC › GUTHRIE › 2025 ›
    Legal Opinion`) AND a Group-Led Project (`PCAR › EUB › BMW › vv › 2024 › Licenses`), each set
    with a mix of HC and non-HC files (2 of 5 per set). **CONFIRMED WORKING**: My Submissions
    correctly showed one combined entry ("... + 1 more destination", 10 files), both HC and non-HC
    files tagged into the right libraries, Group-Led Project's label correctly read
    "Group-Led Project" rather than "Segment".
12. Asked to narrow `scripts/check-tagging-status.js` down to specific folders instead of scanning
    the whole `CRS Submissions` list every time. **BUILT**: added a `PATH_FILTER` config array
    (case-insensitive substring match on `ItemPath`) — matches BOTH the HC and non-HC library copy
    of a folder automatically, since only the library segment differs, never the tail of the path.
    Verified against a synthetic harness before handing off.
13. Running the narrowed script on the item-11 upload flagged 5 of 10 files as "tagged but blank"
    (`ProjectName`). **Diagnosed as a FALSE POSITIVE, not a bug** — `ProjectName`/`Vendor_x002f_
    CustomerName`/`Remark`/`Keyword` are all optional per `missingForFile` in `Form.tsx` (confirmed
    by reading that function directly), and the composed filenames themselves proved it (2-segment
    names for the set missing Project Name vs. 3-segment names for the set that had it filled in).
    **FIXED the script**: added an `OPTIONAL_FIELDS` list and a separate, clearly-labelled "BLANK,
    BUT ONLY OPTIONAL FIELDS — NOT A PROBLEM" bucket, excluded from the summary's problem count.
    Verified against a synthetic harness (required-field-blank still flagged; optional-field-blank
    correctly separated). Re-run on the real data: **0 rows need attention, 10 confirmed clean.**
14. Client then re-uploaded the SAME 10 files to the SAME two locations to test the Replace/clash
    flow. Pressed "Yes" on the "Replace Existing Files" dialog (all 10 files listed across both
    sets). **Two findings:**
    - **Took ~8 minutes** (client: "was never this slow"). **ROOT-CAUSED**: since the 2026-09-18
      tag/approve-by-proxy migration, replacing an existing pending file goes through
      `deleteClashingDraftByProxy` in `Form.tsx` — writes a self-approved `CRS Requests` deletion
      row, then **polls up to 30 seconds per file** waiting for a Power Automate flow
      (`CRS — Execute approved deletion`) to actually recycle it, before falling back to the old
      direct `overwrite=true` (which still works today since `CRS Upload` hasn't had `Edit Items`
      revoked yet — Task 13 of the 2026-09-18 design). With all 10 files clashing, that's up to
      5 minutes of pure polling before every one falls back.
    - **Client found the actual cause directly, live**: `CRS — Execute approved deletion`'s
      SharePoint connection ("GDC Proxy - SharePoint") was **not actually connected** — the flow's
      own page showed a red trigger error ("hasn't been triggered successfully in the last 28
      days") and a completely empty 28-day run history, despite being created and "On" since Sep 19.
      All 10 self-approved `CRS Requests` deletion rows from this test were confirmed stuck forever
      at `Status: Approved`, `Modified` = `Created` (never touched again).
    - ⚠⚠ **SEVERITY, FLAGGED TO THE CLIENT**: Replace has a fallback (direct overwrite), but a
      PLAIN DELETE does not — `writeApprovedDeletionRequest` in `MySubmissions.tsx` (and the
      equivalent in `Requests.tsx`) writes the request and marks the UI as done with **no
      confirmation and no fallback**. Since this flow has never fired even once since creation,
      **every plain "Delete" pressed anywhere in the app since 2026-09-17 has appeared to succeed
      while the actual file was never recycled** — it is still sitting in the library, un-deleted.
      Nothing is *lost* (the opposite problem — nothing was ever removed), but anyone who believes a
      document was deleted since then should not assume it actually is.
15. Client went through **all 23 flows**, reconnecting SharePoint (and Office 365 Outlook where
    needed) to `gdc@sdguthrie.com`. **Reported done.** Re-ran the full test cycle: a fresh Form.tsx
    upload (2 sets, GHO/Legal Opinion + a Group-Led Project, mixed HC/non-HC), then checked My
    Submissions and the SharePoint library.
16. Two new findings from that test, both from the SAME root cause (leftover same-named `test1–
    test5` files sitting in the destination folders from earlier testing, so the fresh upload's
    clash/replace path fired):
    - **"Why am I getting a file-deletion-request email for an upload?"** — **explained, not a bug.**
      A same-named draft already existed in the destination folder, so the upload's clash dialog
      ("Replace Existing Files?") triggered the replace path. Since the 2026-09-18 tag-by-proxy
      change, replace no longer overwrites directly — it writes a self-approved `CRS Requests` row
      (`RequestType: Deletion`, `Status: Approved` from creation) so the proxy account can recycle
      the old draft on the uploader's behalf. `CRS — Notify request activity` fires on ANY
      create/modify of that list regardless of origin, and its `Deletion`/`Approved` branch emails
      **the requester** — which is the uploader themselves, since they're the one who clicked
      replace. Mechanically correct; just an unexpected side effect of clicking "replace" during an
      ordinary upload. Worth a client decision on whether self-generated replace-clash deletions
      should suppress this notification (the runbook itself flags "nothing here dedupes" as an open
      item — this is the same class of gap).
    - **⚠⚠ SUSPECTED SERIOUS RACE — "the file is actually deleted."** My Submissions showed the
      newest submission's `test1` flip-flopping between "1 deleted" and briefly "replaced" across
      refreshes, and the client confirmed the file is genuinely gone from the library, not just a
      display artifact. **Leading theory, NOT YET CONFIRMED against a live flow run:**
      - `deleteClashingDraftByProxy` (Form.tsx) polls up to 10×3s (~30s) for the old draft's GUID to
        404 (confirming the proxy flow recycled it). If it doesn't see that within 30s, it gives up
        and falls back to a **direct `Files/Add(overwrite=true)`** on the same name — which
        SharePoint applies IN PLACE, keeping the SAME list item / SAME UniqueId as the file being
        replaced.
      - `CRS — Execute approved deletion` doesn't know the client gave up. It still eventually
        processes the `CRS Requests` row it was handed and calls `recycle()` on the GUID it was told
        to target — regardless of how long that takes.
      - If the flow's actual run happened later than the client's 30-second patience (very plausible
        right after reconnecting 23 flows at once — cold starts, a backlog, or simple throttling),
        the fallback overwrite will have ALREADY landed the new content onto that same GUID by the
        time the flow finally fires. The delayed-but-genuinely-successful recycle then deletes the
        file the uploader just replaced-in, minutes after the upload looked complete.
      - This fits every symptom: the notification email (proves the request was real), the
        multi-minute delay across 10 clashing files (each burning up to 30s before falling back),
        the flip-flop in My Submissions (briefly resolves live via the same-UniqueId fallback path,
        then reverts to "deleted" once the delayed flow's recycle actually lands and the item is
        genuinely gone — well past the 5-minute settling grace by that point, so it correctly reads
        "deleted" rather than "settling"), and the client's own direct confirmation that the file is
        gone.
      - **NOT YET CONFIRMED.** The decisive check: open `CRS — Execute approved deletion`'s 28-day
        run history and see whether there is a run for `test1`'s request, and if so, what time it
        actually fired relative to the upload. A run landing noticeably after the ~30s mark would
        confirm the race.
      - **The real fix, once confirmed:** before falling back to a direct overwrite, cancel or
        delete the self-approved `CRS Requests` row that was just written, so a late-arriving flow
        run has nothing left to act on. Widening the poll window alone does not close this — any
        fixed timeout can still be beaten by a large enough backlog on the flow side.
      - **No code change made yet** — deliberately holding off until the run-history check confirms
        (or rules out) this theory, since a wrong fix here risks making the replace path worse, not
        better.

## Current uncommitted state — `git status` shows nothing committed this session either

Modified (carried over from 2026-09-19, plus this session's two fixes on top):
- `CLAUDE.md` — this pointer + entries for both fixes above.
- `docs/superpowers/specs/2026-09-19-service-account-migration-crs-to-gdc-runbook.md` — untouched
  this session (carried over).
- `src/shared/naming.ts` + `naming.test.ts` — untouched this session (carried over from 2026-09-19;
  not investigated further, see that handoff's own notes if picking this thread back up).
- `src/shared/spSubmissionRecords.ts` — untouched this session (carried over).
- `src/shared/submissionRecords.ts` + `submissionRecords.test.ts` — **THIS SESSION's fix #4 above**
  (the `indexLiveByUniqueId` settling-window join).
- `src/webparts/bulkUpload/components/BulkUpload.tsx` — untouched this session (carried over; the
  `buildLevelFormValues` fix from 2026-09-19).
- `src/webparts/form/components/Form.tsx` — untouched this session (carried over; the
  `pollForTagStatus` self-approve change from 2026-09-19).
- `src/webparts/mySubmissions/components/MySubmissions.tsx` — **THIS SESSION's fix #7 above** (the
  batch-signature refresh fix), on top of the 2026-09-19 `settling` record-state display change.

New, untracked (carried over, untouched this session):
- `docs/2026-09-19-session-handoff.md`
- `scripts/check-audit-recording.js` — **still not run by the client.**
- `scripts/check-tagging-status.js` — used earlier this session to help confirm tagging status
  during diagnosis.

⚠ **Nothing above is committed.** The working tree now holds THREE sessions' worth of unrelated
in-flight threads (2026-09-18/19 tag-by-proxy work, 2026-09-19 bulk-upload fix, this session's two
My Submissions fixes) with no commit boundary between any of them. Splitting this into sensible
commits is still an open item — see below.

## Build status

Both of THIS session's fixes (steps 4 and 7 above) have been **built and deployed by the client
themselves** (self-served, `npm run build` → upload/update in Site Contents, to save tokens) and are
**confirmed working live** — see steps 5 and 9. No further build/deploy action is pending for these
two specifically. Anything the client does next (the Form.tsx HC/Group-led-Project pass) runs
against this already-deployed state.

## Open items — what to pick up next

0. **⚠⚠ TOP PRIORITY, NARROWED — apply the flow-side runbook (item 31 above).** The recycle bin
   check (item 27) already confirmed the file is recoverable there, and the client chose and the
   CODE side of the fix is BUILT (item 30). The only remaining action is the Power Automate edit in
   `docs/superpowers/specs/2026-09-20-etag-guard-execute-approved-deletion-runbook.md` — follow it
   step by step, then test per its §7 before considering this closed.

1. **~~Confirm or rule out the delayed-recycle race~~ — UN-DOWNGRADED, see item 0.** The earlier
   "did not reproduce in the latest 10-file test" note (items 17-18) is still true FOR THAT
   particular batch, but item 24-26 above found it DID recur on a different upload
   (`test1-test1-test1-20092026.xlsx`, uploaded across 17:18/17:46/18:54) discovered via the
   evidence-collection script. Do not read the earlier "downgraded" note as still current.
   Open `CRS — Execute approved deletion`'s 28-day run history for `test1`'s request from this
   session's test and check WHEN it actually ran relative to the upload. If it ran after the ~30s
   poll gave up, that confirms: the client-side fallback (`Files/Add(overwrite=true)`, which keeps
   the SAME UniqueId) and the now-genuinely-working proxy flow (which recycles that same UniqueId
   once it eventually fires) can both act on one identity — the flow's late-but-real recycle then
   deletes the content the fallback just wrote.
   - ⚠⚠ **CORRECTION, 2026-09-20 (next session): the handoff's own proposed fix — "cancel/delete the
     self-approved `CRS Requests` row before falling back" — DOES NOT WORK, and should not be built
     as written.** Read the flow's own runbook
     (`docs/superpowers/specs/2026-09-17-execute-approved-deletion-flow-runbook.md` §3/§5): both the
     True-branch recycle action and the Failure-branch update action read **only**
     `triggerOutputs()` — the payload frozen at the moment the trigger fired — and neither ever
     re-reads the `CRS Requests` item's live state during execution. So once a run has been
     triggered and queued (even if its actual execution is delayed), modifying or deleting the row
     afterward has nothing left to act on — there is no lookup of the row's current state anywhere
     in the flow's logic for it to see the cancellation.
   - **A second, more fundamental point, independent of confirming this specific incident:**
     "When an item is created or modified" is a **polling** trigger, not instant/push-based — the
     standard SharePoint connector typically checks on an interval (commonly ~1 minute, sometimes
     longer under throttling), not sub-30-seconds. The client-side poll budget in
     `deleteClashingDraftByProxy` (`Form.tsx`) is 10×3s ≈ 30s. So the fallback to
     `Files/Add(overwrite=true)` — same identity, same `UniqueId` — will plausibly fire **on most
     replace-clashes**, independent of whether the connection is healthy or the reconnection sweep
     "cold start" has settled. This is not a one-off from reconnecting 23 flows at once; it can
     recur any time the flow's actual execution takes longer than 30s, which a polling trigger makes
     common.
   - **Real fix options, once the client is ready (needs a decision, not yet made — asked, held off
     pending the run-history check):**
     1. **Safest, code-only:** on timeout, refuse the upload and tell the uploader to retry shortly,
        instead of falling back to `overwrite=true`. Never risks two writers on one identity.
     2. **Code-only, more work:** upload the replacement under a temporary name on timeout, then
        rename into place once the deletion is later confirmed.
     3. **Flow + code change (structurally correct):** stamp something identifying (e.g. the target
        file's `ETag`/`Modified`) onto the `CRS Requests` row at write time; have the flow re-check
        that value against the live file immediately before recycling, and skip if it has changed
        since. Needs a Power Automate edit alongside the code change.
   - **Do not widen the poll window as a fix on its own** — any fixed timeout can still be beaten by
     a large enough backlog; that was already noted and still holds.
   - **Status as of 2026-09-20 (next session): asked the client whether to (a) treat the race as
     confirmed by the reasoning above and fix it now, or (b) check the run history first — they
     chose to check first — and (c) which of the three fix options above to build — they chose to
     hold off entirely until the check is done. No code changed. Pick this thread up by asking
     whether the run history has been checked yet.**
2. **Decide whether replace-clash auto-deletions should notify anyone at all** (item 16's first
   finding). `CRS — Notify request activity` currently emails the requester for every
   `Deletion`/`Approved` row regardless of whether it was raised on the Requests page or
   auto-generated by an ordinary upload's clash resolution. Mechanically correct, just surprising —
   worth a client decision on whether to suppress notification for this specific origin.
3. **⚠⚠ AUDIT WHETHER ANY DOCUMENT "DELETED" SINCE 2026-09-17 IS ACTUALLY STILL PRESENT.** Direct
   consequence of item 14's severity finding — a plain Delete has no fallback and, until today's
   reconnection sweep, this flow had never fired, so nothing deleted through the app since that
   date was actually recycled. Now that connections are reportedly fixed, this needs re-checking in
   the OTHER direction too (item 1 above) — a genuinely working-but-slow flow is its own risk.
4. **The `CRS Requests` rows stuck at `Status: Approved` from earlier failed tests will NOT
   self-resolve** now the connection is fixed — a create-or-modify trigger only fires on FUTURE
   changes to a row, not retroactively. Decide with the client what to do with them.
5. **Why bulk upload tags slower than Form.tsx** — open question, not investigated, not blocking.
6. **Audit log — still unresolved, carried over from 2026-09-19.** `scripts/check-audit-recording.js`
   has still not been run by the client. Now that flows are reconnected, this is worth checking
   again — per the earlier hypothesis, the SAME connection issue may explain why no `Uploaded`/
   `Approved` audit rows were appearing.
7. **Splitting the accumulated uncommitted work into commits** — now spans three sessions' worth of
   unrelated threads (see the file list above), plus `scripts/check-tagging-status.js`'s two new
   features (`PATH_FILTER`, `OPTIONAL_FIELDS`) from this session. Still not done; decide with the
   client when there is a natural pause.

## Update — later the same session, after checking run history / mailboxes

17. Client checked `CRS — Execute approved deletion`'s 28-day run history directly (screenshot: many
    runs between 02:54–03:00 AM, all "Succeeded", most under 1 second). Couldn't tell from
    Start/Duration/Status alone which run was `test1`'s, so was pointed at `CRS Requests` + the
    `CRS Audit Log` directly instead, to compare the request's `Created` time against the flow's
    actual `Deleted` audit-log time for the same file.
18. **That comparison came back clean for THIS test run — no race observed.** The audit log
    (`CRS Audit Log` / "History of this file" style view) shows, for all 10 test1–test5 files across
    both sets (GHO/COSEC/GUTHRIE/2025/Legal Opinion and PCAR/EUB/BMW/vv/2024/Licenses):
    `Uploaded` → `Deleted` → `Request approved` landing in the **same minute**, repeating in sequence
    from 02:54 to 03:00. No multi-minute gap between a `CRS Requests` row's `Created` time and its
    corresponding `Deleted` audit row. **This does NOT rule out the race under heavier load or a
    slower flow run — it only means it did not reproduce in this specific 10-file batch.** No
    correction needed for this batch; nothing currently stuck at `Status: Approved` forever from it.
19. **⚠⚠ CORRECTED — my own earlier theory about the "gdc sees a deletion request during replace,
    which doesn't make sense" observation was WRONG and should be discarded.** I had guessed it was
    `Requests.tsx` showing gdc every request site-wide because `gdc@sdguthrie.com` satisfies
    `isSystemAdmin` (being in `CRS Owners`). **Never confirmed which screen the client was actually
    looking at before proposing that — and it was wrong.** The client clarified directly: they were
    looking at **gdc's own Outlook mailbox**, not any page in the app. `gdc@sdguthrie.com` is the
    account every notification-sending flow (`CRS — Notify request activity`, the approver-reminder
    flow, etc.) runs as and sends email *from* — so its own mailbox naturally retains a copy of
    every email it has ever sent, plus every bounce-back (NDR) for ones that failed to deliver. That
    is ordinary behavior for a service/sending mailbox, not a UI visibility bug. **No code change is
    needed for this — there was never a real defect here.**
    - **Confirms the routing itself IS correct**, from the other direction: the client separately
      checked `crs@sdguthrie.com` (a genuine **approver** mailbox) and confirmed it received *only*
      "file requires approval" emails — no deletion-request notifications at all. Matches the design:
      deletion-approved notices go to the *requester*, never to an approver.
20. **Re-confirmed, not new: the Exchange tenant mail-flow block from 2026-09-10 is still active.**
    A fresh NDR bounce for "Approved File Deletion Request: test5-test5-20092026.xlsx" (addressed to
    `clarence@trinergydigital.com`) reads: *"A custom mail flow rule created by an admin at
    simedarbyplantation.onmicrosoft.com has blocked your message. You are not allowed to send emails
    to external recipients."* This is the exact, already-documented issue in CLAUDE.md
    ("⚠⚠ EXCHANGE ONLINE BLOCKS EVERY OUTBOUND EMAIL TO AN EXTERNAL ADDRESS", 2026-09-10) —
    `clarence@trinergydigital.com` is external to SD Guthrie's real tenant
    (`simedarbyplantation.onmicrosoft.com`). **No code fix is possible for this; it needs an Exchange
    admin exception on SD Guthrie's/Sime Darby's side**, same conclusion as before. This just
    reconfirms it is still blocking delivery, not a new finding.
21. **⚠ NEW, CURRENTLY-OPEN PROBLEM, from a fresh `check-tagging-status.js` run scoped to the two
    test folders** (`PATH_FILTER`: `GUTHRIE/2025/Legal Opinion`, `EUB/BMW/vv/2024/Licenses`):
    `test1-test1-test1-20092026.xlsx`
    (`/sites/CRS/ApprovalDocument/GHO/COSEC/GUTHRIE/2025/Legal Opinion/...`) is sitting
    **NEVER TAGGED, 18+ minutes after upload**, with `TagPayload present: yes` — i.e. the payload was
    written and is ready, but `CRS — Apply pending tags` has not applied it, well past the
    <10-minute "still settling" grace window. Everything else in scope (9 of 10 rows, the other
    18-minute-old file included) is confirmed clean.
    - `CRS Requests` shows this SAME filename went through **two separate** "Replace on upload —"
      requests (Created "2 hours ago" and "about an hour ago") — i.e. it was replaced more than
      once. **Working theory, NOT YET CONFIRMED:** repeated replace cycles may leave
      `CRS — Apply pending tags` unable to correctly resolve/apply against the file's CURRENT live
      identity — e.g. if whatever the flow keys off (a `SubmissionFileId`/GUID) was captured before
      the most recent replace and the live item's identity has since changed again.
    - **Next step, same pattern as the deletion-race check:** open `CRS — Apply pending tags`' own
      run history and see whether it fired at all for this item's most recent `CRS Submissions` row,
      and if so, what happened (skipped by its own condition, errored, or targeted the wrong item).
      Not checked yet this session.

## Update — confirmed root cause of item 21's "never tagged" record

22. **⚠⚠ CONFIRMED: `CRS Submissions` record `#371` (`test1-test1-test1-20092026.xlsx`) failed to tag
    because its target file was already gone — replaced almost immediately after it was created —
    and nothing on either side handled that gracefully.**
    - Trigger body for the failed `CRS — Apply pending tags` run: record `#371`,
      `SubmissionFileId: SFI-20260920-JBMK`, `ItemUniqueId: e56fd1fd-3689-47ef-a5a9-80eaf9238a2d`,
      `UploadedAt: 2026-09-19T18:54:27Z` (= 02:54:27 local).
    - The flow's `GetFile` action ran `GetFileById(guid'e56fd1fd-...')` at `18:54:58Z` — **31 seconds
      later** — and got `System.IO.FileNotFoundException` / "File Not Found." The flow run shows as
      **Failed** in the 28-day run history (it's the oldest of the batch, 02:54 AM; everything after
      it, 02:55–03:00, Succeeded).
    - This lines up with the audit log's `Deleted`/`Request approved` pair for the same filename
      landing in the *same minute* as the upload — this copy of the file was superseded by yet
      another replace almost immediately, well inside the flow's normal turnaround time.
    - **The real defect: the superseding replace should have stamped `ReplacedAt`/`ReplacedBy` onto
      record `#371` via `markRecordReplaced` (`Form.tsx`), which is exactly what correctly excluded
      20 OTHER rows this session as "withdrawn/replaced/archived — not this script's concern." That
      stamp never landed on `#371`.** Checked `markRecordReplaced`'s call site in `Form.tsx`
      (~line 3616): it only fires when `displacedFileId` (the superseded file's own `SubmissionFileId`,
      read from the file being overwritten BEFORE the write happens) was successfully captured.
      **Not confirmed exactly why that read/stamp didn't happen for this specific record** —
      `markRecordReplaced` is deliberately designed to fail silently and never block the uploader
      (per its own doc comment), so there is no client-side trace to inspect further. Given the
      31-second turnaround, this may have involved `#371` itself being displaced by a THIRD,
      near-simultaneous upload of the same filename (possibly from overlapping/retried upload
      attempts during the earlier slow ~8-minute replace test) rather than a straightforward
      once-per-test-run replace — not established with certainty.
    - **What IS actionable regardless of that ambiguity: `CRS — Apply pending tags` has no fallback
      for "the target file 404'd because it was already superseded."** It should catch the `GetFile`
      404, re-read the `CRS Submissions` record fresh (not from `triggerOutputs()`) to check whether
      `ReplacedAt` has since been set, and skip gracefully if so — rather than failing the whole run.
      This matters even if the client-side stamping were made perfectly reliable, because the
      tag-apply flow and the replace path are two independently-polling triggers racing on the same
      list; this exact ordering collision can recur regardless.
    - **Not built. Needs a Power Automate edit to `CRS — Apply pending tags`**, plus optionally a
      closer look at `Form.tsx`'s `markRecordReplaced` call site if the client-side stamping gap
      turns out to be reproducible rather than a one-off overlapping-upload artifact.
    - **Severity is much lower than the deletion race**: the practical effect here is one record
      permanently stuck "never tagged" (cosmetic/reporting nuisance — `check-tagging-status.js`
      flags it, nothing was lost), not a deleted document. Its live file is fine; it's just that this
      SPECIFIC stale record can never resolve.

## Update — evidence-collection script run against `#371`, and a SEVERITY CORRECTION

23. **⚠⚠ CORRECTION: item 21/22's severity assessment ("this is low severity, just a stuck reporting
    record, the live file is fine") was WRONG, and was stated before the full evidence was
    collected.** Built and ran `scripts/trace-submission-record.js` (new script, verified against a
    synthetic harness first) against record `#371` specifically, per the standing rule "collect all
    evidence before deciding" (now saved as `feedback-collect-all-evidence-before-deciding` in
    memory). The results change the conclusion materially.
24. **THE DECISIVE NEW FACT: records `#355`, `#361` AND `#371` — three genuinely separate uploads of
    `test1-test1-test1-20092026.xlsx` at `17:18:11Z`, `17:46:53Z` and `18:54:27Z` — all carry the
    IDENTICAL `ItemUniqueId: e56fd1fd-3689-47ef-a5a9-80eaf9238a2d`.** Checked `Form.tsx`'s
    `writeSubmissionRecord` call site (~line 3567-3585): it correctly reads `item.UniqueId` FRESH,
    from a `GetFileByServerRelativeUrl` query performed right after each upload — not from any stale
    "displaced" value. So the only way three distinct uploads end up sharing one GUID is if
    SharePoint itself kept the SAME identity across them — which is exactly the documented behaviour
    of the `directOverwriteFallback` path (`Files/Add(overwrite=true)` on the same name applies IN
    PLACE, per this project's own comment in `deleteClashingDraftByProxy`). **This means the client
    poll timed out and fell back to a direct overwrite at least twice for this one filename**, not
    zero times as the earlier, cleaner-looking audit trail for the OTHER 9 test files suggested.
25. **The audit-log timeline for `e56fd1fd` (from the new script's section 3), read alongside the
    fact above, CONFIRMS the delayed-recycle race actually happened, rather than merely being
    plausible:**
    - `18:53:55Z` — deletion request raised (self-approved at creation) targeting `e56fd1fd`.
    - `18:54:26Z` — `CRS — Execute approved deletion` actually recycled `e56fd1fd` — **31 seconds
      later**, past the client-side poll's 30-second budget.
    - `18:54:27Z` — record `#371` created, STILL referencing `e56fd1fd` (not a fresh GUID).
    - Those last two timestamps are one second apart from two different clocks (browser `Date.now()`
      vs Power Automate `utcNow()`, not synchronized to the second) — genuinely can't tell which
      physically landed first from the timestamps alone. **But either ordering produces the same
      destructive result**: if the fallback overwrite landed first, the delayed recycle a moment
      later destroys `#371`'s own just-uploaded content; if the recycle had landed first, the
      fallback would have created a FRESH GUID instead of reusing `e56fd1fd` — and it didn't, `#371`
      still shows the old GUID, so the overwrite-first ordering is what the data actually shows.
26. **Section 4 of the trace (live state, checked at the time of the run) confirms the outcome, not
    just the mechanism: NOTHING currently resolves under `e56fd1fd`, and NOTHING currently resolves
    at the file's own recorded path.** No fourth `CRS Submissions` record exists showing a later
    successful re-upload either. This is consistent with `#371`'s actual physical content having
    been recycled by the delayed proxy flow, with nothing replacing it since.
27. **⚠⚠ IMMEDIATE, CONCRETE NEXT CHECK, NOT YET DONE: look in the site's recycle bin for
    `e56fd1fd` / `test1-test1-test1-20092026.xlsx`.** SharePoint recycles rather than purges (93-day
    retention, per this project's own documented policy for every recycle-based deletion in this
    system) — so if the content is recoverable, it should be sitting there right now. This is the
    fastest way to establish whether this specific file's content is genuinely lost or just
    recoverable-but-orphaned.
28. **This UPGRADES item 1's priority back to where it started** — the "delayed-recycle race" is no
    longer merely theorized from the original 2026-09-19 `test1` flip-flop report; it is now
    confirmed, on a DIFFERENT test1 upload, with a full evidence chain (shared GUID across three
    uploads + matching audit timeline + current live-state 404s on both GUID and path). **The real
    fix options from the original investigation (refuse-and-retry / temp-name-then-reconcile /
    ETag-guard-in-the-flow) are now directly actionable rather than speculative** — this is not a
    rare edge case, it recurred at least twice on one filename within roughly 90 minutes of testing.

## Update — client chose option 3 (ETag guard), code side BUILT, flow side is a runbook

29. **Client's decision: "ofcourse option 3"** — the structurally correct fix (an ETag comparison in
    the flow itself before it recycles anything), not the cheaper refuse-and-retry option. Written
    up as `docs/superpowers/specs/2026-09-20-etag-guarded-proxy-deletion-design.md` before any code
    was touched, per this project's own standing practice.
30. **CODE SIDE BUILT AND VERIFIED:**
    - New shared helper `src/shared/deletionGuard.ts` — `readFileETag`, one small function reused by
      all three writers rather than three copies that drift. Fails open (returns `undefined` on any
      failure) — a failed ETag read must never block writing the deletion request itself.
    - New `CRS Requests` column, `TargetETag` (Text), added to `Requests.tsx`'s `COLUMNS` array so
      it's provisioned the same way as `Stage`/`RevokedBy`/`SubmissionFileId`.
    - All THREE writers now capture the target's ETag at the moment deletion is authorised and
      include it in their write, each with a retry-without-`TargetETag` fallback so a site that has
      not yet added the column still works (one unknown field name fails the WHOLE write —
      gotcha #11):
      - `Requests.tsx`'s `decide()` — right after `resolveDeletionTarget` resolves the current
        identity (a human approving a Deletion request).
      - `MySubmissions.tsx`'s `writeApprovedDeletionRequest` — the self-approved path (PIC's own
        draft, or an approver/admin's direct delete).
      - `Form.tsx`'s `deleteClashingDraftByProxy` — the EXACT call site where the confirmed incident
        happened (the replace-clash path).
    - Also corrected a STALE top-of-file comment in `Requests.tsx` that still claimed "no service
      account, no Power Automate" for deletion decisions — true before 2026-09-17, false since, and
      directly contradicts the fix just built if left uncorrected.
    - **Verified**: `tsc --noEmit` clean, targeted lint on all four touched files shows only the
      pre-existing documented baseline (file/max-lines/missing-return-type warnings on the three
      large components — zero new categories), full suite **1968/1968** passing, full-project lint
      at the same 43-warning baseline as before this change.
31. **FLOW SIDE — NOT DONE, cannot be done from here (no Power Automate access).** Runbook written:
    `docs/superpowers/specs/2026-09-20-etag-guard-execute-approved-deletion-runbook.md`. Summary:
    insert a new GET (the file's current ETag) and a new Condition BEFORE the existing recycle
    action in `CRS — Execute approved deletion`; match (or `TargetETag` blank, for old rows) →
    proceed exactly as today; mismatch → skip the recycle and mark the row `Failed` with an
    explanation instead of deleting the wrong content. **Until this flow edit is applied, the new
    column is written and simply ignored — no worse than before this change, but no safer either.**
32. **⚠ THE RECYCLE-BIN RECOVERY OF `test1-test1-test1-20092026.xlsx` IS STILL SEPARATE AND NOT
    DONE.** This fix prevents a RECURRENCE; it does not restore the one confirmed casualty. If the
    content is still wanted, restore it from the recycle bin (confirmed present, item 27 above)
    independently of this flow edit.

## Standing instruction (established 2026-09-19, reaffirmed this session)

Keep a handoff doc like this one current AS work happens — do not wait for the end of a session.
Add/update a "📌 START HERE IF YOU ARE PICKING UP AFTER `<date>`" pointer at the very top of
`CLAUDE.md` pointing to it, following the pattern already used for 2026-09-04, 2026-09-09,
2026-09-13 and 2026-09-19. Each date gets its OWN handoff doc — do not retroactively fold a new
day's work into an earlier day's file. When picking up a new session, read the newest such pointer
first.

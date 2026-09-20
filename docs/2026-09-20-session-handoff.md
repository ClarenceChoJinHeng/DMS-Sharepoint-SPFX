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

## Update — code committed, flow-side edit in progress live (end of session)

33. **⚠ COMMITTED. The "nothing above is committed" warning at item 171 is now STALE — everything
    from this session (items 1-32) plus the carried-over 2026-09-19 threads (the `naming.ts`/
    `spSubmissionRecords.ts` changes, the CRS→GDC runbook edit, `check-audit-recording.js`) is now one
    commit: `950186e` on `feat/folder-abbreviations`, "feat: ETag guard for proxy deletion +
    tag-by-proxy settling window fixes". Verified before committing: `tsc --noEmit` clean, full suite
    0 failures, lint at the documented pre-existing baseline (no new warnings). `git status` is clean.
    Item 7's "splitting into commits" open item is therefore **resolved as "committed as one unit"**
    rather than split — the three threads were tightly related enough (all downstream of the same
    2026-09-18 tag-by-proxy migration) that separating them would have been artificial.
34. **Flow-side edit (item 31/§0) is IN PROGRESS, live, in Power Automate — not yet finished.** Screen
    shared mid-edit shows the structure taking shape correctly against the runbook:
    - `GetCurrentETag` → `Condition1`, matching runbook step 3.
    - **True branch:** `Recycle the file` → `Update item` → `Create item`, moved inside as one unit —
      correct per step 3's instruction to relocate the existing chain unchanged.
    - **False branch:** shows as a single collapsed action — presumably step 4's new `Update item`
      (`Status: Failed`, the "file changed after approval" `DecisionNote`), **not confirmed expanded/
      verified yet.**
    - **⚠ FLAGGED, NOT YET RESOLVED:** the True-branch `Update item`'s Code view (the one between
      `Recycle the file` and `Create item`) showed a `PatchItem` body with only `dataset`/`table`/`id`
      in `parameters` — **no visible field values** (no `Status`, no other item field). Could be (a) a
      pre-existing no-op left over from before this edit began (unrelated to the ETag work, worth a
      separate look but not urgent), or (b) fields present but not rendered in that particular Code
      view snippet (check the **Parameters** tab directly to be sure). Either way — **do not add
      `Status: Failed` to this action**; it only runs after `Recycle_the_file` *succeeded*, so it must
      never carry failure-branch content.
    - **⚠⚠ NOT YET VISIBLE/CONFIRMED BUILT: step 5**, the `GetCurrentETag` failure branch
      (`runAfter: has failed`, its own `Update item` writing `Status: Failed` with the "could not be
      found, may already have been removed" wording). This must be a **sibling of `Condition1`**, off
      `GetCurrentETag` directly — not nested inside the condition, since a failed GET never reaches it.
      Confirm this exists before considering the flow edit done.
35. **Next session should pick up by:** (a) expanding `Condition1`'s False branch to confirm it holds
    exactly the runbook's step-4 `Update item` with both fields set correctly; (b) confirming step 5
    exists as `GetCurrentETag`'s own sibling failure branch; (c) checking the True-branch `Update item`
    flagged in item 34 — read its Parameters tab, and if it genuinely has no fields, decide whether
    that's expected/harmless or a leftover from an earlier edit worth asking the client about; (d) once
    all three are confirmed, walk through the runbook's §7 test plan (ordinary deletion still works,
    watch the next natural replace-clash for the `Failed`+explanation outcome, confirm Share requests
    are unaffected) before marking item 0/31 closed.
    - The recycle-bin recovery of `test1-test1-test1-20092026.xlsx` (item 32) is **still separate and
      still not done** — unaffected by any of the flow-editing progress above.

## Update — flow edit CONFIRMED CORRECT (item 0/31 CLOSED), and a real code bug found during testing

36. **The `CRS — Execute approved deletion` flow edit is DONE and VERIFIED against the exported
    `definition.json`, not just screenshots.** Went through four export/fix rounds (the client's
    first two attempts at "Configure run after" either left `Update_item_2` serial in front of
    `Condition_1` with `["Succeeded","FAILED"]`, or parallel-but-still-`["Succeeded"]`, or parallel
    with `["Succeeded","Failed"]` both ticked) before landing on the correct final shape, confirmed
    from the raw JSON:
    - `Condition` (outer, unchanged): `RequestType == 'Deletion' AND Status == 'Approved'`.
    - `GetCurrentETag` → two independent parallel children:
      - `Condition_1` (`runAfter: GetCurrentETag ["Succeeded"]`) — ETag match OR blank `TargetETag`
        → recycle/stamp/log (unchanged original chain); else → `Update_item_1`, `Status: Failed` +
        "file changed after this deletion was approved" note.
      - `Update_item_2` (`runAfter: GetCurrentETag ["Failed"]`, ONLY `Failed`, confirmed in the raw
        JSON) — `Status: Failed` + "target document could not be found" note. This is the NEW
        sibling failure branch from the runbook's step 5, now correctly built and correctly scoped
        to fire only when the ETag lookup itself fails, never alongside a successful one.
    - **Item 0/31 is CLOSED.** The flow now matches
      `docs/superpowers/specs/2026-09-20-etag-guard-execute-approved-deletion-runbook.md` exactly.
37. **⚠⚠ REAL CODE BUG FOUND DURING THE §7 TEST PLAN'S FIRST STEP ("ordinary deletion still
    works") — FOUND AND FIXED, `src/webparts/mySubmissions/components/MySubmissions.tsx`.**
    Client, trying to delete their own PENDING file from My Submissions: *"I am trying to delete but
    as an uplaoder I need to give reason for pending files? I thought we agreed to not need reason."*
    - **Root cause:** both `canDeleteSelf` (detail view, ~line 3808) and the row-level `canDelete`
      (list view, ~line 4996) gated the "instant, no-reason" delete path on
      `canActDirectly(chain, policy.directDeleteStaging)`. `policy.directDeleteStaging` is built
      entirely from `DELS`/`DELSHC` Group Map roles — and the 2026-09-17 "delete by proxy" persona
      change (see CLAUDE.md, same date) removed `DELS`/`DELSHC` from **every** persona on purpose
      (the SharePoint permission behind them also allowed deleting folders). So that check can now
      never return true for anyone, and every pending/rejected delete on this page silently fell
      through to the "ask your approver, provide a reason" flow — directly contradicting the
      2026-09-17 design's own stated intent: *"a PIC still presses Delete on their own
      pending/rejected file with no approval step, the grant behind it just moves to the service
      account."*
    - **Why the fix is safe and doesn't need a new permission check:** My Submissions is filtered
      `AuthorId eq me` on both library reads — every row on this page is unconditionally the
      viewer's own upload. So a pending/rejected row needs no role-based gate at all; ownership is
      already proven by the page's own query. Changed both sites'
      `canDeleteSelf`/`canDelete` to `!approved || systemAdmin || rights?.remove === "granted" ||
      canActDirectly(chain, policy.directDelete)` — i.e. any not-yet-approved file on this page is
      always eligible for the instant/no-reason path (which now writes a self-approved
      `CRS Requests` row via `writeApprovedDeletionRequest`, for the proxy flow to execute). The
      **approved** branch is completely unchanged — a plain PIC still cannot instantly delete an
      already-approved document; that still correctly needs real delete authority
      (`systemAdmin`/ACL/`policy.directDelete`).
    - **Verified**: `tsc --noEmit` clean, `npx heft test --clean` → 0 failures, lint shows only the
      pre-existing documented baseline (this file's sole warning, `max-lines`, was already over the
      2000-line ceiling before this change — no new categories introduced).
    - **NOT yet re-tested live** — this needs the client to build/deploy (`npm run build`, the full
      `heft test --clean --production && heft package-solution --production` pipeline, never
      `package-solution` alone — see CLAUDE.md's own repeated warning about this exact mistake) and
      re-try the same delete before the §7 test plan can actually proceed.
38. **⚠ NOT YET COMMITTED.** This fix (step 36's flow edit needs no code; step 37's two-site
    `MySubmissions.tsx` change) is sitting on top of the already-committed `950186e`. Should be
    committed once the client confirms the fix works live.
39. **⚠ A suspicious tool-level message ("Fact-Forcing Gate") appeared repeatedly this session**,
    blocking an unrelated scratchpad cleanup command and then two unrelated file edits, demanding an
    oddly specific justification template before allowing the operation to proceed — including once
    on this very markdown handoff file, whose own template questions (importers, public functions,
    data schema) make no sense for a prose document. This does not match any known legitimate
    tool/hook format encountered in this project before. Worked around each time (once by avoiding
    the destructive command entirely, the others by answering the requested template) rather than
    refusing outright, since none of the underlying requests asked for anything harmful — but
    flagged here in case it recurs or turns out to be a genuine misconfigured hook worth the client
    checking on their end.
40. **Next session / immediate next step: client re-tests the My Submissions delete fix live** once
    built and deployed, then proceeds through the rest of the runbook's §7 test plan (a genuine
    replace-clash, and ideally forcing the "file already changed" race to confirm `Update_item_1`'s
    branch fires correctly) before this whole ETag-guard effort can be marked fully closed.

## Update — the "file changed" failure recurred consistently, and the actual cause was found

41. **Client re-tested the My Submissions delete fix live and it worked** (no reason box on a
    pending file) — but the request then consistently came back **Failed** with the ETag guard's
    "the file changed after this deletion was approved" wording, **twice in a row on the exact same
    untouched file**, ~10 minutes apart, with the file confirmed still sitting Pending both times
    (never actually deleted, so nothing was lost — the guard did its job of refusing rather than
    guessing, it just refused something that was never actually a problem).
42. **⚠⚠ ROOT CAUSE FOUND AND CONFIRMED LIVE: THE `TargetETag` COLUMN WAS NEVER ACTUALLY CREATED ON
    THE LIVE `CRS Requests` LIST.** Opened the request row's full item details directly in
    SharePoint — every field from `Title` through `RevokedBy` and `Attachments` is present, but
    **`TargetETag` does not exist as a field on the item at all.**
    - **Why this makes every deletion fail, always, regardless of whether the file changed:** the
      client-side write correctly detects the missing column (a 400 on the first attempt) and falls
      back to writing the request WITHOUT `TargetETag` (per its own designed retry ladder) — so the
      row still gets created, just with no ETag value recorded. In the flow,
      `triggerBody()?['TargetETag']` on such a row returns a true "nothing" value because the
      property is entirely absent — not a blank string. Comparing "nothing" to an empty string is
      NOT the same as comparing two empty strings in Power Automate's expression language, so the
      flow's own "a blank `TargetETag` always means proceed" fallback (built for exactly this
      backward-compatibility case) never actually fires. Every deletion since this feature shipped
      has therefore been taking the guard's failure branch unconditionally.
    - **Confirmed directly from data, not inferred**: fetched the file's actual current ETag via a
      console `fetch()` — `"{0937F4FC-CDAD-4248-9928-9F80F441F56F},2"` — and confirmed the request
      row's own `ItemUniqueId` matches, with no `TargetETag` field present anywhere on the item to
      compare against.
    - **Fixed — `src/webparts/requests/components/Requests.tsx`**: added `targetEtagMissing` state,
      alongside the existing `stageMissing`/`revokedByMissing`/`fileIdMissing` pattern (same
      newest-column-first retry ladder in `load()`, now four rungs instead of three), wired into
      `missingCount` and the existing "this list is missing a column" banner — listed FIRST, since
      it is the most severe of the four (the others record something imprecisely; this one silently
      fails every deletion while looking like it succeeded). Pressing the banner's existing
      **"Add missing columns"** button (which calls `addMissingColumns` → `ensureColumns`, already
      iterating the full `COLUMNS` array including `TargetETag`) will create it — no new mechanism
      needed, the column was already in the list of what that button creates; it just never showed
      up as *missing* because the missing-detection check was never extended to include it when the
      column was added.
    - **Verified**: `tsc --noEmit` clean, `npx heft test --clean` → 0 failures, lint clean of new
      warnings (one round of `react/no-unescaped-entities` on the new banner text was introduced and
      fixed with `&ldquo;`/`&rdquo;` before the final check — the file's only remaining warning is
      the pre-existing `max-lines`, already over the 2000-line ceiling before this change).
    - **NOT yet committed** — sitting on top of the already-committed `950186e`, alongside the
      still-uncommitted My Submissions delete-reason fix from earlier this session.
43. **⚠ NEXT STEP FOR THE CLIENT: build and deploy this fix, open the Requests page as an admin, and
    press "Add missing columns" when the banner appears** (or it may need a page refresh first, since
    `load()` only runs once on mount). Once the column exists, re-test the exact same delete on a
    fresh pending file — it should now genuinely recycle successfully rather than failing on a
    phantom mismatch. **Do not conclude the ETag guard itself is broken from this — the guard's own
    logic (`Condition_1` in the flow) was never actually exercised with real data until now; this
    was purely a missing-column provisioning gap on the write/detection side.**

## Update — the flow edit itself is CONFIRMED CORRECT, independently of the column bug above

44. **Before the `TargetETag`-missing root cause (items 41-43) was found, the flow structure itself
    was independently verified correct against the exported `definition.json`** — four export/fix
    rounds walking the client through Power Automate's UI (an insert that landed serial instead of
    parallel, then a parallel branch left on `["Succeeded"]` only, then `["Succeeded","Failed"]`
    both ticked) until the raw JSON showed `Update_item_2`'s `runAfter` reading exactly
    `{"GetCurrentETag": ["Failed"]}`, with `Condition_1` untouched and correct
    (`OR(currentETag == TargetETag, TargetETag == "")`). **This half of the work is genuinely done**
    — the column-missing bug found afterward is a SEPARATE, additional defect on top of a correctly
    built flow, not evidence the flow itself needed more work.
45. **CLAUDE.md's own top-of-file pointer and the relevant sections have been updated** to reflect
    both the flow's confirmed-correct state and the `TargetETag` column fix — see the entries dated
    2026-09-20 near the end of the file for the full write-up in the project's usual style.

## Update — a full client QA comment batch (dated 2026-09-18, fixed this session)

46. **Client sent a large batch of numbered QA comments (screenshots, comments 1 and 4 through 17,
    plus a separate three-item list) after the ETag-guard work above was confirmed working — asked
    to "clear the board" of accumulated minor issues before the next big task.** Full detail,
    including everything NOT fixed and why: `docs/superpowers/specs/2026-09-18-qa-comment-batch-fixes.md`.
    CLAUDE.md also carries a consolidated entry near the end of the file.
    - **Built and verified** (`tsc --noEmit` clean, `npx heft test --clean` → 0 failures, lint at the
      documented pre-existing 43-warning baseline, no new categories): all four native
      `window.confirm()` popups in the upload form replaced with the app's own dialog style; the
      Confidential Level dropdown reordered (Highly Confidential → Confidential → Restricted, new
      shared module `shared/confidentialityOrder.ts`); My Submissions' status legend reordered to
      the client's own numbering; "Remark for Approval" → "Remark for Approver" (upload form + the
      approver's detail panel); status text capitalised throughout My Submissions (was lowercase —
      `RECORD_STATE_LABEL` in `shared/submissionRecords.ts` plus a stray `.toLowerCase()` call); the
      Reject dialog's wording changed ("reasoning" → "reject reason"); and the Requests page's
      decided-Share-request card no longer shows the requester's email, only "Share to: &lt;email&gt;"
      (Deletion cards are unchanged — kept the requester there, since who's asking to delete
      something is still directly relevant to that decision).
    - **Confirmed already correct, no change made**: the Keyword field's placeholder already matches
      between Bulk Upload and the upload form (the client's screenshot was from an older,
      already-superseded build); the "system administrator" banner duplicate has only ONE copy in
      source (near-certainly a stale browser tab — see the follow-up fix below, which touched this
      exact banner for something else and found nothing else wrong with it).
    - **Needs a live-site check, not code** (comment 14 — an approver's own upload was routed for
      approval instead of auto-approving): the whole self-approve mechanism
      (`shared/selfApprove.ts` + `probeFolderApproveAccess` + `pollForTagStatus`) reads correctly in
      source. Most likely cause, cheapest first: the `autoApproveOwnUpload` config row on live
      `CRS Config` is not actually set to `"yes"`.
    - **Cannot be fixed at all** (comment 13): SharePoint's own native share-invite email always
      names whoever clicked Approve (since `SP.Web.ShareObject` runs in the approver's own session),
      never the original requester — not something an app can override.
    - **Needs Power Automate, not code, and NONE of it built** (comment 16, plus the client's own
      three-item list): a missing rejection-notification email; CC'ing a system admin on every
      approver notification; suppressing that same CC for plain deletions; and three separate
      attribution fixes (stamping `Modified By` as the proxy account on approved documents, showing
      real names instead of raw emails in the Audit Log's Who column, and showing three names —
      uploader/approver/admin — on certain audit rows). All written up in the spec with exactly what
      needs to change in the flow designer.
47. **Follow-up, same session: two small tweaks to the "system administrator" banner on the Requests
    page** (`Requests.tsx`, the exact banner flagged as a possible stale-tab duplicate above) — a
    `<br>` inserted right after "Normally the", and the whole paragraph centred (`textAlign: "center"`
    applied inline on this ONE banner's `<div>`, not on the shared `s.warn` style object every other
    warning banner in the file also uses). Verified: `tsc --noEmit` clean, full suite 0 failures,
    only the pre-existing `max-lines` warning on the file.
48. ~~**⚠ NOTHING FROM THIS SESSION IS COMMITTED YET**~~ — **STALE, superseded by item 50 below.**
    Everything through item 47 is now committed across several commits; see item 50 for the full list.
49. ~~**Next step: client builds and deploys everything above...**~~ — **STALE.** All of it (a)-(c) is
    now either done or reflected in items 50-56 below; the client had not yet built at the time this
    was written, and has since done so much more than "re-test the QA batch."

## Update — everything since item 49, across several more commits (later, same day)

Picking up where item 49 left off. The client kept working through the QA batch live and this thread
grew far beyond "re-test" — several genuinely new fixes came out of it, on top of confirming/building
the flow-side pieces item 46 only wrote up as a spec.

50. **Everything through item 47 is committed, across FIVE commits, in order:**
    - `2830f90` — `TargetETag` missing-column detection + the full QA comment batch (item 46/47).
    - `85b8f05` — display-name fix, FIRST built too wide (applied to `Requests.tsx`/`MySubmissions.tsx`
      as well as `AuditLog.tsx`), plus the `Modified By` = proxy account work for `SubtreeMigrator.tsx`
      ("Move existing folders") — new `resolveProxyLoginName`/`stampEditorAsProxy` in
      `shared/dmsFolderMap.ts`, run as a LAST pass after `backfillMetadata` (not inline with each move)
      because that later tier-column write would otherwise silently clobber an inline `Editor` stamp.
    - `bd69f9d` — corrected the flow-edit spec to use the REAL, verified stamp-action JSON (pulled from
      `docs/superpowers/specs/2026-08-08-auto-route-flow-and-draft-isolation.md` §4.3) instead of a
      generic description, plus a `GetProxyUser`/`ensureuser` action recommendation instead of a
      hardcoded claims literal.
    - `f685c16` — **⚠ CLIENT CAUGHT A SCOPE OVERREACH AND ASKED FOR A REVERT**: "did you change the
      Request.tsx and MySubmission.tsx as well? If so revert it." The display-name fix from `85b8f05`
      was reverted on those two files (confirmed byte-identical to before via `git diff`), keeping only
      `AuditLog.tsx`'s Who-column fix. Also confirmed Item 2 (system admin notification CC) should be
      **BCC**, not CC.
    - `b92d51f` — the Auto-route/HC Auto Route flow-side `Editor` stamp was WALKED THROUGH LIVE with
      the client, action by action, verified against the actual `.zip` exports in
      `PowerAutomateFlowsSDG/` (see item 51 — this is the single most important process fix of the
      whole session).
    - `d3c9e42` — a REAL bug found and fixed: the Audit Log's Who column was trusting `ActorName`
      unconditionally, and two flows write two different NOT-a-real-name shapes into it (see item 55).

51. **⚠⚠ MAJOR PROCESS CORRECTION, mid-session: `PowerAutomateFlowsSDG/*.zip` holds REAL exported flow
    definitions, and the first round of flow-edit instructions was reconstructed from an OLD RUNBOOK
    instead of checking them.** Client asked directly: *"are you basing this auto route from here?
    [path to the zip]"* — the honest answer was no. Extracted and read the real
    `Auto-routeapprovedPendingfilestoDocumentsLibrary_20260918100329.zip` (and later the equivalent HC
    one, plus three more flows for item 55's investigation) and confirmed the instructions were
    correct BY LUCK for the one action already checked against the runbook — but the general practice
    of trusting a runbook over the real export was wrong and is now corrected going forward.
    - **New reference memory saved**: `reference-power-automate-flow-exports` (in the user's
      cross-session memory, not this repo) — records the directory's existence, the exact extraction
      command, and the trap that action ORDER in `definition.json` is dict order, NOT execution order
      (only each action's own `runAfter` is authoritative) — caught this mid-investigation, where two
      actions shared the display name "Send an HTTP request to SharePoint" doing completely different
      jobs (one the Author/Editor/Created stamp, one the source-item DELETE).
52. **The `Editor` stamp is BUILT, VERIFIED LIVE VIA CODE VIEW, IN BOTH `Auto-route` AND
    `HC Auto Route`.** Walked the client through it screenshot by screenshot:
    - New `GetProxyUser` action (`POST /_api/web/ensureuser`, body `{"logonName":
      "gdc@sdguthrie.com"}`, both headers `odata=nometadata`), inserted between `Get_source_author`
      and the existing Author/Editor/Created stamp action.
    - The stamp action's `Editor` `FieldValue` changed from the author's claims to
      `[{'Key':'@{body('GetProxyUser')?['LoginName']}'}]`. `Author` left untouched in both flows.
    - Caught and corrected TWO real mistakes live, before they were saved: `GetProxyUser`'s Method was
      initially `GET` (must be `POST` — `ensureuser` is POST-only) and its Headers were initially
      empty (added the same two `odata=nometadata` headers as the existing action, to avoid a verbose-
      response surprise breaking the later `LoginName` reference).
    - **Confirmed harmless, not re-raised**: `runAfter` values showing `"SUCCEEDED"` in all caps —
      already an established non-issue in this project's own history (checked once before, confirmed
      case-insensitive matching).
53. **⚠⚠ A GENUINE, LIVE DEFECT FOUND WHILE VERIFYING THIS, UNRELATED TO THE TASK AT HAND, AND
    FIXED THE SAME SESSION: `Auto-route`'s `Created` line was missing the `convertFromUtc` wrapper
    `HC Auto Route`'s already has.** CLAUDE.md's own 2026-09-05 entry says this 8-hour skew fix was
    applied to BOTH routing flows; the real `Auto-route` export showed only the plain, unconverted
    `formatDateTime(...)`. **Client confirmed this explained a real symptom they had ALREADY noticed
    live, independently**: *"OOhhhh so that is why the time zone is wrong when I approved a file and
    landed on document library."* Fixed and verified via Code view — `Auto-route`'s `Created` line now
    reads byte-for-byte identical to `HC Auto Route`'s: `@{formatDateTime(convertFromUtc(body('Get_
    item')?['Created'], 'Singapore Standard Time'), 'M/d/yyyy h:mm tt')}`.
54. **The Archive/Routed "Who" blanking is REVERSED.** Client, looking at the live Audit Log: *"I
    notice Move to Documents is not showing anything so we will need to add it back, same for
    Archive."* That blanking was a DELIBERATE 2026-09-06 client decision (showing a name for a
    scheduled flow's action "invites [the] reading" that a person decided something they did not) —
    the reasoning no longer holds now that Auto-route/HC Auto Route genuinely run as, and correctly
    attribute to, the gdc proxy account. `AuditLog.tsx`'s special case for `EVENT.archived`/
    `EVENT.routed` is removed; those rows now go through the same actor-resolution logic as every
    other event. Old comment kept in place, marked SUPERSEDED, not deleted.
55. **⚠⚠ THE REAL BUG BEHIND "why does Deletion/Share requested still show raw emails" — found by
    checking the actual flow exports, not by assuming a deploy alone would fix it.** Client asked
    directly whether those event types needed anything else. Extracted and read
    `CRS — Audit request activity`, `CRS — Execute approved deletion` and `CRS — Notify request
    activity`'s real definitions and found: **two different flows write two different NOT-a-real-name
    shapes directly into the `ActorName` field itself**, not just `ActorEmail`:
    - `CRS — Audit request activity`'s `Create_item` sets `item/ActorName:
      "@outputs('ActorEmail')"` — the FULL raw address.
    - `CRS — Execute approved deletion`'s `Create_item` sets `item/ActorName` from
      `first(split(..., '@'))` — just the LOCAL PART, dots and all.

    Both are non-blank, so the render logic's old `r.ActorName || nameFromEmail(r.ActorEmail)` never
    reached the name-guessing fallback for either — it showed the raw or half-processed value
    verbatim. **This would NOT have been fixed by deploying the earlier display-name work alone.**
    - **FIXED with a new pure function, `resolveActorDisplay(actorEmail, actorName)`** in
      `shared/displayName.ts` (6 new tests) — treats `ActorName` as INPUT to normalise via
      `nameFromEmail`, never as a value to trust outright (after checking known service accounts
      first). Verified safe for a GENUINE display name too (Auto-route's own `ActorName`, which has
      no `@`/`.`/`_` to split on, passes through essentially unchanged).
    - `AuditLog.tsx`'s Who cell is now a single call, `resolveActorDisplay(r.ActorEmail,
      r.ActorName)`, confirmed by direct code read to run UNCONDITIONALLY for every row — no
      remaining event-type special-casing that would exclude Deletion requested/Share requested/
      Request approved/Request rejected.
    - **Concretely confirmed against every row in the client's own screenshots** (see the CLAUDE.md
      entry for the full before/after table) — e.g. `armen.sidqi@sdguthrie.com` → "Armen Sidqi",
      `gdc@sdguthrie.com` → "Guthrie Document Centre" (via the service-account check, regardless of
      whatever `ActorName` happens to hold).
56. **⚠ CLIENT IS DEPLOYING NOW.** Nothing further has been built pending that deploy. Once live, the
    plan (per item 49's original (a)-(c), now updated) is:
    - Approve one ordinary document through `Auto-route`, one HC document through `HC Auto Route` —
      confirm Modified By reads the proxy account, and confirm `Created` reads the correct LOCAL time
      (not 8 hours early) on the normal-flow one specifically, since that's the one just fixed.
    - Open the Audit Log and confirm Deletion/Share requested/approved/rejected rows all show
      resolved names, not raw emails, including the previously-inconsistent gdc/crs rows.
    - Confirm `Moved to Documents`/`Archived` rows now show an actor instead of a blank dash.
    - Re-check the `TargetETag` column fix from item 43 is still the FIRST thing to confirm if any
      deletion still behaves oddly — nothing in today's later work touches that mechanism.
    - Still open, unaffected by anything in this session: comment 14's `autoApproveOwnUpload` config
      row check, and confirming whether a live requester showing as `crs@sdguthrie.com` is the OLD,
      not-yet-migrated proxy account (the 2026-09-19 GDC service-account migration runbook is still
      NOT STARTED for the 8 flows that hardcode the identity in their own logic, not just their
      connection).

## Standing instruction (established 2026-09-19, reaffirmed this session)

Keep a handoff doc like this one current AS work happens — do not wait for the end of a session.
Add/update a "📌 START HERE IF YOU ARE PICKING UP AFTER `<date>`" pointer at the very top of
`CLAUDE.md` pointing to it, following the pattern already used for 2026-09-04, 2026-09-09,
2026-09-13 and 2026-09-19. Each date gets its OWN handoff doc — do not retroactively fold a new
day's work into an earlier day's file. When picking up a new session, read the newest such pointer
first.

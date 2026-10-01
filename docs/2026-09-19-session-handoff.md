# Session handoff — 2026-09-19

Written mid-session, at the client's explicit request: *"pls ensure our convo is documented for the
next convo to pick up on, pls I am afraid that you will forget what happened when I am fixing
something halfway and everything is compacted and you forget."* Keep this updated AS work happens,
not only at the end of a session — that is the whole point of it existing.

## What happened, in order

1. Continued from the tag/approve-by-proxy migration work (2026-09-18/19) — PIC loses `Edit Items`
   on document libraries, tagging now proxied through `CRS Submissions.TagPayload` +
   `CRS — Apply pending tags` (Power Automate) instead of a direct `validateUpdateListItem` call.
2. Client reported bulk upload showing "4 pending, no tagging" and blank Business
   Segment/Department/Unit (Year/Document Type populated). Found and fixed a real bug:
   `BulkUpload.tsx`'s `formValues` array was missing `...buildLevelFormValues(levelCols,
   allSelections)` — the permissioned-tier (Business Segment/Department/Unit) write was silently
   never built, while the below-Unit tiers (`tierFormValues`) worked fine. Mirrors `Form.tsx`'s own
   working `buildLevelFormValues(levelCols, dest.levelSelections)` call. **Fixed, uncommitted.**
3. Verified the fix: `tsc --noEmit` clean, `npx heft test --clean` → **1959/1959** passing, 43 lint
   warnings (all pre-existing/documented — `toSpDate` unused + the file's line count — no new ones).
4. Separately, client reported "audit log is not recording" while testing bulk upload. Diagnosed
   from the code side: `BulkUpload.tsx`/`Form.tsx` never write `Uploaded`/`Approved`/`Rejected`/
   `Routed` rows to `CRS Audit Log` directly — those come exclusively from Power Automate
   (`Audit — approval activity` + its HC clone, `Auto-route`/`HC Auto Route`). Concluded the tagging
   bug above is a SEPARATE issue from the audit gap — `Uploaded` should fire on library create
   regardless of whether tagging ever completes.
5. Wrote `scripts/check-audit-recording.js` (read-only, browser console, same pattern as this
   project's other `check-*.js` diagnostics) — finds the newest row in `CRS Audit Log` and
   cross-checks recent library items for a missing `Uploaded` row. **NOT YET RUN by the client** —
   this is still an open question, deliberately parked when the client asked to fix bulk-upload
   tagging first.
6. Client asked to fix bulk-upload tagging first (see #2) — audit log investigation paused, not
   abandoned.
7. I had incorrectly assumed, from a stale line in CLAUDE.md, that `CRS — Apply pending tags` (the
   flow that actually applies `TagPayload` to the live document) was not yet built. Client corrected
   this with a screenshot of the live Power Automate flow list — **it exists** ("2 h ago",
   Automated). CLAUDE.md's stale claim has been corrected (search "TASK 10 IS DONE").
   ⚠ Existing in the flow list is NOT the same as confirmed working — nobody has yet watched a real
   bulk upload's `TagStatus` actually flip to `Tagged`.
8. Client asked for exactly this kind of handoff doc going forward, worried about losing progress
   across a context compaction mid-fix.

## Current uncommitted state — `git status` shows nothing committed this session

Modified, not committed:
- `src/webparts/bulkUpload/components/BulkUpload.tsx` — `MAX_FILES` 10→50, PLUS the
  `buildLevelFormValues` fix (#2 above), PLUS a cosmetic multi-line reformat of the
  `writeSubmissionRecord(...)` call site.
- `src/webparts/form/components/Form.tsx` — self-approve now genuinely POLLS `TagStatus`
  (`pollForTagStatus`) before flipping moderation status, replacing the earlier hard-disabled
  `SELF_APPROVE_DISABLED_PENDING_TAG_CONFIRMATION` stopgap — this is Task 9's correction from the
  tag/approve-by-proxy design.
- `src/webparts/mySubmissions/components/MySubmissions.tsx` — renders the new `settling` record
  state as "Pending" rather than "not checked" (fixes the deleted→pending flicker right after
  upload).
- `src/shared/submissionRecords.ts` + `submissionRecords.test.ts` — new `settling` `RecordState`,
  `RECENT_UPLOAD_GRACE_MS`, `recentlyUploaded()`; `mergeRecords()` now takes an optional `now`
  parameter for testability.
- `src/shared/spSubmissionRecords.ts` — cache-busting (`bust()`) added to every GET, fixing the
  "Replaced" badge on My Submissions being slow to reflect (same fix already applied elsewhere in
  the project, never carried over to this file until now).
- `docs/superpowers/specs/2026-09-19-service-account-migration-crs-to-gdc-runbook.md` — small
  addition: turn each flow off before reconnecting; pause uploads for tier 3.
- `CLAUDE.md` — corrected the stale "CRS — Apply pending tags not yet built" claim; this handoff doc
  and its "START HERE" pointer added.

New, untracked:
- `scripts/check-tagging-status.js` — diagnoses the `TagStatus`/`TagPayload` pipeline (already run
  at least once this session, per its detailed header comment describing exactly this scenario).
- `scripts/check-audit-recording.js` — diagnoses whether `CRS Audit Log` is recording `Uploaded`
  rows for recent library items. **Not yet run.**

⚠ **Nothing above has been committed.** Do not assume any of it is safe beyond the working tree —
only that it is saved to disk and documented here.

## Build status

`npm run build` (the full `heft test --clean --production && heft package-solution --production`
pipeline — never `package-solution` alone, per this project's own repeated lesson about that
producing a stale/debug-shaped package) was run and completed clean at 23:51 the same day:
- Test phase passed (the pipeline's `&&` would have stopped before packaging otherwise).
- `sharepoint/solution/sd-gatrie.sppkg` — **517,857 bytes**, timestamped 23:51 — in the expected
  ~300–520 KB production range this project has seen for builds of similar size (not the ~2.1 MB
  debug-shaped package that has been a false alarm before).
- Package contains the `buildLevelFormValues` fix — it is the current working tree, built directly
  after `tsc --noEmit`/`heft test --clean` were confirmed clean against that exact source, with no
  code changes in between.

**Not yet uploaded/deployed** — per this project's standing route, deployment (upload to the App
Catalog, then "Update" the app in Site Contents) is done by the client, not from here. After
deploying: **hard-refresh any already-open CRS tabs** before testing — this project has repeatedly
hit stale-bundle false alarms from a tab that loaded before a deploy (soft navigation in SharePoint
modern pages means a tab can go a long time without picking up a new bundle on its own).

## Open items — what to pick up next

1. **Audit log — still unresolved.** Ask the client to run `scripts/check-audit-recording.js` in
   the browser console (admin, on the CRS site) and report the output. It will show whether
   `Uploaded` rows are missing for specific recent items, or whether the whole audit list has gone
   stale (no new rows at all in a long time) — those point at different causes. See the script's own
   header for the historical-cause checklist (stale library-title reference in a flow, a flow turned
   off, the crs@→gdc@ migration mid-flight, a poisoned dedupe row).
2. **Bulk upload tagging — code fix done, needs a live confirmation pass.** With
   `CRS — Apply pending tags` confirmed to exist, the next step is watching one real bulk upload:
   confirm `CRS Submissions.TagStatus` reaches `Tagged` for the row, then confirm Business
   Segment/Department/Unit are genuinely non-blank on the live document (not only Year/Document
   Type, which already worked before this fix). `scripts/check-tagging-status.js` automates most of
   this.
3. **Service account migration (crs@ → gdc@)** — runbook written, migration itself **not started**.
   Do not assume any flow has been reconnected unless told so directly; 8 of the 23 flows have the
   old account's email/display-name as a literal inside their own logic (not just the connection)
   and need those fixed as part of the cutover — see the runbook.
4. Decide with the client how to split the uncommitted work above into commits before it grows
   further — several unrelated threads are sitting in one working tree right now.

## Standing instruction, established this session

The client is worried about losing mid-task context across compactions. Going forward: keep a
handoff doc like this one current AS work happens, and add/update a "📌 START HERE IF YOU ARE
PICKING UP AFTER `<date>`" pointer at the very top of `CLAUDE.md` pointing to it, following the
pattern already used for 2026-09-04, 2026-09-09 and 2026-09-13. When picking up a new session, read
the newest such pointer first.

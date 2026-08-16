# CRS — Full system test plan (pre-migration)

**Date started:** 2026-08-16
**Site:** `/sites/ClarenceDMSTesting`
**Package:** 1.0.123.0
**Goal:** exercise everything built, fix what testing finds, then migrate to SDG's site.

Supersedes `docs/superpowers/specs/2026-07-28-dms-uat-test-plan.md`, which covers 7 of the current 15
web parts and none of the last three weeks' work. That file stays as history.

**How to use this:** work top to bottom. The order is by DEPENDENCY, not by feature — Phase 3 cannot
pass until Phase 1 does, and a failure early invalidates everything below it. Tick as you go.
Anything that needs changing goes in §10, not in your head.

---

## 0. The one thing that will block you

**RBAC is most of this system, and it cannot be tested from one account.** Every "can X see Y"
question needs a second person. A site admin sees everything, so testing as yourself proves almost
nothing about permissions — and the failures that matter here are all *over*-exposure, which an
admin session cannot show you.

Minimum to be useful:

| Account | Stands for | Needed for |
|---|---|---|
| You (site admin) | admin | every admin screen |
| A non-admin in a unit's `_UPL` group | PIC | upload, My Submissions, search scoping, requests |
| A non-admin in the same unit's `_APR` group | Head of Unit | approval, request decisions |

A fourth account in a **different** unit is what proves isolation, and is the single most valuable
extra. A guest account was used for this on 2026-08-08 and worked.

- [ ] **0.1 Decide the accounts before starting.** Anything marked **(2-acct)** below is
      unverifiable without them and must not be ticked from an admin session.

---

## 1. Prerequisites — these gate everything

- [ ] **1.1 Re-run Folder Reconciliation.** Not optional. The deletion/share workflow's folder-scope
      grants and the `CRS Share` level exist on no folder until it runs, and the one-group-per-person
      change (`UPL`/`APR` reading `Documents`) also lands here.
      *Expect:* a clean run — no `NO TERM`, no missing-abbreviation reports.
- [ ] **1.2 `CRS Share` permission level exists** and actually contains the sharing rights.
      Site Settings → Site permissions → Permission levels.
- [ ] **1.3 `Highly Confidential` exists as a TERM** in the confidentiality set (`0d6d1da8-…`).
      Suspected missing — the upload form offered only Confidential and Restricted on 2026-08-15.
      Without it the entire HC vertical (§7) is untestable.
- [ ] **1.4 `Documents` content approval is OFF.** Tell-tale without a query: the view bar shows
      `Approve/reject Items` + `Show All Files` when it is on. If on, every routed file is invisible
      to viewers and presents as a permissions bug that is not one.
- [ ] **1.5 Column parity across all FOUR libraries.** Diff `/fields?$select=Title,InternalName`.
      A matching display name over a different internal name fails exactly like an absent column and
      looks correct in the UI — and one unknown field name fails the WHOLE metadata write.
- [ ] **1.6 Audit log list provisioned** (open CRS Audit Log once; it self-provisions) and its
      permissions set by hand — provisioning deliberately does not touch them.

---

## 2. Foundations — cheap, and a failure here invalidates everything below

- [ ] **2.1 DMS Config `mode` rows load** — 5 segments expected (GHO, Minamas HO, NBPOL HO, Upstream
      Operations Malaysia, Industrial).
- [ ] **2.2 Allowed file types** — the form rejects `.exe`, accepts `.pdf`. Settings are read once at
      mount, so **every config change needs a hard refresh** before it can be tested.
- [ ] **2.3 Term sets resolve** — Document Type, Year and Confidentiality all populate.

---

## 3. Admin setup chain — the order the client will actually work in

Run this through the **guided flows** (Folder Administration → picker), since that is the new front
door and the thing the client said they could not operate.

- [ ] **3.1 CRS Settings landing page** — every row resolves to a real page. A row that cannot find
      its page must be disabled and name the page to create, never a dead arrow.
- [ ] **3.2 Back band** on all nine linked pages, returning to CRS Settings. **Not** on Upload Form,
      Approval Document, My Submissions or CRS Search.
- [ ] **3.3 Guided flow: "Add a department or unit"** — the everyday one. Name the subject, walk the
      rail. A lock must arm only on a fact that is KNOWN and unmet; anything unread is not a lock.
- [ ] **3.4 Term Abbreviations** — fill a missing code; a sibling collision blocks save and names the
      other term; "Same as term name" works.
- [ ] **3.5 Group Management** — create a group and add a member, then leave it **unmapped**. That
      state is the whole reason this page exists and could not be expressed before.
- [ ] **3.6 Folder Access** — map that group to a segment/tier/role by persona.
- [ ] **3.7 Reconciliation** creates the folder and grants the ACL. **(2-acct)** the new member can
      actually reach it.
- [ ] **3.8 Deleting a mapping row does NOT delete the group**, and the screen says folder
      permissions survive until reconciliation runs.

---

## 4. The core lifecycle — upload → approve → route

The part the client cares about most; everything else is scaffolding around it.

- [ ] **4.1 Upload a file as the PIC** **(2-acct)**. Metadata written; lands at the END of the chain.
- [ ] **4.2 Batched multi-file upload** — two batches to different destinations, several files each.
      Then the real test: edit the pickers so they describe batch 3, upload, and **batch 1 must still
      land where it was staged**.
- [ ] **4.3 Name composition** — `[Project] - [Vendor] - [Name] - [Date]`. Two files with different
      typed names but matching project/vendor/date must be caught as a collision.
- [ ] **4.4 Success removes, failure stays.** After a partial failure the list holds exactly what
      still needs doing, and Retry cannot double-send.
- [ ] **4.5 A peer PIC cannot see the pending file** **(2-acct)** — Draft Item Security.
- [ ] **4.6 The Head of Unit CAN see it, and approves it** **(2-acct)**.
- [ ] **4.7 Auto-route moves it** to `Documents`, keeps the uploader in Created By, and **deletes the
      source**. That delete is load-bearing for security, not housekeeping.
- [ ] **4.8 Folders created during upload arrive Approved**, or other PICs cannot navigate to their
      own files.
- [ ] **4.9 Reject a file** — the reason reaches the uploader.
- [ ] **4.10 Provisioned-path gating** — a unit the uploader cannot upload into does not appear at
      all; an admin sees it labelled *"not fully set up yet"* and can still select it.

---

## 5. The user-facing pages

- [ ] **5.1 My Submissions** **(2-acct)** — the PIC's own files across both libraries. Folders must
      NOT be listed. The detail panel shows that segment's own tier names.
- [ ] **5.2 CRS Search** — free text finds a recent upload; each dropdown filters; choosing a segment
      reveals its own tiers, and changing a parent clears its children.
- [ ] **5.3 CRS Search — the open question.** Search for something uploaded **a week or more ago**.
      Found ⇒ managed properties are fine, nothing to do. Only recent files ⇒ needs the
      `RefinableString` mapping in Site Settings → Search Schema plus a one-function change.
- [ ] **5.4 CRS Search scoping** **(2-acct)** — a PIC sees their unit; an account in a different unit
      sees neither that unit's documents nor any sign they exist.
- [ ] **5.5 Deletion request** — a PIC raises one on an approved file; the Head of Unit approves; the
      file is **recycled**, not purged. A failed action records `Failed`, never `Approved`.
- [ ] **5.6 Share request** — an approved share reaches the recipient; external is refused unless
      `allowExternalSharing` explicitly says yes. HTTP 200 does not mean it worked — the per-recipient
      result is in the body.
- [ ] **5.7 Per-person access removal** on both access screens. The "also in another group" note must
      appear BEFORE the click, and a removal that does not end access must say so.

---

## 6. Structure change — the highest-risk area

Pausing Auto-route and folder approval is **not required** (a moved pending file takes the False
branch), but it avoids hundreds of no-op runs burning the daily quota — and an exhausted quota means
the next real approval is not routed, silently.

- [ ] **6.1 Structure Manager** — add a level below Unit. It writes `PendingLevels`; `Levels` is
      untouched, so uploads keep using the old shape.
- [ ] **6.2 Subtree migration** — scan, review, typed `MOVE`, run. Files move, folders are
      ensure-created, metadata is re-stamped from the path.
- [ ] **6.3 It applies `PendingLevels` itself** as its last step, and only when a FRESH scan finds no
      drift left.
- [ ] **6.4 Stale-page guard** — leave an upload tab open across the change. The next upload from it
      must refuse with "reload the page" rather than filing into the old shape.
- [ ] **6.5 Add a segment** end to end, then **delete** it and confirm no document and no column was
      touched.

---

## 7. Highly Confidential — a whole vertical, with prerequisites

**Nothing here works until 1.3 passes and the setup below is done.** This is the least-tested thing
in the system, and the consequence of a mistake is the worst.

- [ ] **7.1** `CRS Folder` content type on both HC libraries, with `Full Name` on it.
- [ ] **7.2** Create `*_UPL_HIGHLY_CONFIDENTIAL` / `*_APR_HIGHLY_CONFIDENTIAL` for one unit and map
      them with the HC personas.
- [ ] **7.3** Re-run reconciliation — it must now build and grant **all four** libraries.
- [ ] **7.4** Build the two Power Automate flows **as the service account**: HC Auto-route with
      `{IsFolder}` = **false**, HC folder approval = **true**. The polarity is the whole thing, and
      both mistakes have already been made once on the normal pair.
- [ ] **7.5 `HC Documents` content approval OFF.**
- [ ] **7.6** A cleared uploader sees Highly Confidential and files into the HC library. **(2-acct)**
- [ ] **7.7 AN UNCLEARED UPLOADER DOES NOT SEE THE LEVEL AT ALL** **(2-acct)** — the most important
      single test in this plan.
- [ ] **7.8** HC files appear in CRS Search for the cleared user, marked HC, and not at all for the
      uncleared one — with **no error banner**, because a refusal is not a failure.
- [ ] **7.9** Approval and routing complete into `HC Documents`.
- [ ] **7.10** Note and decide: documents already labelled Highly Confidential sit in the NORMAL
      libraries and stay there. Migration is out of scope and must not be silently skipped.

---

## 8. Audit log

- [ ] **8.1** Admin actions write rows: policy change, abbreviation change, reconciliation run,
      access granted/revoked, segment created, structure changed, migration run, group map changed.
- [ ] **8.2** ONE row per run for reconciliation and migration, not one per folder.
- [ ] **8.3** A half-done permission change records as half-done, never flattened to success.
- [ ] **8.4** The viewer distinguishes "could not read" from "nothing matched".
- [ ] **8.5** Note: file-lifecycle events (Uploaded/Approved/Rejected/Routed/Deleted) come from flows
      that are **not built**. Until they are, an empty file history does NOT mean nothing happened —
      and the page must say so.

---

## 9. Cross-cutting regression

- [ ] **9.1 `+ New Folder`** on every CRS library, with upload options hidden behind it.
- [ ] **9.2 Dates read `DD/MMM/YYYY`** everywhere on screen.
- [ ] **9.3** Every page renders on a narrow window with no horizontal scroll.
- [ ] **9.4** No page shows a raw internal name, a term GUID or a raw byte count.
- [ ] **9.5** Every "could not read" state is distinguishable from "empty" — the most repeated rule
      in this codebase, and the one whose absence caused a duplicate upload once already.

---

## 10. Change register — fill in as we go

Anything to change, however small. A note here costs five seconds; a forgotten one costs a re-test.

| # | Where | What | Severity | Status |
|---|---|---|---|---|
| 1 | CRS Settings + the 9 pages it links to | **Any uploader can open every admin page.** `CRS_SITE_MEMBERS` grants site-level `Read` — it must, or a guest cannot reach Home — and every admin page inherits site permissions. Confirmed live 2026-08-16 with a guest holding only Read + Limited Access. Not a code bug: `pageAccessPolicy.ts` is a UI filter and says so; the boundary is the SharePoint page grant, and nothing performs it or prompts for it. Would have shipped to SDG identically. | **blocker** | open |
| 2 | `CrsSettings.tsx` | No `IsSiteAdmin` check, unlike `AuditLog.tsx` and `BulkUpload.tsx`. Defence in depth only — the real boundary is #1 — but it means the page lists every admin tool to anyone who reaches the URL. | major | open |
| 3 | Migration runbook | Per-page restriction is a required per-site step and is documented nowhere. | major | open |
| 4 | `shared/spGroups.ts` | **No throttle handling on any group operation. Add/remove member and create/delete group are single-shot, so a burst of writes gets a 429/503 and fails outright. Hit live 2026-08-16 removing one account from 8 groups. `dmsFolderMap.ts` retries 429/503 honouring `Retry-After` in six places — group operations are the one bulk-write path without it, and group creation in bulk is exactly what migration to SDG's tenant will do. | major | **fixed 1.0.122.0** |
| 6 | `FolderManager.tsx` reconciliation, page pass | **ROOT CAUSE OF #1.** Reconciliation does break inheritance on Site Pages and grant from Group Map rows — but it iterates only pages that HAVE grant rows. An `adminOnly` page has no eligible groups, so no rows, so reconciliation never visits it and it stays inheriting and open to every site member. The lock mechanism is driven by grants; admin pages by definition have none. | **blocker** | open |
| 7 | `FolderManager.tsx` library pass | **The site-entry grant is inverted on this site, and cannot self-heal.** The restore is conditional on the group having had access BEFORE the break, so whatever state existed at first break is frozen. Live: `Approval Document` HAS `CRS_SITE_MEMBERS` Read (the code comment says it must NOT) and `Documents` has NONE (the comment says it MUST). The block is guarded on `HasUniqueRoleAssignments !== true`, so re-running reconciliation skips it. Risk named in that same comment: the approval guard resolves the destination folder in Documents as the approver and needs that Read — without it approvers 404 on every destination while the run logs success. | **blocker** | **fixed 1.0.123.0** — reconciliation now asserts the state every run, in both directions |
| 5 | `shared/spGroups.ts` `fail()` | A throttled response body is HTML, and it is surfaced raw to the user — a wall of `<!DOCTYPE html>` where a sentence belongs. Should name the status and say "SharePoint is busy, retry shortly". | minor | **fixed 1.0.122.0** |

Severity: **blocker** (migration cannot proceed) · **major** (wrong behaviour, workaround exists) ·
**minor** (cosmetic, wording, layout).

---

## 11. Exit criteria — what "done" means before migrating

1. Every **blocker** and **major** closed, or explicitly accepted in writing.
2. §4 passes end to end with **real second accounts**.
3. §7.7 passes — an uncleared uploader cannot see the HC level.
4. `npx heft test --clean` green, warnings at the 19 baseline.
5. The migration runbook updated with everything this testing changed — especially the per-site setup
   steps: column parity, `Created By` index on `Documents`, audit list permissions, Search Schema
   mappings if 5.3 needs them, and the customizer registration if the feature does not provision.

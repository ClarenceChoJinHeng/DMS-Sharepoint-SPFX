# CRS — Full system test plan (pre-migration)

**Date started:** 2026-08-16 · **last updated 2026-08-19**
**Site:** `/sites/ClarenceDMSTesting`
**Package:** 1.0.170.0
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

- [x] **1.1 Re-run Folder Reconciliation.** **DONE 2026-08-18** — ~40 min, 2,706 steps, all four
      libraries. No `SKIPPED (no abbreviation)`, no collision aborts. A clean re-run afterwards
      created 0 folders, which is the other half of the proof. Not optional. The deletion/share workflow's folder-scope
      grants and the `CRS Share` level exist on no folder until it runs, and the one-group-per-person
      change (`UPL`/`APR` reading `Documents`) also lands here.
      *Expect:* a clean run — no `NO TERM`, no missing-abbreviation reports.
- [ ] **1.2 `CRS Share` permission level exists** and actually contains the sharing rights.
      Site Settings → Site permissions → Permission levels.
- [x] **1.3 `Highly Confidential` exists as a TERM** — confirmed 2026-08-19 in the site term store
      (`DMS` group -> Confidentiality Level -> Highly Confidential, alongside Confidential and
      Restricted). The whole HC vertical was blocked on this.
- [ ] ~~1.3 (old wording)~~ **`Highly Confidential` exists as a TERM** in the confidentiality set (`0d6d1da8-…`).
      Suspected missing — the upload form offered only Confidential and Restricted on 2026-08-15.
      Without it the entire HC vertical (§7) is untestable.
- [x] **1.4 `Documents` content approval is OFF.** **PROVEN 2026-08-18 the only way that counts** — a
      non-admin uploader opened `Documents/GHO/GF/TAX/2024/Tax Return` and saw the routed file. Visible
      to an admin proves nothing; admins see drafts. The `Show All Files` view on that library is a
      leftover from an earlier setting, not an active one.
- [ ] ~~1.4 (old wording)~~ **`Documents` content approval is OFF.** Tell-tale without a query: the view bar shows
      `Approve/reject Items` + `Show All Files` when it is on. If on, every routed file is invisible
      to viewers and presents as a permissions bug that is not one.
- [x] **1.5 Column parity — HC PAIR VERIFIED 2026-08-19** by `scripts/check-hc-setup.js`: all 21
      columns present on both, every one the right type, all three taxonomy columns bound to the right
      term sets. **Run that script rather than diffing XML by hand** — it also catches a
      `Confidentiality Level` created as plain Text and a column bound to the wrong term set, both of
      which look correct in the UI. The normal pair is unverified by the script but demonstrably works.
- [ ] ~~1.5 (old wording)~~ **Column parity across all FOUR libraries.** Diff `/fields?$select=Title,InternalName`.
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

> ✅ **4.1, 4.6 and 4.7 PASSED on 2026-08-18 with two guest accounts** — the first end-to-end pass on
> any site. Upload auto-detected `Group Head Office > Group Finance > Tax` from group membership and
> locked the tiers; the file landed Pending with every field; the approver saw it with its metadata and
> approved; Auto-route moved it to `Documents/GHO/GF/TAX/2024/Tax Return`, deleted the source, and
> **kept chocheetuck4 in `Created By`** with all metadata intact. The approval email carried the right
> path and a working link. The two steps that had never been proven are the uploader stamp and metadata
> surviving the copy; both hold.

- [x] **4.1 Upload a file as the PIC** **(2-acct)**. Metadata written; lands at the END of the chain.
- [ ] **4.2 Batched multi-file upload** — two batches to different destinations, several files each.
      Then the real test: edit the pickers so they describe batch 3, upload, and **batch 1 must still
      land where it was staged**.
- [ ] **4.3 Name composition** — `[Project] - [Vendor] - [Name] - [Date]`. Two files with different
      typed names but matching project/vendor/date must be caught as a collision.
- [ ] **4.4 Success removes, failure stays.** After a partial failure the list holds exactly what
      still needs doing, and Retry cannot double-send.
- [ ] **4.5 A peer PIC cannot see the pending file** **(2-acct)** — Draft Item Security.
- [x] **4.6 The Head of Unit CAN see it, and approves it** **(2-acct)**. Passed 2026-08-18.
- [x] **4.7 Auto-route moves it** to `Documents`, keeps the uploader in Created By, and **deletes the
      source**. That delete is load-bearing for security, not housekeeping. Passed 2026-08-18.
- [ ] **4.8 Folders created during upload arrive Approved**, or other PICs cannot navigate to their
      own files.
- [ ] **4.9 Reject a file** — the reason reaches the uploader.
- [ ] **4.10 Provisioned-path gating** — a unit the uploader cannot upload into does not appear at
      all; an admin sees it labelled *"not fully set up yet"* and can still select it.

---

## 4b. PAGE ACCESS — added 2026-08-19, after a full day lost to it

Every failure here presents as "the app is broken" and none of it is code. Test with a **real second
account**, and re-test after every reconciliation.

- [ ] **4b.1 An uploader can open the Upload Form.** Needs a **page-scope grant for that unit's
      `_UPLOADER` group**, made on Page Access. Verified 2026-08-18 for `GHO_GF_TAX_UPLOADER`.
- [ ] **4b.2 An approver can open the Upload Form too** — `APR` is on that page's policy since
      2026-08-17, because a Head of Unit uploads as well. Verified: the form auto-detected the path for
      an approver group whose rows are `APR + DELS + DEL + SHARE + UPLHC + DELSHC`.
- [ ] **4b.3 An uploader is DENIED on `ApprovalDocument.aspx`** — verified 2026-08-18. The asymmetry is
      deliberate and pinned by test; mirroring the two would let every PIC approve their own documents.
- [ ] **4b.4 A restricted page with NO grant rows locks everyone out silently.** Clearing
      `CRS Group Map` deletes the Page-scope rows while leaving inheritance broken, and reconciliation
      reports `Page access: no Page-scope mappings` as a green tick. Two user-facing pages were
      unreachable by everyone but Owners for hours. **Re-check after any Group Map reset.**
- [ ] **4b.5 A new group does NOT reach a session already open.** After adding someone to a group they
      may keep being bounced until they **sign out and back in** — every permission read back correct
      (membership, page assignment, `Read` binding, their own `currentuser/groups`) and it was the
      SESSION that was stale. Cost an hour on 2026-08-19.
- [ ] **4b.6 A gmail guest can exist TWICE for one address** (an OTP guest and a personal Microsoft
      account). Check `LoginName`, not the email, when "I added them and they still cannot get in".

---

## 4c. STALE CACHES — added 2026-08-19

Settings, library names and modes are read once in a mount-time effect. Everything here is a real
failure that looks like a bug in the feature you were testing.

- [ ] **4c.1 After renaming a library, hard refresh every open tab.** The name cache is a module-level
      promise primed once per page load. A tab that resolved `HCDocuments` before the rename kept using
      it and 404d on a library that exists — reported as *"could not read its columns"* on a screen
      that had nothing to do with names.
- [ ] **4c.2 After any config change, hard refresh** before concluding anything about it.
- [ ] **4c.3 Legally Privileged: the config row holds LEVEL NAMES.** Set
      `legallyPrivilegedFor = Confidential;Highly Confidential`; the tick must appear for BOTH. Until
      1.0.169.0 the check compared the dropdown's raw value — a term GUID — so the row had to hold a
      GUID to match anything at all, and no screen could explain it.

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

> **State on 2026-08-19:** libraries exist and pass the column/type/term-set check; the `CRS Folder`
> content type is attached to both and reconciliation has stamped the folders (`Full Name` shows
> `Group Head Office` on the GHO folder in both). **Both HC flows are built** with the right polarity.
> An HC upload now lands in `HC Approval Document` **with its metadata** (1.0.170.0 — before that it
> could never tag, on any site). What remains untested is everything about who can SEE it.

- [x] **7.1** `CRS Folder` content type on both HC libraries, with `Full Name` on it. Done 2026-08-19.
- [ ] **7.2** Create `*_UPL_HIGHLY_CONFIDENTIAL` / `*_APR_HIGHLY_CONFIDENTIAL` for one unit and map
      them with the HC personas.
- [x] **7.3** Re-run reconciliation — it must now build and grant **all four** libraries. Done
      2026-08-18; the HC folder tree exists and carries the folder content type.
- [x] **7.4 BUILT 2026-08-19** and both trigger conditions verified by eye: HC Auto Route `false`,
      HC folder approval `true`. ⚠ **Built under `clarence@trinergydigital.com`, not a service
      account** — acceptable here, MUST be the service account on CRS or both flows stop silently the
      day that password changes.
- [ ] ~~7.4 (original)~~ Build the two Power Automate flows **as the service account**: HC Auto-route with
      `{IsFolder}` = **false**, HC folder approval = **true**. The polarity is the whole thing, and
      both mistakes have already been made once on the normal pair.
- [x] **7.5 `HC Documents` content approval OFF.** Verified by script 2026-08-19.
- [x] **7.6** A cleared uploader sees Highly Confidential and files into the HC library. **(2-acct)**
      Passed 2026-08-19 — and **with metadata**, which is the part 1.0.170.0 fixed.
- [ ] **7.6b Approve an HC file and confirm HC Auto-route moves it to `HC Documents`** with
      `Created By` and metadata intact. The flows exist but have never fired.
- [ ] **7.7 AN UNCLEARED UPLOADER DOES NOT SEE THE LEVEL AT ALL** **(2-acct)** — the most important
      single test in this plan.
- [ ] **7.8** HC files appear in CRS Search for the cleared user, marked HC, and not at all for the
      uncleared one — with **no error banner**, because a refusal is not a failure.
- [ ] **7.9** Approval and routing complete into `HC Documents`.
- [ ] **7.11 THE HC PAIR IS NOT MIGRATED BY A STRUCTURE CHANGE — KNOWN BROKEN.** `SubtreeMigrator`
      walks a hardcoded two-element list (`Staging`, `Documents`) and `LibTarget` is typed to those two,
      so the HC libraries are invisible to it. Found 2026-08-19: after adding a level, `HC Approval
      Document/GHO/GF/TAX` holds BOTH `2025` (old shape) and `Archive 1` (new) while the normal pair was
      migrated cleanly. Same class as register #6 and #10 — a mechanism written for two libraries when
      there are four. See register #15.
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
| 1 | CRS Settings + the 9 pages it links to | **Any uploader can open every admin page.** `CRS_SITE_MEMBERS` grants site-level `Read` — it must, or a guest cannot reach Home — and every admin page inherits site permissions. Confirmed live 2026-08-16 with a guest holding only Read + Limited Access. Not a code bug: `pageAccessPolicy.ts` is a UI filter and says so; the boundary is the SharePoint page grant, and nothing performs it or prompts for it. Would have shipped to SDG identically. | **blocker** | **fixed 1.0.129.0** — see #6 |
| 2 | `CrsSettings.tsx` | No `IsSiteAdmin` check, unlike `AuditLog.tsx` and `BulkUpload.tsx`. Defence in depth only — the real boundary is #1 — but it means the page lists every admin tool to anyone who reaches the URL. | major | open |
| 3 | Migration runbook | Per-page restriction is a required per-site step and is documented nowhere. | major | open |
| 4 | `shared/spGroups.ts` | **No throttle handling on any group operation. Add/remove member and create/delete group are single-shot, so a burst of writes gets a 429/503 and fails outright. Hit live 2026-08-16 removing one account from 8 groups. `dmsFolderMap.ts` retries 429/503 honouring `Retry-After` in six places — group operations are the one bulk-write path without it, and group creation in bulk is exactly what migration to SDG's tenant will do. | major | **fixed 1.0.122.0** |
| 10 | `FolderManager.tsx` library pass | **BOTH HC libraries inherited site permissions — every site member could read them.** Found live 2026-08-17 by the new site-entry pass. Folders inside were locked, but the library root inherited site Read from `CRS_SITE_MEMBERS`, on the two libraries where that matters most. Cause: inheritance is broken inside the library-scope ROWS loop, and the HC libraries have no rows — the same structural gap as #6, where the mechanism is driven by grant rows and the thing needing protection has none. | **blocker** | **fixed 1.0.124.0** — an inheriting CRS library is now broken and Owners restored, not merely reported |
| 11 | `FolderManager.tsx` progress feed | **Library-scope grants were all reported under the Documents panel**, whatever library they targeted — so 17 `Staging → …` rows appeared under "Documents group assignments" while the Staging panel stayed empty. Spotted live 2026-08-17. The grants always landed correctly; only the reporting was wrong, on the screen whose job is showing what happened to which library. Page-scope grants still report under Documents by necessity — there is no page panel — but the line says "page scope". | minor | **fixed 1.0.125.0** |
| 8 | Below-Unit tiers (`Form.tsx`, `BulkUpload.tsx`, `StructureManager`, `DocumentSearch`) | **SubUnit must read its options from the chosen UNIT's child terms, not from a standalone term set.** Client's data: some units have subunits, some do not, and they are unit-specific — a flat set offers every unit the same list. Reconciliation already stops at `permissionedDepth` and its comment already describes below-Unit tiers as living in the term tree, so authoring under the unit is safe today; only the dropdown source is wrong. Fails SILENTLY — an empty dropdown, not an error. | major | open — spec then build after permissions testing |
| 9 | Reconciliation runtime | **~60 minutes, single tab, no resume.** Scales with folders × libraries; creating the HC pair doubled it (4 libraries now). Dies if the tab closes, with no way to tell how far it got and no "what remains" summary. Not viable to hand to SDG, whose site is larger. Cheap mitigation exists (`recon_gridMode` off) but is already off here — the cost is the per-folder permission work across four libraries. | major | open |
| 6 | `FolderManager.tsx` reconciliation, page pass | **ROOT CAUSE OF #1.** Reconciliation does break inheritance on Site Pages and grant from Group Map rows — but it iterates only pages that HAVE grant rows. An `adminOnly` page has no eligible groups, so no rows, so reconciliation never visits it and it stays inheriting and open to every site member. The lock mechanism is driven by grants; admin pages by definition have none. | **blocker** | **fixed 1.0.129.0** — a separate lockdown pass locks every `adminOnly` page to site Owners and strips every other grant, asserted each run; the row pass now refuses a row targeting one so the two cannot fight. Spec `2026-08-17-admin-page-lockdown-design.md`. **The coverage test written with it found FIVE admin pages the policy never matched** — Group Management (`group.?manager` does not match "Management"), Site Access, Page Access and CRS Audit Log (all fell to `DEFAULT_POLICY`), and Approval Library Access, which matched `/approv/i` and was classified as an *approver* page. Each would have been left open by a lockdown that looked like it worked. |
| 7 | `FolderManager.tsx` library pass | **The site-entry grant is inverted on this site, and cannot self-heal.** The restore is conditional on the group having had access BEFORE the break, so whatever state existed at first break is frozen. Live: `Approval Document` HAS `CRS_SITE_MEMBERS` Read (the code comment says it must NOT) and `Documents` has NONE (the comment says it MUST). The block is guarded on `HasUniqueRoleAssignments !== true`, so re-running reconciliation skips it. Risk named in that same comment: the approval guard resolves the destination folder in Documents as the approver and needs that Read — without it approvers 404 on every destination while the run logs success. | **blocker** | **fixed 1.0.123.0** — reconciliation now asserts the state every run, in both directions |
| 12 | `groupMapModel.ts` (`GroupName` column) + every screen and log that renders it | **A Group Map row caches the group's NAME, and a rename desynchronises it forever.** The row stores the group's integer `GroupId` — which is what reconciliation grants against, so grants are always correct — plus a `GroupName` snapshot taken when the mapping was created. Nothing refreshes it. Found live 2026-08-17: Folder Access, Group Management's filter and every reconciliation log line named `DMS_GHO_GF_CORU_UPL`, a group that **does not exist**; the real group is `GHO_GF_CORU_UPL` (id 53), renamed in the `DMS_` → `CRS` pass. The grant was right the whole time and the entire audit surface was lying. **Migration makes this universal, not incidental** — SDG renames every group at import (memory `dms-to-crs-rename-pending`), so every row on every screen and every `↳ … → CRS Upload` line will name a group nobody can find, across 50+ groups, read by people who do not know the system. Cost 20 minutes to diagnose here with full context. Fix: resolve the display name from the live group by **id** at render/log time; fall back to the stored name only when the id does not resolve, and SAY so — a deleted group and a renamed one need opposite fixes (gotcha #9's rule again). | major | open |
| 5 | `shared/spGroups.ts` `fail()` | A throttled response body is HTML, and it is surfaced raw to the user — a wall of `<!DOCTYPE html>` where a sentence belongs. Should name the status and say "SharePoint is busy, retry shortly". | minor | **fixed 1.0.122.0** |

| 13 | `FolderManager.tsx` page pass + `pageAccessPolicy.ts` | **Page access is not derived from anything — it needs a Page-scope Group Map row per group, per page.** The policy file is a UI filter over a manual action and says so, but its name and CLAUDE.md both read as though a role grants a page. Consequences seen live 2026-08-18: `Upload-Form.aspx` and `ApprovalDocument.aspx` sat restricted with NO grants after `CRS Group Map` was cleared, so only Owners could open them, and reconciliation logged `Page access: no Page-scope mappings` as a green tick. Doing it properly by hand means ~120 rows here and ~260 on CRS — every one saying the same thing, because the page needs ROLE granularity, not unit granularity. Client chose to keep the pages restricted rather than let them inherit. **Agreed fix (option B):** reconciliation derives these grants from the Folder rows' roles, so a page can never be stranded and a new unit needs no manual step; plus a `⚠ restricted but nothing grants access` warning in the same pass. | **blocker** | open — agreed, not yet built |
| 14 | `Form.tsx` `uploadStagedFile` | **HC uploads could never tag, on any site.** The file is placed by FOLDER ID (`GetFolderById` reaches into any library, so it landed in `HC Approval Document` correctly) while the tagging call named the library by TITLE and was hardcoded to `settings.stagingLibrary` — the NORMAL approval library. Item ids are per-list, so `validateUpdateListItem` looked for the new id in the wrong list: 404 on a good day, and on a bad one it finds a DIFFERENT document with that id and writes the metadata onto it, reporting success. Symptom was the generic *"Uploaded, but tagging metadata failed"*. `BulkUpload` already did this right via `targetListTitle()` — two upload paths, one pattern, one written by hand. | **blocker** | **fixed 1.0.170.0** — the target title is a parameter, the HC branch passes `libApiTitle("StagingHC")`, and the failure now names the STATUS and the LIBRARY |
| 15 | `SubtreeMigrator.tsx` | **A structure change migrates only two of the four libraries.** Hardcoded `[Staging, Documents]` with `LibTarget` typed to match, so the HC pair keeps the old shape while the normal pair is migrated and `Levels` goes live — one segment with two shapes, in the libraries where a mistake matters most. Found 2026-08-19. `allLibraryTitles()` exists precisely for this and its own comment warns that a two-element literal is "the bug, waiting". | **blocker** | open |
| 16 | `Form.tsx`, `BulkUpload.tsx` | **Legally Privileged compared the confidentiality TERM ID, not the label**, so `legallyPrivilegedFor` had to hold a GUID to match anything — while HC routing resolved the label first. Two rules over one dropdown wanting opposite shapes in config, with nothing on screen able to explain either. Client asked for the tick on Highly Confidential and no config value could have delivered it. | major | **fixed 1.0.169.0** — both compare the label; the row now holds LEVEL NAMES and accepts a `;`/`,` list |
| 17 | `CRS Config` | Two stale rows: **`term_highlyConfidential`** (a GUID) is read by NO code — HC routing uses `hcConfidentialityLevel`, absent here, and correctly defaults to the literal `Highly Confidential`; and **`stagingLibrary = Staging`** is superseded and would 404 if anything trusted it (the live title wins). Harmless, and exactly the sort of row that misleads the next reader. | minor | open |
| 18 | Reconciliation | **A full run re-walks every segment across four libraries — ~40-60 min — even when one unit changed.** Client's request 2026-08-19: pick which business segments to run, and ideally be TOLD which ones changed. Change detection is answerable from data already loaded (term tree vs Folder Map vs abbreviation rows) but must **recommend, never restrict** — grants drift for reasons the term tree cannot see, so a full run stays one click away. | major | open — spec next |

Severity: **blocker** (migration cannot proceed) · **major** (wrong behaviour, workaround exists) ·
**minor** (cosmetic, wording, layout).

---

## 11. Exit criteria — what "done" means before migrating

1. Every **blocker** and **major** closed, or explicitly accepted in writing. **Open blockers as of
   2026-08-19: #13 (page grants not derived) and #15 (HC not migrated).**
2. ✅ §4 passes end to end with **real second accounts** — done 2026-08-18.
3. §7.7 passes — an uncleared uploader cannot see the HC level. **Still the single most important
   untested thing in the system.**
4. §7.6b passes — an HC file actually routes to `HC Documents`.
5. §4b passes — page access, with a second account, re-checked after a reconciliation.
6. `npx heft test --clean` green, warnings at the 15 baseline (was 19; the count fell as files were
   split out, not by suppressing anything).
5. **The migration runbook — `docs/2026-08-17-sdg-migration-runbook.md`** — updated with everything
   this testing changed — especially the per-site setup
   steps: column parity, `Created By` index on `Documents`, audit list permissions, Search Schema
   mappings if 5.3 needs them, and the customizer registration if the feature does not provision.

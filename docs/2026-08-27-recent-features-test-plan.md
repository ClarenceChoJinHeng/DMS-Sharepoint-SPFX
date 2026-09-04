# CRS — Test plan for the last two sessions' work

**Date started:** 2026-08-27 · **Package: 1.0.305.0** · **Site:** `/sites/ClarenceDMSTesting`

**Scope:** the submission record (built 2026-08-27, never run on a site) and the items from the
1.0.304.0 handoff. **Not** a full system test — `docs/2026-08-16-full-system-test-plan.md` still
covers that, and nothing here replaces it. That file is pinned at 1.0.170.0 and mentions none of this.

**How to use this:** top to bottom. The order is by DEPENDENCY. §1 gates everything — the submission
record does not exist on a site until reconciliation has run, so testing §2 before §1 tests nothing.
Findings go in §7, not in your head.

**(2-acct)** marks a test that CANNOT be ticked from an admin session. An admin sees everything, and
the failures that matter in §2.9, §2.10 and §3.3 are all *over*-exposure, which your own session
cannot show you.

---

## 0. Before starting

- [ ] **0.1 Second account ready.** A non-admin in a unit's `_UPLOADER` group. Several tests below are
      meaningless without one. A guest account has worked for this since 2026-08-08.
- [ ] **0.2 An HC-cleared and an UNCLEARED uploader.** §3.2 and §3.3 need both, and "uncleared" means
      genuinely in no `_UPLOADER_HIGHLY_CONFIDENTIAL` group.

---

## 1. Prerequisites — these gate everything below

- [ ] **1.1 Deploy 1.0.305.0, then UPDATE the app in Site Contents.**
      One package supersedes 301–304; 301 was never deployed and 302/303 are folded in.
      ⚠ Feature elements run on install/update only, and the bundle refreshes independently — so web
      parts can appear updated while nothing else did.
- [ ] **1.2 Close every CRS tab and reopen.** A tab open across a deploy holds the old bundle and
      requests a content hash that no longer exists. The symptom is a dead web part reading
      `Something went wrong` / `ERROR: [object Object]`, with the real cause only in the console as a
      MIME-type refusal. Not a code defect — a stale tab.
- [ ] **1.3 Run Folder Reconciliation.** **Nothing in §2 works before this.** It creates the
      `CRS Submissions` list, its 11 columns, the `Author` index, the `SubmissionFileId` column on
      every library, and the two-list ACL grants.

      Watch the log for these, and record what you actually saw in §7:

      | Expect | Meaning |
      |---|---|
      | `CRS Submissions: list created ✓` | first run only; absent on re-runs |
      | `CRS Submissions: N column(s) created ✓` | expect 11 on the first run, 0 after |
      | `submission reference / bulk import column(s) created ✓` | `SubmissionFileId` added per library |
      | `CRS Requests: N group(s) granted, M already correct` | the option-2 per-group grants |
      | `CRS Submissions: N group(s) granted, M already correct` | same, second list |
      | `… REMOVED (N binding(s)) — no longer readable by every site member` | **the revoke** |

      ⚠ **On this site the revoke is the line to check hardest.** The requests grant was added by hand
      earlier, so the grant path may report `already correct` — but the revoke has **never fired
      anywhere**. If instead you see `still readable by every site member — … LEFT IN PLACE`, read the
      reason: no groups derived, or a grant failed. Both are deliberate refusals, not bugs.

- [ ] **1.4 No `"Created By" could not be indexed` warning.** If it appears, the list works until it
      passes 5,000 rows and then My Submissions starts failing on it. Worth fixing before SDG.
- [ ] **1.5 A second reconciliation run creates nothing.** `0 column(s)`, no list created, grants all
      `already correct`. That idempotence is half the proof.

---

## 2. The submission record — never run on a site

All new in 1.0.305.0. Spec: `docs/superpowers/specs/2026-08-27-submission-record-design.md`.

- [ ] **2.1 A row is written on upload.** Upload one file through the upload form, then open the
      `CRS Submissions` list directly. Expect one row: `SubmissionFileId` like `SFI-20260827-XXXX`,
      `Source` = `Form`, `LibraryTitle` = the approval library, `ItemPath` the full server-relative
      path, `MetadataSnapshot` holding readable JSON.
      ⚠ **Read the LIST, not the page.** A blank stamp is invisible on the upload form — the upload
      succeeds either way, which is the design.
- [ ] **2.2 The upload still succeeds if the record cannot be written.** No way to stage this cleanly;
      confirm the negative instead — nothing in the upload's own messages mentions submissions, and the
      console shows at most a `[submissions]` info line. The record must never be able to fail an
      upload.
- [ ] **2.3 ⚠⚠ THE ONE THAT MATTERS — an APPROVED file still resolves.** Upload, approve it, wait for
      Auto-route (~18s plus flow polling), then reload My Submissions.
      **Expect: the row reads `Approved` and is NOT greyed.**
      This is the test the whole design correction exists for. Auto-route copies then deletes, so the
      routed file carries a **new `UniqueId`** — the key the agreed design would have joined on. If
      this row reads **Deleted**, the stamp is not surviving the routing copy and the feature is
      inverted: every approved document reported as destroyed. Stop and report it.
- [ ] **2.4 A deleted file keeps its row.** Recycle an uploaded file from the library, reload My
      Submissions. Expect the row still listed, **greyed**, badge **`Deleted`**, the name as plain text
      with no button, no `Open file`, and the action cell reading `no longer here` with no
      Delete/Share buttons.
- [ ] **2.5 The note appears once, and reads calmly.** Above the table on **All** or **Submissions**:
      *"N file(s) you uploaded are no longer in the library…"*. Grey, not red — a deleted document is a
      fact about the past, not a failure of the page.
- [ ] **2.6 A deleted row is on the right tabs and NOT on the wrong ones.**
      - [ ] **Submissions** — present, inside its own batch
      - [ ] **All** — present
      - [ ] **Pending / Approved / Rejected** — **absent from all three**

      ⚠ If a deleted file appears under **Pending**, the status placeholder is leaking and the page is
      telling someone a destroyed file awaits approval.
- [ ] **2.7 The tallies ignore it.** Tab counts and the batch's own `N approved / N pending` line must
      not count the deleted file. **Submissions** and **All** counts DO include it.
- [ ] **2.8 A deleted file's details come from the snapshot.** Open its submission → its batch. Expect
      the card greyed, no `Open file`, and metadata rows populated from the record — Unit, Year,
      Document name, Confidentiality. It must **not** read *"No details were recorded for this file"*:
      that message means the snapshot is not being read.
- [ ] **2.9 (2-acct) One uploader cannot see another's records.** As the second account, open My
      Submissions and confirm only their own files appear. The list holds everyone's rows; the page
      filters `AuthorId eq me`.
- [ ] **2.10 (2-acct) A plain site member cannot read the list.** After 1.3's revoke, open the
      `CRS Submissions` list via Site Contents as someone in **no** uploader/approver/HoD group.
      Expect **access denied**. If they can read it, the revoke did not happen — re-read the §1.3 log.
- [ ] **2.11 Bulk Upload writes records too.** Import two files. Expect two rows with `Source` =
      `BulkUpload`, **different** `SubmissionFileId` values, and the same `SubmissionRef`.
      ⚠ **Different stamps per file is the specific thing to check.** Bulk Upload shares one
      `formValues` array across the whole run, so a stamp added in the wrong place would put one id on
      every file and collapse every record onto a single document.
- [ ] **2.12 Files uploaded BEFORE today still behave as they did.** An older file has no record, so
      deleting it makes its row vanish exactly as before. **Expected, not a bug** — nothing recorded
      them and no retro-fix exists. The client will test this first; tell them beforehand.

### Not practical to stage, but know the symptom

**If every row reads `Not checked` instead of `Deleted`,** the live reads could not return
`SubmissionFileId` — the three-rung ladder fell back, or an HC/archive library read failed.
Deliberate: the page refuses to claim a deletion it could not verify. Check the console for a
`[submissions]` line, then confirm §1.3 created the column on every library.

---

## 3. From the 1.0.304.0 handoff

Cheapest first, as that handoff had it. The last is the real change.

- [ ] **3.1 The Add box refuses an unknown address in plain English.** Group Management → type
      `minecraftx@gmail.com` → Add. Expect a readable refusal naming both causes (nobody has that
      address / more than one account matches), **beside the field** rather than as a toast — and the
      suggestion dropdown must **close** so the message can be read.
      ⚠ `ensureuser` **resolves, it never invites.** A correction from that session: this failing is
      not about a sharing setting.
- [ ] **3.2 The HC dropdown is stable on both upload screens.** Select a unit, change a below-Unit tier
      several times, reload, repeat — on the **upload form** and on **Bulk Upload**. The Highly
      Confidential level must appear and stay for a cleared uploader.
- [ ] **3.3 (2-acct) Bulk Upload offers an uncleared uploader NO HC level.** **New behaviour there** —
      Bulk Upload had no clearance probe before 1.0.304.0. Test as a genuinely uncleared account.
- [ ] **3.4 A CRS Owners member can decide every request.** As a non-SCA member of `CRS Owners`, open
      Requests. Expect the amber *"system administrator"* note, other people's pending requests under
      **Waiting for you**, and the misleading *"not recorded as the approver"* banner **gone**.

---

## 4. Never exercised — and the client is relying on the first

- [ ] **4.1 ⚠ Auto-route's Replace writes a VERSION, not a destruction.** Send a document through
      approval as a replacement for one already filed, approve it, then open **Version History** in
      `Documents`. **Expect 2.0, with the old content still at 1.0.**
      If it comes back as a single 1.0, the copy is replacing the ITEM rather than versioning it and
      **the previous content is gone** despite versioning being on. That is the client's own
      `nameConflictBehavior: 1` decision, so it must be proven rather than assumed.
      Versioning confirmed on `Documents` 2026-08-27: on, major, 500 kept.
      ⚠ This now has a second dependant: the submission record's stamp is **reassigned** by a Replace,
      so the replaced file's row reads `Deleted`. If Replace destroys rather than versions, that row is
      recording real data loss.
- [ ] **4.2 `HC Documents` versioning is UNCONFIRMED.** Same page, same check. Neither library on
      SDG's tenant has been checked at all.
- [ ] **4.3 Removing yourself from `CRS Owners` keeps your SCA.** Expect the removal to happen and the
      toast to say the demotion was skipped, naming why. Demoting yourself removes the very right the
      call needs, leaving a half-changed state.
- [ ] **4.4 Removing the LAST administrator is refused.** Expect a refusal that says so. A site
      collection with no administrator cannot be administered, and nothing in this app could put one
      back — only a tenant admin.
- [ ] **4.5 Promote the remaining `CRS Owners` members to SCA.** They predate the coupling, Crystal's
      `ckyan90@live.com` included.

---

## 5. What cannot be tested from the app at all

Stated so nobody spends an afternoon looking for a screen.

- **Power Automate flow configuration.** Fourteen flows, none in source control.
  `nameConflictBehavior`, trigger conditions and library references are visible only in the flow's own
  Code view.
- **Whether the RIGHT people are in a group.** Existence is checkable; intent is not. Every tick in
  this plan means "this exists", never "this is correct".
- **Reads and downloads.** Unrecordable by us — Purview is unreachable from a site-collection package.
- **Who DELETED a document, from My Submissions.** The record says who *uploaded* it and that it is
  gone. The **CRS Audit Log** answers who deleted it (`ActorName` from `DeletedByUserName`, path
  recovered from the recycle bin). ⚠ The two cannot be joined by GUID — the delete trigger returns no
  `UniqueId` — so a deletion is matched to its upload by filename, path and time.

---

## 6. Verified log

Only what was actually confirmed, and how. **Nothing in §2 is here yet.**

| Date | What | How confirmed |
|---|---|---|
| 2026-08-27 | Submission record builds clean at 1.0.305.0 | `tsc --noEmit` clean; 1452 tests, 0 failures; no new lint warnings beyond the MySubmissions line count |
| 2026-08-27 | `PRIMED_SUFFIXES` covers `submissions` | `naming.test.ts` pins both directions — adding the suffix unprimed fails the build |
| 2026-08-27 | The stamp joins across libraries in principle | pinned by test (`matches a resolved record across libraries`); **the live routing case is §2.3 and is NOT confirmed** |
| 2026-08-27 | `Documents` versioning on, major, 500 kept | read on site |
| 2026-08-23 | Audit deletion flows record the deleter | verified live per `docs/superpowers/specs/2026-08-23-audit-log-flows-runbook.md` |
| **2026-08-27** | **§1.1 1.0.305.0 is deployed on ClarenceDMSTesting** | the `CRS Submissions` list exists with its columns, and only 1.0.305.0 creates it |
| **2026-08-27** | **§1.3 full reconciliation completed clean** | 1,643 steps, all six libraries, **Warnings (0), Errors (0)**, *"every folder got a group"*. Closing passes all ✓: site-entry member added, Folder Map no orphans, Term Abbreviation no orphans, every folder maps to a live term |
| **2026-08-27** | **§1.4 `Author` indexed on the submissions list** | implied by Warnings (0) — a failure emits `⚠ "Created By" could not be indexed` |
| **2026-08-27** | **The submissions list has its columns, correctly typed** | `SubmissionRef` (SP.FieldText), `SubmissionFileId` (SP.FieldText), `MetadataSnapshot` (SP.FieldMultiLineText), read from `/fields` |
| **2026-08-27** | **⚠⚠ THE REVOKE FIRED — first live execution anywhere, on BOTH lists** | `roleassignments/getbyprincipalid(1581)` returns `Can not find the principal` on `CRS Requests` **and** `CRS Submissions`. Per-group `CRS Request` grants confirmed present in the UI (`_UPLOADER`, `_UPLOADER_HIGHLY_CONFIDENTIAL`, `_APPROVER`, `_APPROVER_HIGHLY_CONFIDENTIAL` per unit, `_HOD` per department) |
| **2026-08-27** | **Site Pages is correct on BOTH sites** | ClarenceDMSTesting: reset to inherit, `CRS_SITE_MEMBERS` Read, and **all 14 page items kept their own ACLs** (`HasUniqueRoleAssignments: true` on every admin page; only `Home.aspx`/`CollabHome.aspx` inherit). SDG `/sites/CRS`: inherits, Read present |
| **2026-08-27** | A list-scope inheritance reset does NOT cascade to items | observed directly, above — correcting a CLAUDE.md caution that said it "risks" them and had never been tested |
| **2026-08-27** | SDG's interrupted run left production safe | 13 groups granted `CRS Request` on `CRS Requests`; `CRS_SITE_MEMBERS` still holds Read; revoke never fired. Strictly better than before the run, and unfinished rather than broken |

| **2026-08-27** | **§2.1 records are written, from a non-admin uploader's session** | 3 rows in `CRS Submissions`, uploaded as `chocheetuck4` (`CRS_SITE_MEMBERS` + `GHO_GCA_EG_UPLOADER`, "Upload only"). `SubmissionRef`/`BatchRef`/`SubmissionFileId` correctly formatted, `LibraryTitle` = `Approval Document` (live title, not legacy `Staging`), `UploadedBy` lower-cased, `Source` = `Form`. A POST to `items` returning **201** seen in the Network tab |
| **2026-08-27** | **✅✅ §2.3 THE STAMP SURVIVES AUTO-ROUTE — the design correction validated on live data** | `TEST - Clarence - CHO - 27-08-26.pdf` recorded as `SFI-20260827-BGWU` against `Approval Document`, approved, routed. Now `Documents` item **6793** at `/Shared Documents/GHO/GCA/EG/2024/Agreement/`, carrying `SubmissionFileId = SFI-20260827-BGWU`. **The new item id proves a COPY, so its `UniqueId` is new — on the agreed key this file would have read `Deleted`** |
| **2026-08-27** | **A non-admin PIC can still raise a request after the revoke** | deletion request raised as `chocheetuck4`, sits Pending with a Cancel button |
| **2026-08-27** | The already-asked guard works | a file showing `Approved` **and** `deletion asked` at once, with `asked` replacing the buttons |

**Still NOT confirmed:**

- **§2.4 — deleting a recorded file and confirming its row greys out rather than vanishing.** **This is now the main outstanding test of the feature.** §2.3 proved a record survives *routing*; §2.4 is what proves it survives *deletion*, which is the client's actual requirement. Records 2 and 3 are available for it.
- **§2.6 / §2.7** — a deleted row's tab placement, and its exclusion from the tallies.
- **§2.11** — Bulk Upload writing records with a distinct stamp per file.
- **§1.5** — a second reconciliation run creating nothing (only one run so far).
- **§4.1** — whether Auto-route's Replace versions or destroys.

**⚠ Known display state, not a defect:** a file approved but **not yet routed** reads `Approved` with the *approval-library* path, so once Auto-route deletes the source the preview 404s and the details read *"No details were recorded for this file"*. Seen live; self-corrects on reload.

**⚠ An empty `CRS Submissions` read is not evidence of failure.** One was read that way on 2026-08-27; the query had run four minutes before the upload it was checking. Check the clock, and look for a `201` on `items` in the Network tab.

⚠ **A row here means it was observed.** Everything else is unproven, however confident the code
comments sound — this project's standing rule.

---

## 7. Findings

One line each, as you go. Anything needing a change goes here rather than being remembered.

| # | Test | What happened | Action |
|---|---|---|---|
|  |  |  |  |

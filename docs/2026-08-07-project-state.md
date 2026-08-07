# Project State — 2026-08-07

**Read this first in a new session.** It records what is true on the test site
`/sites/ClarenceDMSTesting` right now, what changed on 2026-08-06/07, and what is still
outstanding. CLAUDE.md holds the rules; this file holds the *state*.

Branch `feat/folder-abbreviations`, pushed to `origin`. Package built at
`sharepoint/solution/sd-gatrie.sppkg` v1.0.84.0. **Not yet deployed.**

---

## 1. The site was substantially rebuilt

### 1.1 Lists — all four now CRS-named

| List | URL | Items | Note |
| --- | --- | --- | --- |
| `CRS Config` | `/Lists/Config` | 14 | **Renamed**, not recreated — hence the short URL |
| `CRS Folder Map` | `/Lists/CRS Folder Map` | 0 | Recreated. Empty by design; reconciliation rebuilds it |
| `CRS Group Map` | `/Lists/CRS Group Map` | 1 | Recreated. Holds one row (see §1.4) |
| `CRS Term Abbreviation` | `/Lists/CRS Term Abbreviation` | **175** | Recreated + re-imported, verified field-by-field |

The old `DMS *` lists were **deleted**. The 175 abbreviation rows were re-imported from
`DMS Term Abbreviation.csv` (kept untracked in the project root) with a browser-console script,
then verified: 0 missing, 0 duplicate GUIDs, 0 field mismatches, and all **61 fullwidth `＆`**
preserved. 42 Department + 133 Unit.

**Recreating a list changes `ListItemEntityTypeFullName`; renaming does not.** `FolderMap.tsx`
reads it off the list at runtime for that reason — a hardcoded value breaks every `__metadata`
write. That is the only file writing `__metadata`.

### 1.2 The library was deleted and recreated

`Staging` is **gone** — deleted, along with its 182 test items and every folder ACL. Replaced by:

- **Title:** `Approval Document`
- **URL:** `/ApprovalDocument`

Created without the space so the URL stayed clean, then retitled. **They differ, and both are
needed** — see CLAUDE.md gotcha #12 and `naming.ts`. Verified state:

- All **15 columns** present with correct internal names, including `Full_x0020_Name` (the old
  library had `FullName0`; `Documents` still does)
- `CRS Folder` content type attached, `Full Name` on it, hidden from the New button
- Content approval **ON**, Draft Item Security **"Any user who can read items"**
- Permission inheritance **broken** at library level; only `CRS Owners` (Full Control)
- **0 items** — no folders yet, nothing reconciled

### 1.3 Content type

`CRS Folder` created 2026-08-07 as a **site** content type, group `CRS Content Types`, parent
group **Folder Content Types**, parent **Folder**. `DMS Folder` still exists and **cannot be
deleted** until reconciliation re-stamps the `Documents` folders — SharePoint blocks deleting a
content type still in use.

### 1.4 `CRS Group Map` holds exactly one row

| Title | GroupId | Segment | UnitTermGuid | Role | Scope |
| --- | --- | --- | --- | --- | --- |
| `DMS_GHO_GF_CORU_UPL` | 53 | `08dd94cb-…` | `3be3e50c-…` | UPLOADER | Folder |

Deliberately one row: a small blast radius for the first reconciliation on a library that has
never been reconciled.

**The old Group Map had 21 rows and they were NOT restored.** Only one was real permission data;
the other 20 were `ENTRY`/`Page` rows, four of which pointed at SharePoint **system principals**
(`Limited Access System Group`, `SharingLinks.…`), and several at MEMBER/GLOBAL/SEGVIEW groups
that the page policy now excludes. Page access is to be re-added through the Page Access tab.

That single row produced **two** silent-failure findings on 2026-08-07. Both are fixed; both are
worth knowing, because neither would have raised an error.

**(a) The `UnitTermGuid` was hand-typed and wrong — FIXED.** It held
`9aef23bf-62a9-4ff9-aaad-d02fc90e5eaf` (`Group Finance`, a **Department**) instead of
`3be3e50c-b90a-4187-9915-85f4a5f1eb49` (`CORU`, the **Unit**). With `recon_departmentFanOut`
defaulting ON, that would have fanned the uploader grant to **all nine units under GF**.

No safety net catches this: orphan repair only fires on a GUID **absent** from the term store,
and this one was perfectly valid — just wrong. Only the group *name* disagreed with it, and
nothing compares those. **Never hand-type a term GUID; copy it.**

**(b) The `Role` column read `UPLOADER`, which the code did not recognise — FIXED IN CODE.**
`ROLE_TO_PERMISSION` keys on the short codes (`UPL`, `APR`, …). The long-form suffixes belong to
group *names* (`_UPLOADER`), and an admin filling the column in by hand naturally matched them.
`accepts()` found no permission level, so the row was **skipped**, and the log would have said
*"no group-map groups for this unit (locked admin-only)"* — indistinguishable from a missing row.

`normalizeRoleValue()` in `groupMapModel.ts` now accepts both forms (and tolerates casing,
spaces and hyphens), wired into all six read sites across `FolderManager`, `GroupMapBuilder` and
`StagingAccess`. An unrecognised value passes through unchanged so it still fails `accepts()` —
defaulting it to anything would grant access nobody asked for.

---

## 2. Code changes, 2026-08-06/07

Commits `9167b9f`, `34eeabb` (plus `8aaa234`, docs only).

### 2.1 Library name resolution (`9167b9f`)

`"Staging"` was hardcoded in ~15 places as **both** a title and a URL segment. Now resolved as a
**pair** by `primeNames()` → `libraryTitle()` / `libraryUrlSegment()`, probing
`Approval Document` → `ApprovalDocument` → `Staging`, reading both halves from one response.

Why a pair: a wrong title 404s loudly; a wrong URL segment fails **silently** —
`split("/Staging/")` on a non-matching path returns a one-element array, so the caller reads
`undefined` and routes the file nowhere while reporting success.

`LibTarget`/`TARGETS` keep `"Staging"` as a **logical key** — it is also the stored `Target`
value on Group Map library-scope rows — and translate at the API boundary via `libApiTitle()`.
**This is why reconciliation logs still read `Staging/GHO/…`.** Those are labels, not paths.

The folder content type is likewise resolved: `CRS Folder` then `DMS Folder`, preferring the
first.

### 2.2 Six-persona role model (`34eeabb`)

See `docs/superpowers/specs/2026-08-07-role-model-simplification-design.md`. Twelve personas → six.
Approval moved from Head of Department to Head of Unit; HoD is now view + delete-in-Documents.
`SEGVIEW` un-retired for the per-segment C-Level. PIC's Highly Confidential variant removed —
confidentiality is metadata, not a permission.

### 2.3 The navigation corridor was restored — this one nearly bit

On 2026-08-04 the client said "Head of Unit cannot see the department folder". That was
implemented: the **ancestor-Read grant was deleted** and replaced with an opt-in revoke pass. The
client then **withdrew** the requirement.

Deleting the grant was invisible on an existing site, because reconciliation only ever added and
earlier runs' grants were still in place. **It bites on a new library** — every folder is new,
none gets ancestor Read, and the library root renders **empty for every non-admin**.

The grant is restored (idempotent), and `recon_revokeAncestorRead` is **hard-off with its config
row ignored** — with granting back, the two passes would fight.

`recon_departmentFanOut` now **defaults ON**, required by the department-scoped HoD persona.

---

## 3. Outstanding — in order

1. ~~Fix `CRS Approve`~~ — **VERIFIED CORRECT 2026-08-07.** It grants **View Items, Edit Items,
   Approve Items, Open Items** and *not* Add, Delete or Manage Lists. Exactly what draft-item
   security needs, and Delete correctly withheld so approving is not deleting.
   > Decoding note, because this cost a false alarm: in a `BasePermissions` mask the **bit index
   > is `PermissionKind - 1`**, not `PermissionKind`. Read off by one and the result is plausible
   > but shifted — View reads as Add, Edit as Delete. If you decode a mask by hand, sanity-check
   > it against the Permission Levels UI before acting on it.
2. **Deploy** v1.0.84.0. Bump to 1.0.85.0 and rebuild if the site keeps serving the cached bundle.
3. **`Documents` library** — attach `CRS Folder`, add `Full Name` to it.
4. ~~Verify the Group Map row's `UnitTermGuid`~~ — **DONE**, see §1.4.
5. ~~Reconcile one unit~~ — **DONE 2026-08-07, and it passed.** The log showed
   `↳ DMS_GHO_GF_CORU_UPL → CRS Upload`, `↳ … Read (browse) on /GHO` and `on /GHO/GF`, and
   `content type → CRS Folder` across every Documents folder. Only CORU was granted, confirming
   the term-GUID fix; the row being honoured at all confirms the Role normaliser.
   > **Folders came out Pending** (`OData__ModerationStatus = 2` on all of them). Harmless while
   > Draft Item Security is "any user who can read items", but under approver-only every folder
   > would vanish for every non-approver. Reconciliation now approves the folders it provisions
   > — only in libraries that actually moderate, and only when not already approved.
   >
   > ✅ **VERIFIED WORKING 2026-08-07 06:31**, after deploying the build with the fix. Querying
   > `$filter=FSObjType eq 1 and OData__ModerationStatus ne 0` returned just **two** rows —
   > `2024` and `Term Sheet` — so every structural folder is now Approved.
   >
   > ⚠ Those two survivors are the point: they are **Year / Document Type folders created by an
   > upload**, and reconciliation will never approve them. `recon_gridMode` is **off** on this
   > site, so the grid is not pre-created — the upload form mints these on first use, and the
   > caller is a PIC who **cannot approve** (`OData__ModerationStatus = 0` needs `ApproveItems`).
   > Turning `recon_gridMode` on would not fix it either: the grid's fast path settles the whole
   > grid on one probe, so a per-folder approve inside that loop never fires on a settled site.
   >
   > **Therefore do NOT tighten Draft Item Security to approver-only.** Under that setting a PIC
   > would see their unit folder as empty and could not browse to their own pending files.
   > Closing it needs an end-of-run library sweep approving every pending FOLDER — page by
   > `ID gt <last>`, since `ID` is always indexed and cannot trip the 5,000-item list-view
   > threshold the way a filter on `FSObjType` / `OData__ModerationStatus` can. **Not built:**
   > per-uploader isolation, the only requirement that ever wanted approver-only, is out of
   > scope (`dms-per-uploader-isolation-rejected`). Build it only if that decision reverses.
   >
   > Reading the run log: on a **re-run** the approval lines are absent because there is nothing
   > left to approve. Silence here means settled, not skipped — confirm with the query above
   > rather than by counting `↳ approved` lines.
6. **Upload → approve** end to end. **UPLOAD HALF DONE 2026-08-07** — verified by a real
   unprivileged PIC (member of `GHO_GF_CORU_UPL` only), which proved the whole chain: site
   entry, ancestor-browse corridor, folder ACL, group→role resolution, term ancestry, form.
   > This is what caught the `collectMembership` bug (`2d73c47`): `Role` is a Choice column
   > holding LONG FORM values, so live rows read `UPLOADER` while the form compared against
   > `UPL` and refused a correctly provisioned uploader. Reconciliation had already normalised
   > and granted the folder, so SharePoint looked right and only the form disagreed.
   >
   > **Test as a NON-ADMIN or the test is worthless.** `Form.tsx` gates the error on
   > `!privileged && validPaths.length === 0`, so a site collection admin — or an M365 group
   > owner, who is an SCA via the `…_o` claim without appearing under their own name — always
   > sees a working form whether the code is right or not. One tester was an explicitly-added
   > SCA and would have produced a false pass.
   >
   > Expected UI for a single-path user: Department and Unit locked with a breadcrumb rather
   > than dropdowns. That is correct, not a failure — a two-unit user gets live dropdowns.
7. **Auto-route flow**, in this order: path split `Staging/` → `ApprovalDocument/`; then stamp
   `Author` + `Created` on the Documents copy via `validateUpdateListItem`; then verify the copy
   and delete the source. The path split alone is **already broken** by the rename.
8. **Re-add page access** — now on its own **Page Access** page, not a tab. Audited 2026-08-07:
   **zero Page-scope rows exist**, so all ten Site Pages run on inherited site permissions and
   every admin page is reachable by anyone who can open the site. Five pages need rows:
   `Folder-Access.aspx`, `Site-Access.aspx`, `Approval-Library-Access.aspx`, `Page-Access.aspx`
   and the Folder Manager page.
   > `Target` matches on **`FileLeafRef`**, the page's FILE name — which SharePoint fixes at
   > creation and does NOT change when the page title is edited (same trap as a list's URL).
   > A typo'd Target writes a row reconciliation re-asserts forever against nothing, silently.
   > `Site-Acess.aspx` was created misspelled and renamed to `Site-Access.aspx` before any row
   > referenced it.
9. **Delete `DMS Folder`** once reconciliation has re-stamped `Documents`.
10. ~~Rename `DMS_SITE_MEMBERS` → `CRS_SITE_MEMBERS`~~ — **DONE 2026-08-07, with all 17 groups.**
    Every `DMS_`-prefixed site group was renamed prefix-less (`DMS_GHO_GF_CORU_UPL` →
    `GHO_GF_CORU_UPL`); `DMS_SITE_MEMBERS` alone kept a prefix and became `CRS_SITE_MEMBERS`,
    because that group is resolved BY NAME from a fixed candidate list — strip its prefix and
    the site-entry pass creates a second, empty one and starts feeding it members while the
    real access sits in the original. Ran via `rename-crs-groups.js` (Downloads): 17 renamed,
    0 failed, 0 skipped. Verified after: zero `DMS_` groups remain and the site-entry group is
    still **id 40** — renamed, never recreated, so its membership and every folder grant
    survived.
    > The code never needed this: no production path matches on a prefix. `roleFromGroupName`
    > reads the SUFFIX, the Group Map joins on the integer `GroupId`, and `spGroupsFilter`
    > excludes by id. `matchesAnyGroupPrefix()` and `groupPrefix()` in `naming.ts` DO require
    > one, but have **zero production callers** — dead code, and a hazard: wiring
    > `matchesAnyGroupPrefix` back into a group search would silently drop every prefix-less
    > group. Delete them or leave them, but do not use them.
11. **Update `GroupName` on the CRS Group Map rows** to the renamed titles (the CORU row →
    `GHO_GF_CORU_UPL`). `GroupId` is the join so access is unaffected, but reconciliation
    compares the stored name against the site-entry title to decide NOT to grant that group
    ancestor-browse Read — a stale name defeats that check.

---

## 4. Open with the client

- **Per-uploader isolation in Staging.** Achievable as: Draft Item Security = *"Only users who
  can approve items (and the author)"* — which hides pending **and rejected** files from other
  PICs while an approver sees all — **plus** the Auto-route flow **moving** approved files out of
  Staging, since approved items are visible to everyone with Read. Both halves are needed; the
  draft setting alone leaves approved files visible.
  - Requires **every** folder to be Approved — including the Year × Document Type grid, which
    reconciliation does **not** approve and uploaders **cannot** (see the ⚠ under §3 item 5).
    Without that sweep the corridor goes invisible again (`dms-content-approval-blocks-uploader`).
  - Requires `CRS Approve` to contain `ApproveItems` — see §3.1.
- **Documents library isolation is not replicable.** Draft security needs an unapproved state,
  and Documents is the approved archive. `Created By` there is the flow's account, not the
  uploader (hence the stamping in §3.7). The honest question for the client: *once a document is
  approved and filed, should a colleague in the same unit read it?* If no, those people belong in
  different units — the unit folder is the smallest boundary SharePoint offers
  (`dms-per-uploader-isolation-rejected`).

---

## 5. Known-stale things a new session should not trust

- **`uploadCommand` extension is dead code.** Its id is absent from `componentIds`, and
  `UPLOAD_PAGE` holds the retired sandbox's absolute path, which `onExecute` concatenates onto the
  current web. Pending the delete/keep decision in `2026-08-04-deletion-request-approval-design.md`
  (D3). Left untouched deliberately.
- **`config/serve.json`** still points several configurations at the retired
  `/sites/SPFX-Sandbox-Testing-Ground`.
- **The two 2026-08-03 group specs are superseded** by
  `2026-08-07-role-model-simplification-design.md`. The visibility spec's five open decisions are
  void.
- **`stagingLibrary` in `CRS Config` is superseded** — it still reads `Staging` and would 404. The
  live title wins unless the row holds some other non-legacy name.
- **`DMS Group Map.csv` / `DMS Term Abbreviation.csv`** in the project root are untracked exports,
  kept as the only offline copy of the 175 hand-authored abbreviation rows.
- **Reconciliation logs saying `Staging/…` are not stale** — that is the logical library key, not
  a path. See §2.1. A log is stale if it says `content type → DMS Folder` or *"Ancestor browse
  Read: report only"*.

---

## 6. In design — configurable folder structure (started 2026-08-06/07)

**Spec:** `docs/superpowers/specs/2026-08-06-configurable-folder-structure-chain-design.md`.
Design agreed, **no code written**. Nothing in `src/` has changed for this.

**The ask.** The client wants to add a folder tier below Unit — e.g. `Function` containing
`Human Resource` — and place it anywhere in the path (before Year, between Year and Document Type,
or after Document Type), then upload into it. They cannot edit JSON and will not have a developer
available, so a config-file answer is not an answer.

**Three pieces, build in order:**

1. **The chain model** (spec above). Unify `Levels` and the hardcoded below-Unit folders into one
   ordered chain per mode, each entry flagged `permissioned` true/false. Absent flag means `true`,
   so the three live mode rows are untouched. **Internal milestone — never ships to the client
   alone.**
2. **Structure Manager UI.** An admin web part that authors the chain — name a tier, pick its term
   set, drag it into position — and creates the required columns. This is the actual client
   deliverable. No spec yet.
3. **Subtree migration.** Move existing content under a newly inserted tier. No spec yet; findings
   so far are recorded in the piece 1 spec under "Piece 3 — findings and outstanding tests".

**The hazard to understand before touching `FolderManager.tsx`:** reconciliation creates a folder
per term *and breaks inheritance* for every chain entry. Unfiltered, the new non-permissioned
entries would produce thousands of ACL'd Year and Document Type folders. The client explicitly
declined any inheritance break below Unit (2026-08-06). The guard is a `permissioned === true`
filter before the walk.

**Auto-route findings from this discussion** (observed in the live flow, 2026-08-07):

- Trigger is `When an item is created or modified` — correct, and must not be changed. Power
  Automate offers no "when approval status changes" trigger for SharePoint.
- `Copy file` uses `If another file is already there: Replace`, so a re-fire to the *same* path
  overwrites harmlessly.
- But the destination is built from the **live folder path**, so a folder move sends the copy to a
  new path and orphans the old one. **Migration must therefore run against both libraries at once,
  with the flow disabled** — see §3.7, which already needs the `Staging/` → `ApprovalDocument/`
  path-split fix.
- **Unverified:** the approval email sits in the same branch after `Copy file` and is probably not
  idempotent, so metadata edits on an approved file may re-send it. Read the branch condition to
  confirm. If real, this is a live bug unrelated to the folder-structure work.

**Blocked on site maintenance:** the folder-move test (does approval status survive a move within
the library?). Procedure and expectation table are in the spec. If status does *not* survive,
migration gets materially more expensive and "new uploads only" should be re-proposed to the client
despite their having rejected it.

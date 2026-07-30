# DMS — Full Test Plan (pre-migration UAT)

**Date:** 2026-07-28
**Scope:** everything built to date — 5 web parts, 2 extensions, 4 shared model modules
**Purpose:** prove the solution works end-to-end on `/sites/ClarenceDMSTesting` *before* the client-site
migration, then re-run the same matrix on the client site as the acceptance gate.
**Companion:** `2026-07-28-client-site-migration-runbook.md`

---

## 0. Current automated state (verified 2026-07-28)

`npx heft test` → **57 passing, 0 failing**, 4 suites:

| Suite | Tests | Covers |
|---|---|---|
| `formModel.test.ts` | 27 | `parseLevels`, `parseReconModes`, `collectMembership`, GUID normalisation |
| `groupMapModel.test.ts` | 21 | group-map row parsing, role-from-suffix, dedupe |
| `spGroupsFilter.test.ts` | 4 | SP group name filtering |
| `pathEncoding.test.ts` | 5 | server-relative path encoding |

Build green with **14 lint warnings** (non-blocking): 4× `no-new-null` in `dmsFolderMap.ts`,
4× `eqeqeq`, `max-lines` on `BulkUpload.tsx` (2430 > 2000 limit), 1× `self-closing-comp`, 1× `no-void`.

### Automated coverage gaps (pure logic — cheap to close)

| Module | Lines | Test file | Gap |
|---|---|---|---|
| `src/shared/dmsFolderMap.ts` | 175 | **none** | UniqueId→path resolution, cache behaviour, null handling |
| `src/shared/spGroups.ts` | 175 | partial (`spGroupsFilter` only) | group-name → segment/role/term parsing |

Everything else is REST-coupled and covered by the manual matrix. There is no E2E harness in this
repo and one is not proposed against a live tenant — §3 is the harness.

---

## 1. Prerequisites

Nothing in §3 can run until all of these hold on the target site:

- [ ] `.sppkg` deployed, app added to site, version matches the build under test
- [ ] Libraries `Staging`, `Documents` exist **under exactly those titles** (see A-08 / B-02)
- [ ] Lists `DMS Config`, `DMS Group Map`, `DMS Folder Map` exist under exactly those titles
- [ ] Term sets exist in the site-collection-local group; GUIDs written into `DMS Config`
- [ ] Staging columns provisioned (§2.1) and Documents parity columns provisioned
- [ ] Staging content approval ON; **Draft Item Security = "Any user who can edit items"**
- [ ] Site-entry Read group (`DMS_SITE_MEMBERS`) exists; Staging inheritance broken at *library* level
- [ ] Per-unit SP groups exist (base / `_UPL` / `_APR`); `DMS Group Map` rows written
- [ ] Auto-route flow imported and ON
- [ ] All 5 test accounts provisioned (§1.1)

### 1.1 Test accounts — all five are required; role isolation is the point of the RBAC model

| Account | Membership | Exercises |
|---|---|---|
| `t-admin` | Site Collection Admin | Onboarding, Folder Manager, Group Map Builder, reconciliation |
| `t-upl-a` | one unit's `_UPL` | upload happy path, own-items-only visibility |
| `t-apr-a` | same unit's `_APR` | approve / reject |
| `t-upl-b` | a **different** unit's `_UPL` | cross-unit isolation (negative tests) |
| `t-reader` | base group + `DMS_SITE_MEMBERS` only | read-only; must not reach Staging |

---

## 2. Contract fixtures — verify first; a failure here invalidates everything downstream

### 2.1 Staging internal names

`GET /_api/web/lists/getbytitle('Staging')/fields?$select=Title,InternalName,TypeAsString&$top=500`

Expect **exactly**:

```
Document_x0020_Type            Text
Year_x002f_Period              Text
DocumentDate                   DateTime
Confidentiality_x0020_Level    Text
Vendor                         Text
_ExtendedDescription           Note
Business_x0020_Segment         Text
BusinessSegmentTid             Text
Department                     Text
DepartmentTid                  Text
Unit                           Text
UnitTid                        Text
```

> Internal names are **frozen at creation and not predictable**. Any mismatch is fixed with
> `DMS Config` `col_*` setting rows, or per-level `labelCol` / `tidCol` in the mode's `Levels`
> JSON — **never** by editing code.

### 2.2 DMS Config completeness

`setting` rows the code actually reads (confirmed in `Form.tsx` `loadSettings`). Every one is
**optional**, and each silently falls back to a hard-coded ClarenceDMSTesting value — that silent
fallback is the central migration trap:

```
termSet_documentType   termSet_yearPeriod   termSet_confidentiality   termSet_vendor
col_documentType   col_yearPeriod   col_documentDate   col_confidentiality   col_vendor
col_businessSegment   col_businessSegmentTid
stagingLibrary   allowedExtensions
```

`mode` rows — selected fields are `Title, ModeLabel, Category, TermSetGuid, StagingFolder, Levels,
SortOrder`. **4 rows expected** (Group HO, Upstream Malaysia HO, Minamas HO, NBPOL HO), each
`Category = BusinessSegment`, each
`Levels = [{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]`.

> The column is **`Category`**, not `Side`. `Side` appears in the older seed runbook and is not
> what the code selects.

**T2.2a** — for each of the 4 modes: `GET /_api/v2.1/termStore/sets/{TermSetGuid}/terms`.
Expect 200 + non-empty `value`. A `404` / `apiNotFound` means the GUID belongs to a different site
collection → **the mode renders an empty dropdown with no visible error.**

**T2.2b** — deliberately set a `termSet_*` and a `col_*` row to a wrong value, reload, and record
whether the form fails *visibly* or silently uses a stale default. Highest-risk area of the migration.

**T2.2c** — 3 of the 4 mode rows did not exist as of 2026-07-28. Confirm all 4 are present.

### 2.3 Term label integrity

**T2.3** — every ampersand in a term label must be the **fullwidth ＆ (U+FF06)**, e.g.
`Group Legal, Risk ＆ Compliance`. Search the term response for plain `&` (U+0026); any hit will
break folder-name matching against `DMS Folder Map`.

**T2.4** — NBPOL / Upstream Malaysia / Minamas currently share GHO's terms via
`isAvailableForTagging`. Confirm whether the client is supplying their own Department/Unit trees
before UAT, since that changes the expected dropdown contents for 3 of the 4 modes.

---

## 3. Functional matrix

Each failure records the **actual HTTP status + response body**. A `400` means a malformed request,
not missing data (gotcha #9) — do not diagnose it as a naming/data problem.

### 3.1 Upload Form (`Form.tsx`, 1599 lines) — run as `t-upl-a`

| ID | Test | Pass criteria |
|---|---|---|
| F-01 | Load | Business Segment + Project tabs render; user's segment auto-selected |
| F-02 | Auto-detect | Dept + Unit pre-filled: SP group → Group Map → `loadTermPath` ancestry walk. No manual picking |
| F-03 | Multi-unit user | Add `t-upl-a` to a second unit's `_UPL` → both offered, neither auto-committed |
| F-04 | No-rights user | `t-reader` opens form → clear "no upload rights" message, not a blank cascade or JS error |
| F-05 | Cascade | Choosing Department repopulates Unit; changing Department clears Unit |
| F-06 | Metadata format | Taxonomy-backed text columns get `Label\|GUID`, **not** legacy `-1;#Label\|GUID` |
| F-07 | Date | `DocumentDate` sent as `M/D/YYYY` via `toSpDate()`; stored value is the intended date, no timezone shift |
| F-08 | Field-error trap | Force a bad value → code checks `HasException` per `validateUpdateListItem` result and surfaces it (HTTP is 200 either way) |
| F-09 | Rename | Custom name applied, extension preserved, illegal chars stripped, blank = original filename |
| F-10 | Replace guard | Duplicate name → modal; **Cancel** leaves original intact, **Overwrite** replaces |
| F-11 | Allow-list | `.pdf .doc .docx .xls .xlsx .ppt .pptx .txt .csv .jpg .jpeg .png` accepted; `.exe`/`.bat` rejected **before** any request fires |
| F-12 | Routing | Lands in `Staging/{Segment}/{Dept}/{Unit}/{Year}/{DocType}/`; Year + Type folders ensure-created on demand and **inherit the Unit ACL** |
| F-13 | Deep-path regression | Longest real chain (>330 chars encoded) uploads OK. Gotcha #9: inline-quoted `getfolderbyserverrelativeurl('…')` returns **400**; the `(@f)?@f='…'` alias form works |
| F-14 | Details | "Details" text lands in `_ExtendedDescription` |
| F-15 | Confidentiality | "Highly Confidential" applies the securing behaviour per the 2026-07-16 spec |
| F-16 | Multi-file | Several files per submit → all uploaded, each tagged, sequential not parallel |
| F-17 | Non-admin upload | No 403. A 403 means Draft Item Security is still "Only users who can approve items" |
| F-18 | Isolation | `t-upl-a` cannot reach `t-upl-b`'s unit folder by direct URL |

### 3.2 Bulk Upload (`BulkUpload.tsx`, 2430 lines) — run as `t-admin`

| ID | Test | Pass criteria |
|---|---|---|
| B-01 | Mode load | Reads `ConfigType eq 'mode'`; all 4 HO segments listed |
| B-02 | Direct to Documents | Files land in `Documents` (hard-coded `DOCUMENTS_LIST_TITLE`, line 46), bypassing Staging approval |
| B-03 | Two-batch | Per the 2026-07-24 design: batch 1 uploads, batch 2 tags. A batch-2 failure leaves batch-1 files **identifiable**, not silently orphaned-and-untagged |
| B-04 | Progress UI | Per the 2026-07-28 progress-and-scroll spec: live counter, no freeze at 50+ files, panel scrolls |
| B-05 | Parity | Documents columns receive the same values as their Staging equivalents (2026-07-24 parity spec) |
| B-06 | Partial failure | Drop the network mid-run → per-file error rows, remaining files still attempted, no unhandled rejection |
| B-07 | Throttling | 200-file run → 429s retried with backoff, run completes |
| B-08 | Non-admin | `t-upl-a` opens Bulk Upload → blocked or scoped to own unit. Decide the intended behaviour, then lock it in |

### 3.3 Approval Document (`ApprovalDocument.tsx`, 519 lines)

| ID | Test | Pass criteria |
|---|---|---|
| A-01 | Metadata render | All fields populate. Regression: `FieldValuesAsText` keys are **double-encoded** — read `Department_x005f_x0020_x005f_Type`, not `Department_x0020_Type` |
| A-02 | Approve (`t-apr-a`) | Moderation → Approved; auto-route fires |
| A-03 | Reject | Status → Rejected, comment persisted and visible to uploader |
| A-04 | Wrong approver | `t-apr-a` on another unit's item → denied |
| A-05 | Uploader view | `t-upl-a` sees own pending item read-only, cannot self-approve |
| A-06 | Existing duplicate | Same `FileLeafRef` already in Documents → the lookup finds it and the UI handles the clash, no silent double-write |
| A-07 | Popup polish | Per the 2026-07-27 spec: layout, close behaviour, no scroll-lock leak |
| A-08 | **Hard-coded titles** | Lines 124/137/176/199/229 hard-code `'Staging'` and `'Documents'` — this web part **ignores** the `stagingLibrary` config row. Verify the client site uses these exact titles |

### 3.4 Folder Manager + reconciliation (`FolderManager.tsx`, 1685 lines) — `t-admin`

| ID | Test | Pass criteria |
|---|---|---|
| FM-01 | Browse | Staging tree renders from `DMS Folder Map`; single-click navigates |
| FM-02 | Dry run | Reports what *would* be created; mutates nothing |
| FM-03 | Live run | Missing unit folders created, inheritance broken, correct 3 groups assigned per unit |
| FM-04 | Idempotency | Immediate re-run reports zero changes |
| FM-05 | Live progress | Counter + ETA advance; auto-continue survives a 429 |
| FM-06 | Split log | Successes vs needs-attention listed separately |
| FM-07 | Interrupt | Navigate away mid-run → no half-permissioned folder left silently behind |
| FM-08 | Rename-proof | Rename a folder in SharePoint → stale badge appears; `DMS Folder Map` UniqueId still resolves the new path |
| FM-09 | Member modal | Lists real SP group members; add/remove writes through |
| FM-10 | Delete group | Removes the SP group **and** its `DMS Group Map` rows |
| FM-11 | Open TODOs | 4 items still pending: mobile responsive, split log, progress persistence across nav, horizontally scrollable folder panel. Confirm status before sign-off |

### 3.5 Group Map Builder (`GroupMapBuilder.tsx`, 1181 lines)

| ID | Test | Pass criteria |
|---|---|---|
| GM-01 | One-shot | Creates group + adds members in a single action |
| GM-02 | Role from suffix | `_UPL`→UPL, `_APR`→APR, bare→MEMBER (2026-07-22 spec) |
| GM-03 | `GroupId` type | Written as the SP group **integer** id as a string (e.g. `"27"`) — matches `collectMembership`. **Not** an Entra Object ID |
| GM-04 | GLOBAL role | A `GLOBAL` row is read-only and **ignored** by upload resolution (`collectMembership` skips it) |
| GM-05 | Duplicate guard | Re-creating an existing group does not duplicate Group Map rows |
| GM-06 | Term validity | Rows reference terms that exist in the target term set |

### 3.6 Onboarding (`Onboarding.tsx`, 858 lines)

| ID | Test | Pass criteria |
|---|---|---|
| O-01 | Folder create | Creates unit folder, breaks inheritance **before** `addroleassignment` |
| O-02 | Group claim | `ensureuser` uses `c:0o.c\|federateddirectoryclaimprovider\|{id}` for M365 groups |
| O-03 | Graph-free | Network tab shows **zero** `graph.microsoft.com` calls; `webApiPermissionRequests` is `[]` |
| O-04 | Rollback | Mid-flight failure leaves a reportable state, not silent partial permissions |

### 3.7 Extensions

| ID | Test | Pass criteria |
|---|---|---|
| X-01 | `hideAppBar` application customizer active on Staging views |
| X-02 | `uploadCommand` command-set button on the Staging command bar opens the form |
| X-03 | **Provisioning gap** — no `elements.xml` exists anywhere in the repo, and `uploadCommand`'s id `c2fd5b59-02ed-4892-87f2-2418fefb4980` is **absent from `package-solution.json` componentIds**. Both extensions therefore need explicit provisioning. See runbook §M-04 |

### 3.8 Auto-route flow (Power Automate)

| ID | Test | Pass criteria |
|---|---|---|
| R-01 | Approve → file copied to the matching `Documents` path |
| R-02 | Routing survives a folder rename (routes by live path, not stored name) |
| R-03 | Splits on `Staging/` (no leading slash); trailing slash stripped — Copy file rejects a trailing slash |
| R-04 | Notification sent from the `DMS No Reply` shared mailbox |
| R-05 | Missing target folder → "Create new folder" branch handles it |
| R-06 | Reject → **no** copy occurs |

---

## 4. Cross-cutting regression checklist

Run once at the end. Every item is a bug that has already bitten this project.

- [ ] No `Promise.allSettled` anywhere (tsconfig doesn't target ES2020) — per-item `try/catch` inside `Promise.all`
- [ ] No `getfolderbyserverrelativeurl('<inline literal>')` remains; all use `(@f)?@f='…'`
- [ ] CAML uses `FSObjType` / `_ModerationStatus`, never the REST names (wrong name throws `-2130575340`)
- [ ] Upload bodies are raw `File`/`Blob`, never `FormData`
- [ ] Staging `Name` formatting: folders use `customRowAction: defaultClick`, files use `<a href>`
- [ ] `npx heft test` green; lint warning count has not grown past 14

## 5. Exit criteria

1. All §2 fixtures verified against the target site.
2. §3 matrix executed with all 5 accounts; zero open Sev-1/Sev-2.
3. §4 checklist clean.
4. `npx heft test` ≥ 57 passing.
5. Every row recorded pass / fail / N-A with evidence, attached to migration sign-off.

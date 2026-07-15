# Multi-Segment Upload Form — Flow & Data Model Design

**Date:** 2026-07-15
**Scope:** The upload form (`Form.tsx`) + its DMS Config data + a new DMS Group Map list.
**Companion spec:** `2026-07-15-staging-per-unit-groups.md` (the permission/isolation build sheet this form sits on top of).

---

## Problem

The form today has a **2-mode toggle** (Department / Projects) with a hardcoded 1–2 level cascade and `detectDepartment()` that matches the user's M365 group **by display-name substring** against the department term set's top-level terms.

The real org model turned out to be **5 top-level term sets** of **differing depth and differing level labels**:

| Top-level term set | Toggle side | Level chain (after the segment) |
|---|---|---|
| Group Head Office | Business Segment | Department → Unit |
| Group Upstream Operations | Business Segment | Region → Estate/Mill |
| Group SDGI Operations | Business Segment | Refinery → Department |
| Group Innovation & Technology | Business Segment | I&T Operating Unit *(1 level)* |
| Group-led Projects | Project | Project Name → Department → Unit |

Common tail on every mode: **Year/Period → Document Date → Document Type → Confidentiality → Vendor → Details**.

Two things break under this model and must change:
1. The cascade must become **data-driven variable depth (1–3 levels)** with a **per-level label**, not hardcoded.
2. Group→term detection can **no longer match by name** (renames, punctuation, no reliable substring rule, and per-unit groups like `DMS_GF_TAX_UPL` don't contain the term text). It must use an **explicit mapping keyed by the group Object ID**.

---

## Decisions locked

| # | Decision | Choice |
|---|---|---|
| 1 | Number of modes | **5** (4 Business Segments + 1 Projects), each its own term set |
| 2 | Toggle | **`Business Segment | Project`** |
| 3 | Auto-detect depth | **Full path, pre-filled but editable** (a user can belong to multiple units) |
| 4 | Group→term link | Hand-filled **DMS Group Map** list, keyed by group **Object ID** (never name) |
| 5 | Cascade | **Data-driven**: each mode declares an ordered list of levels, each with its own label |
| 6 | Metadata columns | **One column per distinct label**, typed as **plain text** (Option A), searchable |

---

## 1. DMS Config — mode rows

Extend the existing `ConfigType eq 'mode'` items. New/changed fields:

| Field | Type | Example (GHO) | Purpose |
|---|---|---|---|
| `Title` / `ModeLabel` | Text | `Group Head Office` | display name |
| `Side` | Choice | `BusinessSegment` \| `Project` | which toggle side it belongs to |
| `TermSetGuid` | Text | *(GHO set GUID)* | the term set to walk |
| `StagingFolder` | Text | `Group Head Office` | top staging folder for path building |
| `Levels` | Multiline (JSON) | *(see below)* | ordered levels: label + target column |
| `SortOrder` | Number | `1` | order within its toggle side |

`LookupStyle` / `SubTeamLabel` from the old schema are **retired** — replaced by the generic `Levels` array.

`Levels` JSON (one entry per dropdown, in order):

```json
[
  { "label": "Department", "column": "Department" },
  { "label": "Unit",       "column": "Unit" }
]
```

Per mode:

| Mode | Levels JSON |
|---|---|
| Group Head Office | `[{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]` |
| Group Upstream Operations | `[{"label":"Region","column":"Region"},{"label":"Estate/Mill","column":"Estate_Mill"}]` |
| Group SDGI Operations | `[{"label":"Refinery","column":"Refinery"},{"label":"Department","column":"Department"}]` |
| Group Innovation & Technology | `[{"label":"I&T Operating Unit","column":"IT_Operating_Unit"}]` |
| Group-led Projects | `[{"label":"Project Name","column":"Project_Name"},{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]` |

---

## 2. DMS Group Map — new SharePoint list

Plain list, **hand-filled** by the admin (rows added as each unit group is created). Read by the form at load.

| Column | Type | Example | Purpose |
|---|---|---|---|
| `GroupId` | Text | `a1b2c3d4-…` (Entra Object ID) | **the stable match key** vs `/me/memberOf` |
| `GroupName` | Text | `DMS_GF_TAX_UPL` | human-readable only, never matched on |
| `Segment` | Text | *(GHO set GUID)* | which term set / mode this group belongs to |
| `UnitTermGuid` | Text | *(Tax term GUID)* | the leaf term; form walks **up** for ancestors |
| `Role` | Choice | `UPL` \| `APR` | optional; reserved for later Documents-library logic |

The form does not care who wrote the rows — schema + reader are fixed regardless.

---

## 3. Form flow

**On load:**
1. Load modes from DMS Config; group them by `Side` for the toggle.
2. Graph `GET /me/memberOf` → collect the user's **group Object IDs**.
3. Load DMS Group Map → keep rows whose `GroupId` is in the user's set. Each match = one `{ segment, unitTermGuid }` path.
4. Pick a default: first matched Business-Segment path → set active mode + pre-select its cascade by walking **up** the term tree from `unitTermGuid` to fill each level. Dropdowns stay **editable**.
5. If the user has Project-side matches, they surface when the Project toggle is selected.

**Rendering:**
- Toggle `Business Segment | Project`.
- For the active mode, render **one labelled dropdown per entry in `Levels`**, each populated by walking the term tree (parent → children) as the level above is chosen.
- Then the common tail: Year/Period, Document Date, Document Type, Confidentiality, Vendor, Details.

**On upload:**
- Resolve the destination folder by **UniqueId** (rename-proof — via the existing DMS Folder Map / reconciliation output).
- Upload the raw `File`/`Blob` (never FormData).
- Tag columns: for each level, write its `column` ← selected term **label** (text) **and** `<column>_Tid` ← the term **GUID**; `BusinessSegment` / `BusinessSegment_Tid` ← active segment label + GUID; plus the tail columns. Columns not used by the active mode stay blank.

---

## 4. Metadata columns (Staging library)

Each level stores **two** plain-text columns: a **visible label** column (searchable, clean in views) and a **companion GUID** column `<column>_Tid` (hidden, for rename-safe re-matching). Columns are shared where labels repeat.

| Label column (internal) | GUID column | Used by |
|---|---|---|
| `BusinessSegment` | `BusinessSegment_Tid` | all business-segment modes (segment label) |
| `Department` | `Department_Tid` | GHO, SDGI, Projects |
| `Unit` | `Unit_Tid` | GHO, Projects |
| `Region` | `Region_Tid` | Upstream |
| `Estate_Mill` | `Estate_Mill_Tid` | Upstream |
| `Refinery` | `Refinery_Tid` | SDGI |
| `IT_Operating_Unit` | `IT_Operating_Unit_Tid` | I&T |
| `Project_Name` | `Project_Name_Tid` | Projects |

16 columns total (8 label + 8 GUID). Existing tail columns unchanged: `Year_x002f_Period`, `DocumentDate`, `Department_x0020_Type` (Document Type), `Confidentiality_x0020_Level`, `Vendor`, `_ExtendedDescription` (Details).

**Why text, not taxonomy:** a managed-metadata column binds to exactly one term set, so a *shared* `Department` column can't accept GHO **and** SDGI **and** Projects departments. Correctness is already enforced by (a) the folder path the file lands in and (b) the form's dropdowns being built from the term store. The label column is a searchable record of the choice; the `_Tid` column is the stable identity. Text columns are crawled by search, filterable, sortable, and viewable — everything needed here except the MMD refiner panel, which the folder tree already replaces.

**Rename safety:** if a term is later renamed, the label column keeps the old word but the `_Tid` GUID column still matches the term (and the folder/UniqueId routing stays correct regardless). The `_Tid` columns exist precisely so search/re-tagging survives renames.

---

## 5. Code touch-points

| File | Change |
|---|---|
| `src/webparts/form/components/Form.tsx` | Mode loader reads `Side` + `Levels` (JSON); **replace `detectDepartment()`** with a DMS Group Map reader keyed by Object ID; generic **N-level cascade renderer** driven by `Levels`; payload builder writes per-level text columns |
| `src/webparts/reconciliation/components/Reconciliation.tsx` | Already path-based; add the 5 modes to `DEFAULT_MODES` (fallback) |
| DMS Config list (data) | Add/replace 5 mode rows with `Side` + `Levels`; retire `LookupStyle`/`SubTeamLabel` |
| DMS Group Map list (data) | New list, hand-filled |
| Staging library (data) | Create 16 text columns above (8 label + 8 `_Tid`) |
| `CLAUDE.md` + code fallbacks | Update term-set GUIDs, mode defaults, column names |

---

## 6. Migration order (one-by-one, per user's cadence)

1. ✅ Term Store restructured into 5 sets *(done)*.
2. Rebind / retire the old `Department` taxonomy column binding (old dept set `eaba82e5-…`).
3. Create the 16 plain-text columns on Staging (8 label + 8 `_Tid` GUID).
4. Capture the 5 new term set GUIDs.
5. Update DMS Config: 5 mode rows (`Side`, `TermSetGuid`, `StagingFolder`, `Levels`).
6. Create the DMS Group Map list; fill rows as unit groups are created.
7. Update code: cascade renderer + group-map reader + fallbacks; rebuild/ship.
8. Reconciliation run → term→folder UniqueId map for path routing.

---

## Open items / deferred

- **Documents library** columns + reader groups — later pass (mirror this model).
- **Home page** entry point / navigation — later.
- Exact **internal names** of the 16 new columns to confirm against the live `/fields` API after creation (SharePoint may freeze encoded names — e.g. `_Tid` or `/` handling).

# Segment Onboarding Plan

**Date:** 2026-07-21
**Goal:** Make onboarding a business segment (Upstream, SDGI, I&T, Group‑led Projects) a repeatable, mostly **data‑only** process, and lock down the DMS Config `mode` row schema that drives both the upload form and (after one small code change) the reconciliation provisioner.

Related: `2026-07-15-multi-segment-form-flow-design.md`, `2026-07-16-tenant-seed-data-runbook.md`.

---

## 1. Current state (verified in code)

| Piece | Source of modes | Status |
|---|---|---|
| **Upload form** (`Form.tsx` → `loadModes()`) | DMS Config rows where `ConfigType = 'mode'`; falls back to hardcoded `DEFAULT_MODES` (GHO) when no usable row | **Data‑driven** ✅ |
| **Reconciliation** (`FolderManager.tsx` → `RECON_MODES`) | **Hardcoded** — GHO only | **NOT data‑driven** ⚠️ |
| **Metadata settings** (`termSet_documentType/yearPeriod/confidentiality/vendor`) | DMS Config rows where `ConfigType = 'setting'` | **Live** ✅ |

- The form ignores mode rows with an **empty `Levels`** (`Form.tsx` line ~549), so the retired `department`/`project` rows contribute nothing and the form runs on `DEFAULT_MODES` (GHO).
- **Stale rows to delete** (cleanup, in progress by Clarence): `department`, `project`, `termSet_department`.

---

## 2. DMS Config `mode` row schema (the data contract)

`loadModes()` selects: `Title, ModeLabel, Side, TermSetGuid, StagingFolder, Levels, SortOrder` filtered by `ConfigType eq 'mode'`.

| Column | Type | Required | Example (GHO) | Notes |
|---|---|---|---|---|
| `Title` | Single line | yes | `mode_gho` | Internal key; any unique value |
| `ConfigType` | Single line / Choice | yes | `mode` | **Must be exactly `mode`** |
| `ModeLabel` | Single line | yes | `Group Head Office` | Shown in the form's "Upload into" selector |
| `Side` | Choice / Single line | yes | `BusinessSegment` | `BusinessSegment` or `Project` |
| `TermSetGuid` | Single line | yes | `efa87c6a-9536-4f7c-910f-011bf7413b80` | The segment's term‑set GUID. **No trailing spaces.** |
| `StagingFolder` | Single line | yes | `Group Head Office` | Top‑level folder name under Staging (and Documents) |
| `Levels` | **Multiple lines (plain)** | yes | *(JSON — see below)* | Empty = ignored by the form |
| `SortOrder` | Number | yes | `1` | Display order 1..5 |

### `Levels` JSON

Array of level objects, ordered top → bottom (the cascade the user picks):

```json
[
  { "label": "Department", "column": "Department", "labelCol": "Department", "tidCol": "DepartmentTid" },
  { "label": "Unit", "column": "Unit", "labelCol": "Unit", "tidCol": "UnitTid" }
]
```

| Field | Required | Meaning |
|---|---|---|
| `label` | yes | Dropdown label shown to the user (e.g. "Estate/Mill") |
| `column` | yes | Logical key; resolves to Staging columns via `LEVEL_COLUMNS` **only as a fallback** |
| `labelCol` | **recommended** | Real Staging internal name of the text label column |
| `tidCol` | **recommended** | Real Staging internal name of the term‑GUID column |

> **Always set `labelCol`/`tidCol`.** They win over `LEVEL_COLUMNS`, so a new segment needs **no `Form.tsx` code edit** — the mapping lives entirely in the row. (SharePoint does **not** follow a `<label>_Tid` convention — the Tid column has its own frozen internal name; read it from `/fields`.)

---

## 3. The five mode rows

GHO is the only one with a real GUID today. The other four carry `REPLACE-*` placeholders in `DEFAULT_MODES` — capture the real term‑set GUIDs during onboarding.

| SortOrder | ModeLabel | Side | StagingFolder | Levels (label → column) | TermSetGuid |
|---|---|---|---|---|---|
| 1 | Group Head Office | BusinessSegment | Group Head Office | Department → Department · Unit → Unit | `efa87c6a-9536-4f7c-910f-011bf7413b80` ✅ |
| 2 | Group Upstream Operations | BusinessSegment | Group Upstream Operations | Region → Region · Estate/Mill → EstateMill | **TBD** |
| 3 | Group SDGI Operations | BusinessSegment | Group SDGI Operations | Refinery → Refinery · Department → Department | **TBD** |
| 4 | Group Innovation & Technology | BusinessSegment | Group Innovation & Technology | I&T Operating Unit → ITOperatingUnit | **TBD** |
| 5 | Group‑led Projects | Project | Group‑led Projects | Project Name → ProjectName · Department → Department · Unit → Unit | **TBD** |

> ⚠️ The Projects term‑set GUID `e1163337-93bb-4b6e-847e-346f51cc6806` from the handover PDF **does not exist in the tenant — do NOT use it.** Capture the real one.

---

## 4. Per‑segment onboarding checklist

Do these in order for each new segment:

- [ ] **1. Term store** — create the segment's term set with its `Segment → level1 → level2 …` hierarchy (variable depth). **Record the term‑set GUID.**
- [ ] **2. Staging columns** — for each **new** level, create a **label** text column and a **Tid** text column (see §5). Reuse `Department`/`Unit` where a segment already uses them. Read back the frozen internal names via `/fields`.
- [ ] **3. DMS Config `mode` row** — add one row per §2, with the **real** `TermSetGuid` and a `Levels` JSON whose `labelCol`/`tidCol` point at the Staging columns from step 2.
- [ ] **4. Security groups** — create per‑unit groups (UPL / APR / MEMBER) plus segment‑ and department‑tier reader groups, following the naming convention.
- [ ] **5. DMS Group Map rows** — one row per group: `GroupId`, `SegmentTermGuid`, `UnitTermGuid` (= the term of the tier the group governs), `Role` (`MEMBER`/`UPL`/`APR`). **No trailing spaces in any field** (a trailing space silently drops the row — this bit us on `DMS_GHO_LRC`).
- [ ] **6. Run Folder Reconciliation** — creates + locks folders at every tier in Staging + Documents, auto‑assigns the groups, writes the Staging folder map, and builds the Year × Document Type grid. Verify the log shows `↳ <group> → <perm>` lines (no `⚠ no group‑map groups` on tiers that should have a group).

---

## 5. Level columns to create — in BOTH Staging AND Documents

Add these **label + Tid** text columns per segment (skip ones already present). Keep the internal names for the `Levels` JSON `labelCol`/`tidCol`.

> **Create each column in the Documents library too, with the identical internal name and type.**
> Staging and Documents must mirror (see `2026-07-24-documents-metadata-parity-design.md`). A
> level column that exists only in Staging reproduces the blank-metadata bug on the new segment:
> the Auto-route flow (and the bulk-upload tool) write it, but Documents has nowhere to put it.
> The head-office pilot columns (`Business Segment`/`Department`/`Unit` + Tids) are handled by that
> parity fix; every **new** segment's columns below must be added to both libraries at onboarding.

| Segment | New columns (label / Tid) | Reuses |
|---|---|---|
| Upstream | `Region`/`RegionTid`, `EstateMill`/`EstateMillTid` | — |
| SDGI | `Refinery`/`RefineryTid` | `Department`/`DepartmentTid` |
| I&T | `ITOperatingUnit`/`ITOperatingUnitTid` | — |
| Projects | `ProjectName`/`ProjectNameTid` | `Department`/`DepartmentTid`, `Unit`/`UnitTid` |

> Plain‑text columns (label + its term‑GUID), matching the existing `Business Segment` / `Department` / `Unit` pattern. Because `labelCol`/`tidCol` are supplied in the mode row, **`Form.tsx` `LEVEL_COLUMNS` does not need editing** — though keeping it in sync is a nice belt‑and‑suspenders.

---

## 6. Required code change — make reconciliation data‑driven (recommended)

**Problem:** reconciliation reads `RECON_MODES` (hardcoded GHO). Adding a segment would need a code edit + rebuild + redeploy every time — exactly the manual friction we're trying to remove.

**Change:** have reconciliation load its modes from the **same DMS Config `mode` rows** the form uses, so onboarding a segment is **data‑only**.

- **Extract a shared loader** — move the mode‑loading logic (or a slim version needing just `termSetGuid` + `stagingFolder`) into `src/shared` so both `Form.tsx` and `FolderManager.tsx` call one function. Keep it consistent with `loadModes()` (`ConfigType='mode'`, ignore empty `Levels`).
- **FolderManager** — replace the `RECON_MODES` loop in `buildProvisionTargets` with the loaded modes (`mode.termSetGuid`, `mode.stagingFolder`). Keep a **hardcoded GHO fallback** if the config read fails or returns nothing (mirrors the form's `DEFAULT_MODES` safety net).
- **Result:** onboarding a segment = add a DMS Config `mode` row + groups + map rows, then **Run reconciliation**. No redeploy.

**Alternative (not recommended):** keep `RECON_MODES` hardcoded and add four entries — a rebuild/redeploy per segment, and two places (form + reconciliation) to keep in sync.

This code change should be its own small implementation plan (TDD on the shared loader), tracked separately from the data steps above.

---

## 7. Data to capture from the client (blocking)

Before any of the four remaining segments can be seeded:

| Segment | Needed |
|---|---|
| Upstream / SDGI / I&T / Projects | Real **term‑set GUID**, confirmed **level structure** (labels + depth), and the **Staging folder name** |
| Projects | The **real** Project term‑set GUID (the handover PDF's is bogus) |
| All | The list of **units** per segment and which **security groups** govern each tier + role |

---

## 8. Open decisions

1. **Data‑driven reconciliation** (§6) — recommended; confirm before building.
2. **Group Map Builder tool** (raised separately) — dropdown‑driven row creation to kill the manual GUID/typo/trailing‑space class of bugs. Reuses `searchGroups` + term‑store cascade already in Folder Manager. Spec separately.
3. **Cleanup** (Clarence, in progress) — delete stale DMS Config rows `department`, `project`, `termSet_department`.

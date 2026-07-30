# Folder Abbreviation Naming — Design

**Date:** 2026-07-30
**Status:** Agreed, not yet implemented
**Touches:** new `DMS Term Abbreviation` list, `FolderManager.tsx`, `DMS Config` (`StagingFolder`),
Staging + Documents libraries (new `Full Name` column), message text in `Form.tsx` / `BulkUpload.tsx`

---

## Problem

Folder names come from term labels, and the labels are long. Worst case measured across the client's
133 units:

| | Raw | URL-encoded |
|---|---|---|
| Longest chain today | 205 | **257** |
| Median | — | 185 |
| Abbreviated equivalent | — | **115** |

Encoding is what hurts. The term store requires **fullwidth ＆**, which costs **9 encoded characters**
(`%EF%BC%86`), and every space costs 3. The worst unit — `Org Culture ＆ Workplace Services - Org
Culture ＆ Transformation (OCT)` — spends 48 characters on twelve source characters.

`GetFolderByServerRelativeUrl` starts returning **HTTP 400** at roughly 330 characters (gotcha #9), so
the current worst case leaves about **73 characters of headroom** — one long filename from failure.
Abbreviating to `GHO/GCA/GMB_STRATCOMMS` cuts ~142 characters and takes headroom to ~215.

The client also wants the full name still visible when they click a folder.

## What the client asked for, and what is possible

> Term store keeps `Group Head Office`, the upload dropdown shows `Group Head Office`, the folder is
> named `GHO`, and the details panel shows Name = `Group Head Office`.

The first three are achievable. **The fourth is not:** the details pane's `Name` field *is* the folder
name — the same string that forms the path segment. It cannot be short in the URL and long in `Name`.

Delivered instead: a `Full Name` column on the folder, so the details pane reads

```
Name       GHO
Full Name  Group Head Office
```

Hidden from the **view** (not from the form) so it stays out of the grid but renders in the details
pane.

## Design

### 1. Where abbreviations live

| Level | Source | Why |
|---|---|---|
| Segment | `StagingFolder` on the `mode` row in `DMS Config` | The segment container has **no term** (`FolderManager.tsx:1041` pushes `termGuid: null`), so it cannot be keyed by term GUID. Already config-driven. |
| Department, Unit | New `DMS Term Abbreviation` list, keyed by term GUID | Matches the established shape — `DMS Group Map` keys on `UnitTermGuid`, `DMS Folder Map` on term GUID. |

`DMS Term Abbreviation` columns: `TermGuid` (Text, the key), `Abbreviation` (Text), `FullName` (Text,
for reference), `Level` (Choice: Department | Unit).

Term **labels stay full**. The abbreviation is a separate mapping, never a rename of the term. The
upload dropdown continues to read `labels[0].name` and is unaffected.

### 2. Blast radius is one web part

Only `FolderManager.tsx` derives folder names from labels (`:1041`, `:1054-1070`). Confirmed:

- `Form.tsx` resolves the unit folder by **UniqueId** — `lookupFolderMapping` (`:942`) then
  `resolveFolderServerUrl` (`:962`) — and only appends `Year` / `Document Type`. It never rebuilds the
  segment/department/unit path.
- `BulkUpload.tsx` performs string surgery on the **live** path to swap `Staging/` for the Documents
  library (`:1281`, `:1377`). Name-agnostic.

So uploads keep working across a rename with no code change. This is the rename-proof-by-UniqueId
design doing its job.

### 3. Naming rule

```
/<StagingFolder>/<department abbreviation>/<unit abbreviation>
```

`sanitizeFolderSegment` still applies. A unit abbreviation **must carry its sub-group prefix** — the
client's "Unit" labels are really `SubGroup - Unit`, and dropping the prefix collides (see §4).

### 4. The collision this creates

Today collisions are impossible: names come from term labels, and the term store forbids two siblings
sharing a name. Verified against the client's data — zero duplicate unit names within any department.

**Abbreviations remove that guarantee**, because a hand-maintained list has no such constraint. One
real case exists in the current data:

```
Group Head Office | Group Corporate Affairs
  Global Marketing ＆ Branding - Strategic Communications
  Group Communications        - Strategic Communications
```

Same department, same trailing words. Abbreviate on the post-dash part and both become
`STRATCOMMS` — **one folder, one ACL, two units' documents**, which breaks the isolation the whole
permission model rests on. The client's own tokens already avoid it
(`GHO_GCA_GMB_STRATCOMMS` vs `GHO_GCA_GC_STRATCOMMS`).

This is a risk the change *introduces*, not one being discovered. The safety net ships with it.

### 5. Safety net

**Uniqueness validated at entry.** The admin UI rejects an abbreviation that collides with a sibling
under the same parent. First line of defence, before anything reaches the filesystem.

**Rename on mismatch, by UniqueId.** For each mapped term: `GetFolderById(stored UniqueId)` → compare
the folder's current leaf name to the expected abbreviation → rename if different. Reconciliation
already verifies the stored UniqueId rather than the label path (`:1264`), so this extends existing
behaviour from *tolerating* a mismatch to *correcting* it. Uploads continue working before, during,
and after, because the map row never changes.

**Rename conflict is reported, never skipped or forced.** If a folder already occupies the target
name the rename fails — e.g. the map was changed from `COR` to `CORS`, an earlier run created an
empty `CORS`, and the real `COR` now cannot take that name. Silently skipping leaves a real folder
and an empty decoy side by side with nobody told; forcing risks merging two different ACLs. Report
`cannot rename COR -> CORS: a folder named CORS already exists here` and stop touching that branch.
This also catches any collision that slipped past the entry validation — last line of defence.

**Missing abbreviation skips and reports.** No abbreviation means no folder, which means no
`DMS Folder Map` row, which means uploads to that unit fail. That is intended, but it must be loud:
the term appears in a "needs attention" list naming the cause.

**Orphans reported, never auto-deleted.** A row whose term no longer exists is flagged, and removed
only on explicit confirm. Reconciliation must **not** infer "the term is gone" from a term-store read
that returned nothing — a transient failure would then wipe the mapping list. Delete only when the
read demonstrably succeeded and returned a non-empty set. Note the folder and its documents survive
regardless; flag those too.

### 6. `Full Name` column

Text column on Staging and Documents, written by reconciliation at folder creation with the term's
full label (segment folders get the `ModeLabel`). Hidden from the default **view**; left visible on
the **form**, which is what the details pane renders.

### 7. Admin UI

**Bulk fill** for onboarding: the first run has **178 terms** with no abbreviation (3 segments +
42 departments + 133 units). A per-term prompt would be unusable at that volume — list every missing
term in a grid, fill, submit once, re-run.

**Single prompt** for steady state: one new term added later gets a modal with the warning and a
one-row form. Precedent exists — `GroupMapBuilder` already resolves terms and writes to a list.

### 8. Messages to correct

- **Upload failure** currently tells the user to re-run reconciliation. Re-running does not help when
  the cause is a missing abbreviation. Name the real cause.
- **Reconciliation report** must distinguish: created / renamed / skipped-no-abbreviation /
  rename-conflict / orphan.

## Client-facing guidance

A short written guideline covering: abbreviations are permanent-ish (changing one renames a live
folder); they must be unique within their parent; a unit abbreviation must include its sub-group;
a missing abbreviation blocks that unit's uploads entirely; and the full name remains visible in the
details panel.

## Abbreviation data — RESOLVED 2026-07-31

An earlier draft of this spec claimed the client's combined tokens could not be split mechanically
and that only 133 of the 178 abbreviations existed. **Both were wrong.** Splitting as
*segment / department / everything remaining* succeeds on every row — the extra underscore in
`GHO_GHR_OCWS_IFM` sits inside the unit token (`OCWS_IFM`), it is not a fourth level. So the client
had already supplied all three levels, encoded in one column.

All 178 are generated and validated in **`docs/term-store-import/10-per-level-abbreviations.csv`**
(`Level`, `BusinessSegment`, `Department`, `Unit`, `FullName`, `Abbreviation`, `Source`, `Notes`).
This is a starting set to be corrected by the client, not a final answer; the `Source` and `Notes`
columns record every deviation from their file.

### Conflicts corrected in the client's data

Two department tokens were used for **two different departments** — precisely the failure in §4,
which would have merged distinct departments into one folder and therefore one ACL:

| Token | Was used for | Resolution |
|---|---|---|
| `CEOOA` | CEO Office Administration **and** Corporate Communication | Corporate Communication → `CORPCOMM` |
| `SC` | Supply Chain **and** Sustainability | Sustainability → `SUS` |

Three departments were internally inconsistent (`GCA`/`CGA`, `PNE`/`PROD`, `GCA`/`PMOPS`), and 15
rows carried no department token — mostly single-unit departments whose department and unit share a
name, plus four that simply omitted it (`INEXCOMMS`, `PMOPS`, `SCRA`, `UPSUPPORT`).

### Measured result

| | Before | After |
|---|---|---|
| Worst-case encoded path | 257 | **129** |
| Median | 185 | **115** |
| Headroom to the ~330 limit | 73 | **201** |

Validated: no duplicate abbreviation among siblings at any level, no blanks.

### Still outstanding

The abbreviation list is keyed by **term GUID**, and the CSV has none — it is keyed by label path.
GUID resolution happens at import time by walking the term store and matching labels, which the
admin UI in §7 must do.

## Out of scope

- Restructuring the term store into four tiers to make `SubGroup` a real level. The sub-group is
  carried inside the unit abbreviation instead.
- Flattening the tree to one folder per unit. Rejected — the nested chain is assumed by
  `DMS Folder Map`, reconciliation, and leaf-only authorization.
- Migrating an existing populated tree. The project is pre-production with no real documents; the
  tree can be deleted and rebuilt. **This must be revisited before anything ships to `sdguthrie`.**

## Correction to an earlier spec

`2026-07-29-folder-map-integrity-design.md` is still marked "Agreed, not yet implemented", but
`FolderManager.tsx:1264` implements the UniqueId verification it describes. That status line should
be updated.

# Twelve-Segment Expansion — Head Office Pilot

**Date:** 2026-07-24
**Status:** approved design — 4 head-office term sets exist (GUIDs below, captured
2026-07-24) and are in the code fallbacks; DMS Config `mode` rows + client Department→Unit
trees still pending
**Amends:** `2026-07-21-segment-onboarding-plan.md` — replaces its §3 "five mode rows" table.
The onboarding *process* (§2 schema, §4 checklist, §5 columns, §6 data-driven reconciliation)
is unchanged and still governs how each segment below gets onboarded.

## What changed

The client's structure has more head offices than assumed. "Group Head Office" is not the
only head office — there are four, and the full 2027 Business Segment list has **12 values**,
not 5. Each is its own **mode** (own term set, own DMS Config row) — the existing data-driven
architecture; no umbrella "Head Office" level is added.

**2026 pilot scope: the 4 head-office segments only.** The rest onboard in 2027.

## The 12 segments

| # | Business Segment | Family | Levels after segment | Pilot? |
|---|---|---|---|---|
| 1 | Group Head Office | Head Office | Department → Unit | ✅ live (`efa87c6a-9536-4f7c-910f-011bf7413b80`) |
| 2 | Upstream Malaysia Head Office | Head Office | Department → Unit | ✅ 2026 (`5ab1c7c4-78d2-43b4-869f-3eab4b1c375c`) |
| 3 | Minamas Head Office | Head Office | Department → Unit | ✅ 2026 (`6ba9a64c-a363-48fd-afd1-324897df781c`) |
| 4 | NBPOL Head Office | Head Office | Department → Unit | ✅ 2026 (`21d7e6fe-8f71-4a56-bd2e-e4a2176995a7`) |
| 5 | Upstream Malaysia Operations | Upstream Ops | Region → Estate/Mill | 2027 |
| 6 | Minamas Operations | Upstream Ops | Region → Estate/Mill | 2027 |
| 7 | NBPOL Upstream Operations | Upstream Ops | Region → Estate/Mill | 2027 |
| 8 | SDGI Malaysia | SDGI | Refinery → Department | 2027 |
| 9 | SDGI Overseas | SDGI | Refinery → Department | 2027 |
| 10 | Innovation & Technology Malaysia | I&T | I&T Operating Units/Department → Unit | 2027 |
| 11 | Innovation & Technology Overseas | I&T | I&T Operating Units/Department → Unit | 2027 |
| 12 | Group-led Project | Projects | Department → Unit (top folder is the **Project Name**) | 2027 |
| 13 | Group Business Ventures & Transformation | Head Office | Department → Unit | 2027 |

> **Revised 2026-08-17 against the client's own folder-structure table, then against their spoken
> confirmation.** Three rows changed and one was added:
> - **Group-led Project has two levels, not three.** `Project Name` is the segment's **top folder
>   name** — the slot `Business Segment` occupies for every other family — not a permissioned tier. This
>   row read as three only because the top folder was listed inside the chain.
> - **Segment 13 is new**, same shape as a head office. It needs its own term set and is absent from
>   every code fallback (`DEFAULT_MODES`, `RECON_MODES`) — inert until a `mode` row exists, so this is a
>   note rather than a defect.
> - **I&T is TWO tiers, and "one tier" meant one tier NAME.** This row was recorded as a single level in
>   July, flipped to one again on 2026-08-17, and settled at two the same day when the client clarified:
>   *"what I meant by one tier is the I&T Operating Units/Department not the entire thing"*.
>   `I&T Operating Units/Department` is **one tier name containing a slash**, with `Unit` below it. The
>   slash reads as a tier boundary and is not one — a `/` inside a level name is the one punctuation mark
>   in this system that means nothing structural, which is exactly why it misleads.
>
> **THE RECURRING MISREADING, ACROSS EVERY ONE OF THESE ROWS, IS COUNTING THE BUSINESS SEGMENT AS A
> TIER.** It is the Top folder name, above the chain. It is worth being blunt about the cost, because it
> decides the direction to err in: the depth check demands the term set's depth equal the tier count, and
> reconciliation walks the term tree capped at that count — so **one tier too MANY moves every unit's ACL
> a level down onto a folder no Group Map row points at, grants successfully, and looks correct.** One
> too FEW is loud: the form refuses and writes nothing.
>
> **ALL THIRTEEN FAMILIES ARE EXACTLY TWO TIERS, and `validateNewSegment` now REFUSES fewer than two**
> (client's instruction 2026-08-17: *"I think best to force them to create two not one. atleast two"*).
> That floor is not redundant with the depth check, which is why it is worth having: one declared tier
> against a **one-deep** set passes the depth check cleanly, and a set is one-deep precisely when someone
> has authored the departments and not yet the units. The segment then saves with the **department**
> holding the access, so every unit under it shares one folder and one ACL. It is a refusal rather than a
> warning because the depth check's own get-out is warn-and-allow when depth is unknown, and a second
> soft signal would leave the one shape that reads as complete advisory in both places.
>
> `I&T Operating Units/Department` keeps its slash in the **label**, but `sanitizeFolderSegment` removes
> illegal characters rather than substituting, so the derived column is `I&TOperatingUnitsDepartment` —
> the same way `Estate/Mill` yields `EstateMill`.

Every segment ends with **Year → Document Type** below the leaf level — these remain
ensure-created subfolders inheriting the Unit folder's ACL, **not** term levels.

**Spelling: `NBPOL`** (New Britain Palm Oil), not "NBBOL" — this string ends up in term-set
names, folder names, and group names; get it right at creation.

## Decisions

1. **One mode per segment (Option A).** Rejected: a single "Head Office" term set with the
   four offices as a top-level term (would add a cascade level, need new Staging columns,
   and break the one-mode-one-term-set pattern for no permission or metadata gain).
2. **Term store stays flat.** The 3 new sets are created as siblings of `Group Head Office`
   inside the existing **`DMS Metadata`** term group. Term groups cannot nest, so a "parent
   group" is not possible without either a second term group (splits administration) or
   Option B (rejected above). The code matches sets by GUID; grouping is cosmetic.
3. **Head offices share existing columns.** All four use `Levels = [Department, Unit]`,
   which map to the live `Department`/`DepartmentTid` and `Unit`/`UnitTid` Staging columns.
   **Zero new columns and zero code changes for the 2026 pilot** (given data-driven
   reconciliation from the onboarding plan §6; otherwise `RECON_MODES` needs the 3 GUIDs).
4. **Buffer layers.** Client wants 2 spare levels of future depth. No build needed: the
   `Levels` JSON is variable-depth, and deeper chains are a config row + new columns when
   they materialise. Watch-item: deeper folder paths — the OData parameter-alias rule
   (CLAUDE.md gotcha #9) must be used everywhere or deep paths 400.

## Pilot work items (in onboarding-plan §4 order)

1. ~~**Term store (manual, now)**~~ — DONE 2026-07-24: the 3 sets exist in `DMS Metadata`
   (GUIDs in the table above), pending the client's Department→Unit trees.
   ⚠ The first set was created named **"Upstream Head Office"** — rename it to
   **"Upstream Malaysia Head Office"** to match the client's Business Segment list
   (safe: the GUID is stable; the segment is distinct from "Upstream Malaysia Operations").
2. **DMS Config:** 3 new `mode` rows per the §2 schema — `Side=BusinessSegment`,
   `StagingFolder` = the segment name verbatim, the captured `TermSetGuid`, and the same
   `Levels` JSON as GHO (with `labelCol`/`tidCol` set). No trailing spaces.
3. ~~**Code (small)**~~ — DONE 2026-07-24: the 3 GUIDs added to `DEFAULT_MODES` in
   `Form.tsx` + `BulkUpload.tsx` and `RECON_MODES` in `FolderManager.tsx` (fallback parity
   only — config rows are the source of truth). 2027 placeholders renumbered to
   `sortOrder` 5–8.
4. **Groups + Group Map + reconciliation:** per the onboarding checklist, once the client
   supplies each head office's Department→Unit tree and who belongs to each unit.

## Blocking client input

- Department → Unit trees for Upstream Malaysia HO, Minamas HO, NBPOL HO.
- (2027, unchanged from onboarding plan §7): structures + term data for segments 5–12.

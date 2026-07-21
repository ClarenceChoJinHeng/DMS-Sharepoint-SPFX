# Group Map Builder — Design

**Date:** 2026-07-22
**Goal:** Remove the manual, error-prone step of adding rows to the **DMS Group Map** list. Give admins a dropdown-driven tool that maps an **existing** Entra security group to a segment/tier/role and writes a clean row — no hand-typed GUIDs, no trailing-space bugs.

Related: `2026-07-21-segment-onboarding-plan.md` (§8.2 raised this tool), `2026-07-16-tiered-group-detection-design.md`.

---

## 1. Scope

**In scope (Option A — map existing group):**
- Search for an existing Entra (M365) security group and capture its Object ID.
- Pick the segment (from DMS Config `mode` rows), cascade the term chain to the governed tier, pick the role.
- Write one DMS Group Map row (all 5 fields, trimmed).
- List existing DMS Group Map rows and delete a wrong one.

**Out of scope (deferred):**
- **Option B — creating the Entra group itself** (needs `Group.Create` / `Group.ReadWrite.All`, member management, group-lifecycle governance). Not now.
- Editing a row in place — delete + re-add covers it.

**No new permissions:** group search uses `Group.Read.All` (already granted); writing/deleting rows uses SharePoint REST (`spHttpClient`) with the admin's existing Contribute on the list.

---

## 2. Placement

A new **"Group Map"** tab inside the existing **DMS Admin Tool** (Folder Manager web part), alongside Reconciliation and Folder Map. Reuses `searchGroups` and the term-store cascade (`loadReconTops` / `loadReconChildren`) already in that component. No new web part, no new manifest, no version-gated deploy surface beyond the normal bump.

---

## 3. DMS Group Map schema (the data contract)

Verified against both readers (`Form.tsx loadGroupMap`, `FolderManager.tsx loadGroupMapForAssign`).

| Column | Meaning | Source in the builder |
|---|---|---|
| `GroupId` | Entra Object ID — the match key | picked group's `id` |
| `GroupName` | Cosmetic display name | picked group's `displayName` |
| `Segment` | Term-set GUID of the mode this row scopes to | selected mode's `termSetGuid` |
| `UnitTermGuid` | Term GUID of the governed tier — **or the term-set GUID for a segment-tier row** | the term the admin stops the cascade on (segment tier ⇒ `= Segment`) |
| `Role` | `MEMBER` \| `UPL` \| `APR` \| `GLOBAL` | selected role |

Synthetic row:
```
{ GroupId: "00000000-aaaa-bbbb-cccc-000000000001",
  GroupName: "DMS_LRC_GCO_UPL",
  Segment: "efa87c6a-9536-4f7c-910f-011bf7413b80",
  UnitTermGuid: "479a74d8-4126-40f0-9b1f-3374de74152b",
  Role: "UPL" }
```

> **GLOBAL rows carry no term.** For a `GLOBAL` selection the tool skips the segment/tier steps and writes `Segment = ""`, `UnitTermGuid = ""` (matching how `collectMembership` treats GLOBAL — it ignores the term and flags a privileged uploader).

---

## 4. Add-a-row flow (progressive disclosure)

Each step reveals the next; nothing is a free-text GUID field.

1. **Group** — type-to-search Entra via `searchGroups`; pick one → `GroupId` + `GroupName`.
2. **Role** — `MEMBER` / `UPL` / `APR` / `GLOBAL`.
   - If `GLOBAL`: skip steps 3–4; the row is group-only.
3. **Segment** — dropdown of DMS Config `mode` rows (label = `ModeLabel`, value = `termSetGuid`). Sets `Segment` and seeds the cascade root.
4. **Tier** — cascade down the term chain. At each level the admin can **drill deeper** (load children) or **"Assign at this tier."** The stopping term = `UnitTermGuid`. Selecting the **segment itself** (top) sets `UnitTermGuid = Segment` (segment-tier row).
5. **Add** — validate, then POST one row.

**Live preview:** before Add, show the exact row that will be written (GroupName · Segment label · tier label · Role) so the admin confirms it visually.

---

## 5. Existing rows list + delete

- On tab load, GET all DMS Group Map rows (`$select=Id,GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`).
- Render a table: GroupName · Segment (label if resolvable, else GUID) · tier (label if resolvable, else GUID) · Role.
- **Delete** button per row → confirm → `DELETE` the list item → refresh.
- Purpose: spot duplicates and stale/typo rows (exactly the `DMS_GHO_LRC` trailing-space class of bug) and clean them without opening the raw SharePoint list.

Label resolution is best-effort: resolve term GUIDs to labels via the term store where cheap; fall back to showing the GUID. Never block the list on label lookups.

---

## 6. Guardrails

- **Trim every field** before writing — the single most important fix (the trailing space in `UnitTermGuid` silently dropped `DMS_GHO_LRC`).
- **Duplicate check** — if a row with the same `GroupId` + `UnitTermGuid` + `Role` (all normalised: trimmed + lowercased) already exists, warn and **do not** write. (Same group can legitimately hold different roles or map different tiers — only the exact triple is a dupe.)
- **Confirm before delete.**
- **Disable Add** until a group, a role, and (for non-GLOBAL) a segment + tier are all chosen.

> **The builder only writes the mapping row — it does not apply any permission.** The
> `Role` is a label; the native SharePoint level (`MEMBER`→Read, `UPL`→Contribute,
> `APR`→Design; `GLOBAL`→none) is applied by **Folder Reconciliation** via
> `ROLE_TO_PERMISSION`, using SharePoint's built-in role definitions. So after adding
> or deleting rows, the admin must **re-run reconciliation** for the change to take
> effect on the folders. The tab shows a reminder to that effect after a successful add/delete.

---

## 7. Testable core — `src/shared/groupMapModel.ts` (pure, no `@microsoft/*`)

Mirrors the `formModel` / `shareGuard` / `pathEncoding` pattern so it runs in plain Jest.

```ts
export type GroupMapRole = "MEMBER" | "UPL" | "APR" | "GLOBAL";

export interface GroupMapWriteRow {
  GroupId: string;
  GroupName: string;
  Segment: string;       // "" for GLOBAL
  UnitTermGuid: string;  // "" for GLOBAL; = Segment for a segment-tier row
  Role: GroupMapRole;
}

export interface GroupMapDraft {
  groupId: string;
  groupName: string;
  role: GroupMapRole;
  segmentGuid?: string;   // required unless GLOBAL
  tierGuid?: string;      // required unless GLOBAL; may equal segmentGuid
}

/** Build the exact row to POST. Trims all fields. GLOBAL ⇒ Segment/UnitTermGuid "". */
export function buildGroupMapRow(draft: GroupMapDraft): GroupMapWriteRow;

/** True if an equivalent row exists (GroupId + UnitTermGuid + Role, normalised). */
export function isDuplicateRow(existing: GroupMapWriteRow[], candidate: GroupMapWriteRow): boolean;

/** Field-level validation for enabling the Add button. Returns [] when valid. */
export function validateDraft(draft: GroupMapDraft): string[];
```

The web part supplies Graph search, mode loading, term cascade, and the SP REST POST/DELETE around these pure functions.

---

## 8. Files

| File | Change |
|---|---|
| `src/shared/groupMapModel.ts` | **New** — pure helpers above |
| `src/shared/groupMapModel.test.ts` | **New** — unit tests (build, dupe, validate, GLOBAL, trim) |
| `src/webparts/folderManager/components/GroupMapBuilder.tsx` | **New** — the tab UI (flow + list + delete) |
| `src/webparts/folderManager/components/FolderManager.tsx` | Add the "Group Map" tab; expose/reuse `searchGroups`, mode load, term cascade, SP REST write/delete helpers |
| `config/package-solution.json` | Version bump |

---

## 9. Open decisions

- **None blocking.** GLOBAL role is included as a minor separate option per the design review.

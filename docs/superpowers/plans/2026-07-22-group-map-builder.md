# Group Map Builder Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A dropdown-driven "Group Map" tab in the DMS Admin Tool that maps an existing Entra group to a segment/tier/role and writes a clean DMS Group Map row — no hand-typed GUIDs, no trailing-space bugs.

**Architecture:** A pure, unit-tested core (`src/shared/groupMapModel.ts`) builds/validates/dedupes the row. A new self-contained React component (`GroupMapBuilder.tsx`) supplies Graph group search, DMS Config mode loading, term-store cascade, and the SharePoint REST POST/DELETE around that core. It mounts as a new tab in `FolderManager.tsx`.

**Tech Stack:** SPFx 1.23.0, React 17, TypeScript, Heft build (NOT gulp), Node 22, Jest for pure modules. Design: `docs/superpowers/specs/2026-07-22-group-map-builder-design.md`.

---

## Task 1: Pure model — `groupMapModel.ts`

**Files:**
- Create: `src/shared/groupMapModel.ts`
- Test: `src/shared/groupMapModel.test.ts`

- [ ] **Step 1: Write the failing tests**

```ts
// src/shared/groupMapModel.test.ts
import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  GroupMapWriteRow,
  GroupMapDraft,
} from "./groupMapModel";

describe("buildGroupMapRow", () => {
  it("builds a trimmed row for a unit-tier UPL mapping", () => {
    const draft: GroupMapDraft = {
      groupId: "  00000000-aaaa-bbbb-cccc-000000000001 ",
      groupName: " DMS_LRC_GCO_UPL ",
      role: "UPL",
      segmentGuid: " efa87c6a-9536-4f7c-910f-011bf7413b80 ",
      tierGuid: " 479a74d8-4126-40f0-9b1f-3374de74152b ",
    };
    expect(buildGroupMapRow(draft)).toEqual({
      GroupId: "00000000-aaaa-bbbb-cccc-000000000001",
      GroupName: "DMS_LRC_GCO_UPL",
      Segment: "efa87c6a-9536-4f7c-910f-011bf7413b80",
      UnitTermGuid: "479a74d8-4126-40f0-9b1f-3374de74152b",
      Role: "UPL",
    });
  });

  it("leaves Segment and UnitTermGuid empty for a GLOBAL row", () => {
    const draft: GroupMapDraft = {
      groupId: "g-1",
      groupName: "DMS_GLOBAL_UPLOADERS",
      role: "GLOBAL",
      segmentGuid: "should-be-ignored",
      tierGuid: "should-be-ignored",
    };
    expect(buildGroupMapRow(draft)).toEqual({
      GroupId: "g-1",
      GroupName: "DMS_GLOBAL_UPLOADERS",
      Segment: "",
      UnitTermGuid: "",
      Role: "GLOBAL",
    });
  });

  it("supports a segment-tier row where the tier GUID equals the segment GUID", () => {
    const row = buildGroupMapRow({
      groupId: "g-2",
      groupName: "DMS_GHO",
      role: "MEMBER",
      segmentGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
      tierGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
    });
    expect(row.Segment).toBe(row.UnitTermGuid);
  });
});

describe("isDuplicateRow", () => {
  const existing: GroupMapWriteRow[] = [
    { GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-gco", Role: "UPL" },
  ];

  it("matches on GroupId + UnitTermGuid + Role, case/space-insensitive", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: " G-1 ", GroupName: "A", Segment: "s", UnitTermGuid: "T-GCO", Role: "UPL",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(true);
  });

  it("allows the same group with a different role", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-gco", Role: "APR",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(false);
  });

  it("allows the same group mapping a different tier", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-risk", Role: "UPL",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(false);
  });
});

describe("validateDraft", () => {
  it("returns [] for a complete non-global draft", () => {
    expect(
      validateDraft({ groupId: "g", groupName: "n", role: "MEMBER", segmentGuid: "s", tierGuid: "t" }),
    ).toEqual([]);
  });

  it("requires a group and a role", () => {
    const errs = validateDraft({ groupId: "  ", groupName: "", role: undefined as unknown as "MEMBER" });
    expect(errs).toContain("Select a group.");
    expect(errs).toContain("Select a role.");
  });

  it("requires segment + tier unless GLOBAL", () => {
    expect(validateDraft({ groupId: "g", groupName: "n", role: "UPL" })).toEqual([
      "Select a segment.",
      "Select a tier.",
    ]);
  });

  it("does not require segment/tier for GLOBAL", () => {
    expect(validateDraft({ groupId: "g", groupName: "n", role: "GLOBAL" })).toEqual([]);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx jest groupMapModel`
Expected: FAIL — "Cannot find module './groupMapModel'".

- [ ] **Step 3: Write the implementation**

```ts
// src/shared/groupMapModel.ts
// Pure, SPFx-free helpers for the Group Map Builder. No @microsoft/* imports —
// keep this unit-testable in plain Jest.

export type GroupMapRole = "MEMBER" | "UPL" | "APR" | "GLOBAL";

/** The exact field set POSTed to the DMS Group Map list. */
export interface GroupMapWriteRow {
  GroupId: string;
  GroupName: string;
  Segment: string;       // "" for GLOBAL
  UnitTermGuid: string;  // "" for GLOBAL; equals Segment for a segment-tier row
  Role: GroupMapRole;
}

/** In-progress selections from the builder UI. */
export interface GroupMapDraft {
  groupId: string;
  groupName: string;
  role: GroupMapRole;
  segmentGuid?: string;   // required unless GLOBAL
  tierGuid?: string;      // required unless GLOBAL; may equal segmentGuid
}

const norm = (s: string): string => (s ?? "").trim();

/**
 * Build the row to POST. Trims every field (kills the trailing-space bug that
 * silently dropped DMS_GHO_LRC). A GLOBAL row carries no term, so Segment and
 * UnitTermGuid are forced empty regardless of any stale draft selections.
 */
export function buildGroupMapRow(draft: GroupMapDraft): GroupMapWriteRow {
  const isGlobal = draft.role === "GLOBAL";
  return {
    GroupId: norm(draft.groupId),
    GroupName: norm(draft.groupName),
    Segment: isGlobal ? "" : norm(draft.segmentGuid ?? ""),
    UnitTermGuid: isGlobal ? "" : norm(draft.tierGuid ?? ""),
    Role: draft.role,
  };
}

/** True if an equivalent row exists: GroupId + UnitTermGuid + Role, normalised. */
export function isDuplicateRow(
  existing: GroupMapWriteRow[],
  candidate: GroupMapWriteRow,
): boolean {
  const key = (r: GroupMapWriteRow): string =>
    `${norm(r.GroupId).toLowerCase()}|${norm(r.UnitTermGuid).toLowerCase()}|${norm(r.Role).toUpperCase()}`;
  const k = key(candidate);
  return existing.some((r) => key(r) === k);
}

/** Field-level validation for enabling the Add button. Returns [] when valid. */
export function validateDraft(draft: GroupMapDraft): string[] {
  const errors: string[] = [];
  if (!norm(draft.groupId)) errors.push("Select a group.");
  if (!draft.role) errors.push("Select a role.");
  if (draft.role !== "GLOBAL") {
    if (!norm(draft.segmentGuid ?? "")) errors.push("Select a segment.");
    if (!norm(draft.tierGuid ?? "")) errors.push("Select a tier.");
  }
  return errors;
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx jest groupMapModel`
Expected: PASS — all groupMapModel tests green; overall pure-suite count rises by the new tests.

- [ ] **Step 5: Commit**

```bash
git add src/shared/groupMapModel.ts src/shared/groupMapModel.test.ts
git commit -m "feat: groupMapModel pure helpers (build/dedupe/validate DMS Group Map rows)"
```

---

## Task 2: `GroupMapBuilder.tsx` component

**Files:**
- Create: `src/webparts/folderManager/components/GroupMapBuilder.tsx`
- Reference (patterns to copy): `src/webparts/onboarding/components/Onboarding.tsx` (`searchGroups` at line 206), `src/webparts/folderManager/components/FolderManager.tsx` (`loadReconModes`, `loadReconTops`, `loadReconChildren`)

The component is self-contained: it takes `context: WebPartContext` and `siteUrl: string` props and implements its own data access around the Task 1 pure functions. No changes to FolderManager's internals beyond mounting it (Task 3).

- [ ] **Step 1: Scaffold the component and props**

```tsx
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from "@microsoft/sp-http";
import {
  buildGroupMapRow, isDuplicateRow, validateDraft,
  GroupMapRole, GroupMapDraft, GroupMapWriteRow,
} from "../../../shared/groupMapModel";

type Props = { context: WebPartContext; siteUrl: string };
type GroupPick = { id: string; displayName: string };
type ModePick = { label: string; termSetGuid: string };
type TermLite = { id: string; label: string };
type ExistingRow = GroupMapWriteRow & { itemId: number };
const ROLES: GroupMapRole[] = ["MEMBER", "UPL", "APR", "GLOBAL"];
```

- [ ] **Step 2: Data-access helpers (copy the proven patterns)**

Implement, each mirroring existing code:
- `searchGroups(query)` → `MSGraphClientV3` `/groups` search (copy from `Onboarding.tsx:206` — `.header("ConsistencyLevel","eventual").count(true).select("id,displayName").top(25)`, `.search("\"displayName:<q>\"")`), returns `GroupPick[]`.
- `loadModes()` → GET `DMS%20Config/items?$select=ModeLabel,TermSetGuid,Levels&$filter=ConfigType eq 'mode'&$orderby=SortOrder` (`odata=nometadata`); keep rows with non-empty trimmed `TermSetGuid` AND non-empty `Levels`; map to `{ label: ModeLabel, termSetGuid: TermSetGuid.trim() }`.
- `loadTops(termSetGuid)` → GET `/_api/v2.1/termStore/sets/${termSetGuid}/children?$select=id,labels`; map each to `{ id, label: labels[0].name }`.
- `loadChildren(termSetGuid, parentId)` → GET `/_api/v2.1/termStore/sets/${termSetGuid}/terms/${parentId}/children?$select=id,labels`; same mapping.
- `loadExisting()` → GET `DMS%20Group%20Map/items?$select=Id,GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000` (`odata=nometadata`) → map to `ExistingRow[]` (`itemId: Id`).
- `postRow(row: GroupMapWriteRow)` → `context.spHttpClient.post(`${siteUrl}/_api/web/lists/getbytitle('DMS%20Group%20Map')/items`, v1, { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata" }, body: JSON.stringify(row) })`. (SPHttpClient supplies the request digest automatically.)
- `deleteRow(itemId: number)` → `context.spHttpClient.post(`${siteUrl}/_api/web/lists/getbytitle('DMS%20Group%20Map')/items(${itemId})`, v1, { headers: { "IF-MATCH": "*", "X-HTTP-Method": "DELETE", Accept: "application/json;odata=nometadata" } })`.

> All GET/POST use `context.spHttpClient` + `SPHttpClient.configurations.v1`, matching FolderManager. Term-store URLs use the same `/_api/v2.1/termStore` shape already used by `loadReconTops`. Check `res.ok` on each; on failure show a toast, do not throw uncaught.

- [ ] **Step 3: Component state + cascade logic**

State: `group: GroupPick | undefined`, `role: GroupMapRole | ""`, `mode: ModePick | undefined`, `cascade: TermLite[][]` (option lists per level), `chosen: TermLite[]` (picked term per level), `tierGuid: string`, `existing: ExistingRow[]`, `modes: ModePick[]`, `busy: boolean`, `toast`.

Cascade rules:
- On mount: `loadModes()` → `modes`; `loadExisting()` → `existing`.
- Selecting a `mode` → `loadTops(mode.termSetGuid)` into `cascade[0]`; clear `chosen`, deeper levels; `tierGuid = ""`.
- Selecting a term at level `i` → set `chosen[i]`, truncate `chosen`/`cascade` beyond `i`, `loadChildren` into `cascade[i+1]`, set `tierGuid = chosen[i].id` (assign at this tier by default; drilling deeper reassigns downward).
- An **"Assign at segment level"** control sets `tierGuid = mode.termSetGuid` and clears `chosen`/deeper cascade.
- Live draft: `{ groupId: group?.id ?? "", groupName: group?.displayName ?? "", role: (role || undefined), segmentGuid: mode?.termSetGuid, tierGuid }`.

- [ ] **Step 4: Render — add form + preview + existing list**

- Group search: input calling `searchGroups`, results dropdown; picking sets `group`.
- Role selector: 4 buttons from `ROLES`. When `role === "GLOBAL"`, hide the segment + cascade controls.
- Segment dropdown from `modes`.
- Cascade: one `<select>` per `cascade` level, plus an "Assign at segment level" button and an "Assign at this tier" affordance per selected level.
- **Preview line**: from `buildGroupMapRow(draft)` — GroupName · segment label (from `modes`) · tier label (from `chosen`/segment) · Role.
- **Add** button: `disabled={validateDraft(draft).length > 0 || busy}`. On click:
  ```tsx
  const row = buildGroupMapRow(draft);
  if (isDuplicateRow(existing, row)) { showToast("This exact mapping already exists.", true); return; }
  setBusy(true);
  await postRow(row);
  setExisting(await loadExisting());
  showToast("Row added — re-run Folder Reconciliation to apply permissions.", false);
  setBusy(false);
  ```
- **Existing rows table**: GroupName · Segment (resolve to mode label if a `modes` entry matches, else GUID) · tier (resolve to a `chosen`/loaded label if cheap, else GUID) · Role · Delete. Delete → inline confirm → `deleteRow(itemId)` → `loadExisting()` → toast reminding to re-run reconciliation.

- [ ] **Step 5: Build to typecheck**

Run: `npm run build`
Expected: build succeeds, no TS errors. (This component imports `@microsoft/*`, so it is verified by the compiler, not Jest.)

- [ ] **Step 6: Commit**

```bash
git add src/webparts/folderManager/components/GroupMapBuilder.tsx
git commit -m "feat: GroupMapBuilder component (search group, cascade tier, add/list/delete rows)"
```

---

## Task 3: Mount the tab + version bump

**Files:**
- Modify: `src/webparts/folderManager/components/FolderManager.tsx` (Tab type, tab bar array line ~1207, render branch line ~1222, import)
- Modify: `config/package-solution.json` (version)

- [ ] **Step 1: Extend the Tab type**

Find the `Tab` type (top types block) and add `"GroupMap"`:
```ts
type Tab = "Staging" | "Documents" | "Reconciliation" | "GroupMap";
```

- [ ] **Step 2: Add the tab button**

In the tab-bar array (`FolderManager.tsx:1207`), add `"GroupMap"`:
```tsx
{(["Staging", "Documents", "Reconciliation", "GroupMap"] as Tab[]).map((t, i, arr) => (
```
Update the label (line 1216) so `GroupMap` renders as `"Group Map"`:
```tsx
{t === "Reconciliation" ? "Folder Reconciliation" : t === "GroupMap" ? "Group Map" : t}
```
In the button `onClick` (lines 1209-1213), the existing library-only side-effect must skip `GroupMap` too:
```tsx
if (t !== "Reconciliation" && t !== "GroupMap") { setLibTarget(t as LibTarget); setExpandedIds({}); }
```

- [ ] **Step 3: Import and render the component**

Add import near the other component imports at the top of the file:
```ts
import GroupMapBuilder from "./GroupMapBuilder";
```
Add a render branch at the start of the conditional at line 1222, so `GroupMap` renders the builder:
```tsx
{tab === "GroupMap" ? (
  <GroupMapBuilder context={context} siteUrl={siteUrl} />
) : tab === "Reconciliation" ? (
  // ...existing reconciliation block, unchanged...
) : (
  // ...existing Staging/Documents tree block, unchanged...
)}
```
Verify the final `: ( ...tree... )` branch is NOT rendered for `GroupMap` (the new leading ternary guarantees this).

- [ ] **Step 4: Confirm `context`/`siteUrl` in scope**

`siteUrl` is defined in the component; `context` comes from `props`. Both are already used throughout FolderManager, so they are in scope where `<GroupMapBuilder>` renders.

- [ ] **Step 5: Bump the version**

In `config/package-solution.json`, bump `solution.version` `1.0.22.0` → `1.0.23.0`.

- [ ] **Step 6: Build**

Run: `npm run build`
Expected: build succeeds; `sd-gatrie.sppkg` at 1.0.23.0.

- [ ] **Step 7: Run the full test suite**

Run: `npx jest`
Expected: all pre-existing pure suites pass plus the new `groupMapModel` tests. (The 8 `@microsoft/*` suites still show as transform-failures — pre-existing and unrelated.)

- [ ] **Step 8: Commit**

```bash
git add src/webparts/folderManager/components/FolderManager.tsx config/package-solution.json
git commit -m "feat: mount Group Map tab in DMS Admin Tool; bump to 1.0.23.0"
```

---

## Manual verification (after deploy)

1. Upload the new `.sppkg` to the app catalog; open the DMS Admin Tool → **Group Map** tab.
2. Search a known group (e.g. an LRC unit group), pick **UPL**, pick **Group Head Office** → cascade to a unit → **Add**. Confirm the row appears in the existing list with the right tier label.
3. Try adding the exact same mapping again → expect the "already exists" warning, no dupe.
4. Delete the test row → confirm it disappears.
5. Open the raw **DMS Group Map** list → confirm the written fields are clean (no trailing spaces, GroupId is the Object ID).
6. Re-run **Folder Reconciliation** → confirm the group is assigned at the expected native level (MEMBER→Read / UPL→Contribute / APR→Design).

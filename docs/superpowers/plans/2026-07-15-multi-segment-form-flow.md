# Multi-Segment Upload Form — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convert the upload form from a fixed 2-mode (Department/Projects) cascade into a data-driven 5-mode flow (4 Business Segments + Projects), where each mode declares its own variable-depth level chain, the user's path is auto-detected via a group→term map keyed by Object ID, and files route into `…/Unit/Year/Document Type/` with per-level label+GUID metadata columns.

**Architecture:** Extract all pure, runtime-independent logic (Levels parsing, group→path matching, folder-name building, metadata payload building) into a new SPFx-free module `src/shared/formModel.ts`, unit-tested with Jest. `Form.tsx` and `Reconciliation.tsx` consume that module; their SharePoint-REST/Graph glue is verified by build + SharePoint workbench. Folder isolation is unchanged (Unit is the deepest permissioned level; Year/Document Type folders are ensure-created and inherit the Unit ACL).

**Tech Stack:** SPFx 1.23.0, React 17, TypeScript 5.8, Heft (no Gulp), `@rushstack/heft-jest-plugin` for tests, SharePoint REST v1 + Term Store v2.1 + MS Graph.

**Companion docs:** Spec `docs/superpowers/specs/2026-07-15-multi-segment-form-flow-design.md`; permissions `docs/superpowers/specs/2026-07-15-staging-per-unit-groups.md`.

---

## File map

| File | Responsibility | Change |
|---|---|---|
| `src/shared/formModel.ts` | Pure types + functions: `Level`, `ModeV2`, `GroupMapRow`, `UserPath`, `parseLevels`, `matchUserPaths`, `sanitizeFolderSegment`, `buildLevelFormValues` | **Create** |
| `src/shared/formModel.test.ts` | Jest unit tests for the above | **Create** |
| `src/shared/dmsFolderMap.ts` | Rename-proof folder map + new `resolveFolderServerUrl` + `ensureFolder` helpers | **Modify** |
| `src/webparts/form/components/Form.tsx` | Mode loader (Side/Levels), group-map reader, N-level cascade, on-demand folder creation, payload builder, toggle | **Modify (large)** |
| `src/webparts/reconciliation/components/Reconciliation.tsx` | 5-mode `DEFAULT_MODES` | **Modify** |
| `config/jest.config.json`, `config/heft.json` | Wire the Jest test phase | **Create** |
| `config/package-solution.json` | Version bump | **Modify** |
| `CLAUDE.md` | Update term-set GUIDs, mode/column notes | **Modify** |

---

## Phase A — Tenant data setup (manual, in the browser / PnP; NO app code)

> These are prerequisites the app reads at runtime. Do them first so Phase C has real data to load. All are done by hand per the user's chosen workflow. Record every GUID you capture in a scratch note — later tasks paste them in.

### Task A1: Create the 16 metadata columns on the Staging library

**Where:** Staging library → Settings → Create column (or a PnP PowerShell `Add-PnPField`). Type = **Single line of text** for all 16.

- [ ] **Step 1: Create the 8 label columns** (Single line of text), display names exactly:
  `Business Segment`, `Department`, `Unit`, `Region`, `Estate/Mill`, `Refinery`, `I&T Operating Unit`, `Project Name`.

- [ ] **Step 2: Create the 8 companion GUID columns** (Single line of text), display names:
  `Business Segment Tid`, `Department Tid`, `Unit Tid`, `Region Tid`, `Estate/Mill Tid`, `Refinery Tid`, `I&T Operating Unit Tid`, `Project Name Tid`. Mark these **Hidden** in the column's advanced settings (they exist only for rename-safe matching).

- [ ] **Step 3: Read back the frozen internal names.** SharePoint encodes spaces/`&`/`/` into the internal name (e.g. `Estate/Mill` → `Estate_x002f_Mill`). Fetch them:

Run in browser (logged into the site):
```
{site}/_api/web/lists/getbytitle('Staging')/fields?$select=Title,InternalName,Hidden&$filter=Group ne '_Hidden'&$top=500
```
Expected: a JSON list; copy each new column's `InternalName` into your scratch note. **Do not guess these — the `FIELDS`/`LEVEL_COLUMNS` tables in Form.tsx (Task C5) must use the internal names, and the `_Tid` convention appends `_Tid` to each.**

- [ ] **Step 4: Decide the `column` token.** In `Levels` JSON we store a *logical* column key (e.g. `Department`) and Form.tsx maps it to the real internal name via the `LEVEL_COLUMNS` table (Task C5). So the `column` value in Levels does **not** need the encoded form — but `LEVEL_COLUMNS` in Task C5 **does**. Paste the internal names into that table when you reach it.

### Task A2: Capture the 5 term set GUIDs

- [ ] **Step 1:** In the Term Store admin, click each of the 5 sets and copy its **Unique Identifier**:
  `Group Head Office`, `Group Upstream Operations`, `Group SDGI Operations`, `Group Innovation & Technology`, `Group-led Projects`.
- [ ] **Step 2:** Record them in the scratch note labelled clearly. These feed Task A3 (DMS Config), Task D1 (Reconciliation fallback), Task C2/D2 (Form fallback), and CLAUDE.md.

### Task A3: Update DMS Config `mode` rows

**Where:** the `DMS Config` SharePoint list. Add two columns first, then replace the mode rows.

- [ ] **Step 1: Add list columns** to DMS Config (if absent): `Side` (Single line of text or Choice: `BusinessSegment`,`Project`) and `Levels` (Multiple lines of text, plain text).
- [ ] **Step 2: Delete the old 2 mode rows** (`department`, `project`) — `ConfigType eq 'mode'`.
- [ ] **Step 3: Create 5 mode rows** with these exact values (paste the GUIDs from A2 into `TermSetGuid`; leave `LookupStyle`/`SubTeamLabel` blank — retired):

| Title | ModeLabel | Side | TermSetGuid | StagingFolder | SortOrder | Levels (paste verbatim) |
|---|---|---|---|---|---|---|
| `gho` | Group Head Office | BusinessSegment | *(GHO GUID)* | `Group Head Office` | 1 | `[{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]` |
| `upstream` | Group Upstream Operations | BusinessSegment | *(Upstream GUID)* | `Group Upstream Operations` | 2 | `[{"label":"Region","column":"Region"},{"label":"Estate/Mill","column":"EstateMill"}]` |
| `sdgi` | Group SDGI Operations | BusinessSegment | *(SDGI GUID)* | `Group SDGI Operations` | 3 | `[{"label":"Refinery","column":"Refinery"},{"label":"Department","column":"Department"}]` |
| `it` | Group Innovation & Technology | BusinessSegment | *(I&T GUID)* | `Group Innovation & Technology` | 4 | `[{"label":"I&T Operating Unit","column":"ITOperatingUnit"}]` |
| `projects` | Group-led Projects | Project | *(Projects GUID)* | `Group-led Projects` | 5 | `[{"label":"Project Name","column":"ProjectName"},{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]` |

> The `column` tokens above are **logical keys** (no encoding) — Form.tsx's `LEVEL_COLUMNS` table (Task C5) maps each key to a real internal name + its `_Tid` sibling.

### Task A4: Create the DMS Group Map list

- [ ] **Step 1:** Create a new SharePoint list named **`DMS Group Map`**.
- [ ] **Step 2:** Add columns (all Single line of text unless noted): `GroupId`, `GroupName`, `Segment` (the term set GUID), `UnitTermGuid`, `Role` (Choice: `UPL`,`APR`). (`Title` can hold `GroupName` too.)
- [ ] **Step 3:** Add **at least one real row** for a group you can test with (e.g. `DMS_GF_TAX_UPL`): paste its Entra **Object ID** into `GroupId`, the GHO set GUID into `Segment`, and the **Tax** term's GUID into `UnitTermGuid`, `Role=UPL`. This is the row Phase C is verified against.

---

## Phase B — Pure model module (`formModel.ts`) with tests

### Task B1: Wire the Jest test phase and prove it runs

**Files:**
- Create: `config/jest.config.json`
- Create: `config/heft.json`
- Create: `src/shared/formModel.test.ts` (temporary smoke test)

- [ ] **Step 1: Install the Heft Jest plugin** (matches Heft 1.2.x):

Run: `npm install --save-dev @rushstack/heft-jest-plugin`
Expected: added to devDependencies, no peer-dep errors.

- [ ] **Step 2: Create `config/heft.json`** registering the Jest lifecycle (extends the rig, adds the test plugin):

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/heft/v0/heft.schema.json",
  "extends": "@microsoft/spfx-web-build-rig/profiles/default/config/heft.json",
  "phasesByName": {
    "test": {
      "tasksByName": {
        "jest": {
          "taskPlugin": {
            "pluginPackage": "@rushstack/heft-jest-plugin"
          }
        }
      }
    }
  }
}
```

- [ ] **Step 3: Create `config/jest.config.json`:**

```json
{
  "preset": "@rushstack/heft-jest-plugin",
  "testMatch": ["<rootDir>/src/**/*.test.ts"]
}
```

- [ ] **Step 4: Write a smoke test** `src/shared/formModel.test.ts`:

```ts
describe("jest smoke", () => {
  it("runs", () => {
    expect(1 + 1).toBe(2);
  });
});
```

- [ ] **Step 5: Run the test phase.**

Run: `npx heft test --clean`
Expected: Jest executes, `jest smoke › runs` PASS. **If Heft cannot load the plugin in this environment, fall back:** rename tests to `.test.mts`, use `node:test`+`node:assert`, and run `node --experimental-strip-types --test src/shared/*.test.mts` (Node 22.14 supports type-stripping for annotation-only files — `formModel.ts` qualifies). Record whichever runner works; later tasks say `npx heft test` — substitute your runner if you took the fallback.

- [ ] **Step 6: Commit.**

```bash
git add config/jest.config.json config/heft.json package.json package-lock.json src/shared/formModel.test.ts
git commit -m "test: wire heft jest phase with smoke test"
```

### Task B2: Types + `parseLevels`

**Files:**
- Create: `src/shared/formModel.ts`
- Modify: `src/shared/formModel.test.ts`

- [ ] **Step 1: Write the failing test** (replace the smoke test file contents):

```ts
import { parseLevels, Level } from "./formModel";

describe("parseLevels", () => {
  it("parses a valid Levels JSON array", () => {
    const json = '[{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]';
    const result: Level[] = parseLevels(json);
    expect(result).toEqual([
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ]);
  });

  it("returns [] for empty, null, or malformed JSON", () => {
    expect(parseLevels("")).toEqual([]);
    expect(parseLevels(undefined as unknown as string)).toEqual([]);
    expect(parseLevels("{not json}")).toEqual([]);
    expect(parseLevels('"a string"')).toEqual([]);
  });

  it("drops entries missing label or column", () => {
    const json = '[{"label":"Region","column":"Region"},{"label":"x"},{"column":"y"}]';
    expect(parseLevels(json)).toEqual([{ label: "Region", column: "Region" }]);
  });
});
```

- [ ] **Step 2: Run to verify it fails.**

Run: `npx heft test`
Expected: FAIL — `Cannot find module './formModel'`.

- [ ] **Step 3: Create `src/shared/formModel.ts`** with the types and `parseLevels`:

```ts
// Pure, SPFx-free model helpers for the multi-segment upload form.
// No imports from @microsoft/* here — keep this unit-testable in plain Jest.

/** One cascade level within a mode: a labelled dropdown that writes one column. */
export interface Level {
  label: string;   // shown to the user, e.g. "Estate/Mill"
  column: string;  // logical column key, mapped to a real field in Form.tsx LEVEL_COLUMNS
}

/** A configured upload mode (one term set + its level chain). */
export interface ModeV2 {
  key: string;
  label: string;
  side: "BusinessSegment" | "Project";
  termSetGuid: string;
  stagingFolder: string;
  levels: Level[];
  sortOrder: number;
}

/** One row of the DMS Group Map list. */
export interface GroupMapRow {
  groupId: string;      // Entra Object ID — the match key
  groupName: string;
  segment: string;      // term set GUID
  unitTermGuid: string;
  role: string;         // "UPL" | "APR" | ""
}

/** A user's resolved path from a matched group. */
export interface UserPath {
  segment: string;      // term set GUID (identifies the mode)
  unitTermGuid: string;
  role: string;
}

/** Safely parse a DMS Config `Levels` JSON string into Level[]. Never throws. */
export function parseLevels(json: string): Level[] {
  if (!json || typeof json !== "string") return [];
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return [];
  }
  if (!Array.isArray(raw)) return [];
  return raw
    .filter(
      (e): e is Level =>
        !!e &&
        typeof (e as Level).label === "string" &&
        typeof (e as Level).column === "string",
    )
    .map((e) => ({ label: e.label, column: e.column }));
}
```

- [ ] **Step 4: Run to verify it passes.**

Run: `npx heft test`
Expected: PASS — all `parseLevels` tests green.

- [ ] **Step 5: Commit.**

```bash
git add src/shared/formModel.ts src/shared/formModel.test.ts
git commit -m "feat: formModel types + parseLevels (tested)"
```

### Task B3: `matchUserPaths`

**Files:** Modify `src/shared/formModel.ts`, `src/shared/formModel.test.ts`

- [ ] **Step 1: Write the failing test** (append to the test file):

```ts
import { matchUserPaths, GroupMapRow } from "./formModel";

describe("matchUserPaths", () => {
  const rows: GroupMapRow[] = [
    { groupId: "g-tax", groupName: "DMS_GF_TAX_UPL", segment: "set-gho", unitTermGuid: "t-tax", role: "UPL" },
    { groupId: "g-treas", groupName: "DMS_GF_TREAS_UPL", segment: "set-gho", unitTermGuid: "t-treas", role: "UPL" },
    { groupId: "g-est", groupName: "DMS_UP_EAST_UPL", segment: "set-up", unitTermGuid: "t-east", role: "UPL" },
  ];

  it("returns the paths for the user's matched group ids", () => {
    expect(matchUserPaths(rows, ["g-tax", "g-est", "unrelated"])).toEqual([
      { segment: "set-gho", unitTermGuid: "t-tax", role: "UPL" },
      { segment: "set-up", unitTermGuid: "t-east", role: "UPL" },
    ]);
  });

  it("is case-insensitive and trims group ids on both sides", () => {
    expect(matchUserPaths(rows, [" G-TAX "])).toEqual([
      { segment: "set-gho", unitTermGuid: "t-tax", role: "UPL" },
    ]);
  });

  it("returns [] when nothing matches", () => {
    expect(matchUserPaths(rows, ["none"])).toEqual([]);
    expect(matchUserPaths([], ["g-tax"])).toEqual([]);
  });
});
```

- [ ] **Step 2: Run to verify it fails.**

Run: `npx heft test`
Expected: FAIL — `matchUserPaths is not a function`.

- [ ] **Step 3: Implement** (append to `formModel.ts`):

```ts
/** Return the UserPath for every group the user belongs to, in row order. */
export function matchUserPaths(rows: GroupMapRow[], userGroupIds: string[]): UserPath[] {
  const wanted = new Set(userGroupIds.map((id) => (id ?? "").trim().toLowerCase()));
  return rows
    .filter((r) => wanted.has((r.groupId ?? "").trim().toLowerCase()))
    .map((r) => ({ segment: r.segment, unitTermGuid: r.unitTermGuid, role: r.role }));
}
```

- [ ] **Step 4: Run to verify it passes.**

Run: `npx heft test`
Expected: PASS.

- [ ] **Step 5: Commit.**

```bash
git add src/shared/formModel.ts src/shared/formModel.test.ts
git commit -m "feat: matchUserPaths group->term resolver (tested)"
```

### Task B4: `sanitizeFolderSegment`

**Files:** Modify `src/shared/formModel.ts`, `src/shared/formModel.test.ts`

Year and Document Type folder names come from term labels; strip characters SharePoint forbids in folder names before ensure-creating them.

- [ ] **Step 1: Write the failing test:**

```ts
import { sanitizeFolderSegment } from "./formModel";

describe("sanitizeFolderSegment", () => {
  it("removes illegal folder characters and trims", () => {
    expect(sanitizeFolderSegment('In*voice:2026?')).toBe("Invoice2026");
    expect(sanitizeFolderSegment("  2026  ")).toBe("2026");
  });
  it("collapses internal whitespace to single spaces", () => {
    expect(sanitizeFolderSegment("Board   Papers")).toBe("Board Papers");
  });
  it("returns '' for empty/whitespace input", () => {
    expect(sanitizeFolderSegment("   ")).toBe("");
    expect(sanitizeFolderSegment("")).toBe("");
  });
});
```

- [ ] **Step 2: Run to verify it fails.**

Run: `npx heft test`
Expected: FAIL — `sanitizeFolderSegment is not a function`.

- [ ] **Step 3: Implement** (append; SharePoint forbids `\ / : * ? " < > | # %` in folder names — `&` and `-` are legal and preserved):

```ts
const ILLEGAL_FOLDER_CHARS = /[\\/:*?"<>|#%]/g;

/** Make a term label safe to use as a single folder name. Returns "" if nothing remains. */
export function sanitizeFolderSegment(name: string): string {
  if (!name) return "";
  return name.replace(ILLEGAL_FOLDER_CHARS, "").replace(/\s+/g, " ").trim();
}
```

- [ ] **Step 4: Run to verify it passes.**

Run: `npx heft test`
Expected: PASS.

- [ ] **Step 5: Commit.**

```bash
git add src/shared/formModel.ts src/shared/formModel.test.ts
git commit -m "feat: sanitizeFolderSegment for on-demand subfolders (tested)"
```

### Task B5: `buildLevelFormValues`

**Files:** Modify `src/shared/formModel.ts`, `src/shared/formModel.test.ts`

Builds the label + `_Tid` field-value pairs for the selected levels. Consumers pass a `columnMap` (logical key → real internal name) and the selected `{label,id}` per level.

- [ ] **Step 1: Write the failing test:**

```ts
import { buildLevelFormValues } from "./formModel";

describe("buildLevelFormValues", () => {
  const columnMap: Record<string, string> = {
    Department: "Department",
    Unit: "Unit",
    BusinessSegment: "Business_x0020_Segment",
  };
  it("emits a label value and a _Tid value per selection", () => {
    const result = buildLevelFormValues(
      columnMap,
      [
        { column: "Department", label: "Group Finance", id: "t-gf" },
        { column: "Unit", label: "Tax", id: "t-tax" },
      ],
    );
    expect(result).toEqual([
      { FieldName: "Department", FieldValue: "Group Finance" },
      { FieldName: "Department_Tid", FieldValue: "t-gf" },
      { FieldName: "Unit", FieldValue: "Tax" },
      { FieldName: "Unit_Tid", FieldValue: "t-tax" },
    ]);
  });
  it("uses the mapped internal name and skips unmapped columns", () => {
    const result = buildLevelFormValues(columnMap, [
      { column: "BusinessSegment", label: "Group Head Office", id: "set-gho" },
      { column: "Nope", label: "x", id: "y" },
    ]);
    expect(result).toEqual([
      { FieldName: "Business_x0020_Segment", FieldValue: "Group Head Office" },
      { FieldName: "Business_x0020_Segment_Tid", FieldValue: "set-gho" },
    ]);
  });
  it("skips selections with a blank label", () => {
    expect(buildLevelFormValues(columnMap, [{ column: "Unit", label: "", id: "t" }])).toEqual([]);
  });
});
```

- [ ] **Step 2: Run to verify it fails.**

Run: `npx heft test`
Expected: FAIL — `buildLevelFormValues is not a function`.

- [ ] **Step 3: Implement** (append):

```ts
export interface LevelSelection {
  column: string; // logical key
  label: string;
  id: string;
}
export interface SpFormValue {
  FieldName: string;
  FieldValue: string;
}

/**
 * Build validateUpdateListItem field pairs for the chosen levels:
 * one text pair for the internal-name column, one for its `<name>_Tid` sibling.
 * Skips selections whose column isn't in columnMap or whose label is blank.
 */
export function buildLevelFormValues(
  columnMap: Record<string, string>,
  selections: LevelSelection[],
): SpFormValue[] {
  const out: SpFormValue[] = [];
  for (const s of selections) {
    const internal = columnMap[s.column];
    if (!internal || !s.label) continue;
    out.push({ FieldName: internal, FieldValue: s.label });
    out.push({ FieldName: `${internal}_Tid`, FieldValue: s.id });
  }
  return out;
}
```

- [ ] **Step 4: Run to verify it passes.**

Run: `npx heft test`
Expected: PASS.

- [ ] **Step 5: Commit.**

```bash
git add src/shared/formModel.ts src/shared/formModel.test.ts
git commit -m "feat: buildLevelFormValues label+_Tid payload builder (tested)"
```

---

## Phase C — Wire `Form.tsx` to the new model

> These tasks touch runtime SharePoint/Graph code that isn't unit-tested. `Form.tsx` will **not compile between C2 and C6** (the edits are interdependent). Do **not** commit until the build is green at C6 Step 9. Keep the temporary `[DMS DEBUG]` console logs until Task C6.

### Task C1: Add folder helpers to `dmsFolderMap.ts`

**Files:** Modify `src/shared/dmsFolderMap.ts`

- [ ] **Step 1: Add `resolveFolderServerUrl` and `ensureFolder`** at the end of the file (before EOF):

```ts
/** Get a folder's CURRENT server-relative URL from its stable UniqueId. */
export async function resolveFolderServerUrl(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  uniqueId: string,
): Promise<string | null> {
  const res: SPHttpClientResponse = await spHttpClient.get(
    `${siteUrl}/_api/web/GetFolderById(guid'${uniqueId}')?$select=ServerRelativeUrl`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) return null;
  const d = await res.json();
  return d.ServerRelativeUrl ?? null;
}

/**
 * Ensure a subfolder `name` exists directly under `parentServerRelativeUrl`.
 * Idempotent: creates it, or resolves the existing one on 409/exists.
 * Returns the child's UniqueId + ServerRelativeUrl, or null on failure.
 */
export async function ensureFolder(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  parentServerRelativeUrl: string,
  name: string,
): Promise<{ uniqueId: string; serverRelativeUrl: string } | null> {
  const childPath = `${parentServerRelativeUrl}/${name}`;
  const addRes: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl=@u)?@u='${encodeURIComponent(childPath)}'&$select=UniqueId,ServerRelativeUrl`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (addRes.ok) {
    const d = await addRes.json();
    return { uniqueId: d.UniqueId, serverRelativeUrl: d.ServerRelativeUrl };
  }
  // Already exists (or transient) — try to resolve it by path.
  const existing = await resolveFolderByPath(spHttpClient, siteUrl, childPath);
  return existing;
}
```

- [ ] **Step 2: Build to verify it compiles.**

Run: `npm run build`
Expected: build succeeds (TypeScript clean).

- [ ] **Step 3: Commit.**

```bash
git add src/shared/dmsFolderMap.ts
git commit -m "feat: ensureFolder + resolveFolderServerUrl for on-demand subfolders"
```

### Task C2: Replace the mode type + `loadModes` with Side/Levels

**Files:** Modify `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Import the model** — add after the `lookupFolderMapping` import (line 5):

```ts
import {
  parseLevels,
  matchUserPaths,
  sanitizeFolderSegment,
  buildLevelFormValues,
  Level,
  GroupMapRow,
  UserPath,
} from "../../../shared/formModel";
```

Also add the two new folder helpers to the existing dmsFolderMap import (line 5):

```ts
import { lookupFolderMapping, resolveFolderServerUrl, ensureFolder } from "../../../shared/dmsFolderMap";
```

- [ ] **Step 2: Replace the `UploadMode` type** (lines ~59-67) with:

```ts
type UploadMode = {
  key: string;
  label: string;
  side: "BusinessSegment" | "Project";
  termSetGuid: string;
  stagingFolder: string;
  levels: Level[];
  sortOrder: number;
};
```

- [ ] **Step 3: Replace `DEFAULT_MODES`** (lines ~84-103) with the 5-mode fallback (paste real GUIDs from Task A2 in place of each `REPLACE-*`):

```ts
const DEFAULT_MODES: UploadMode[] = [
  { key: "gho", label: "Group Head Office", side: "BusinessSegment", termSetGuid: "REPLACE-GHO", stagingFolder: "Group Head Office",
    levels: [{ label: "Department", column: "Department" }, { label: "Unit", column: "Unit" }], sortOrder: 1 },
  { key: "upstream", label: "Group Upstream Operations", side: "BusinessSegment", termSetGuid: "REPLACE-UP", stagingFolder: "Group Upstream Operations",
    levels: [{ label: "Region", column: "Region" }, { label: "Estate/Mill", column: "EstateMill" }], sortOrder: 2 },
  { key: "sdgi", label: "Group SDGI Operations", side: "BusinessSegment", termSetGuid: "REPLACE-SDGI", stagingFolder: "Group SDGI Operations",
    levels: [{ label: "Refinery", column: "Refinery" }, { label: "Department", column: "Department" }], sortOrder: 3 },
  { key: "it", label: "Group Innovation & Technology", side: "BusinessSegment", termSetGuid: "REPLACE-IT", stagingFolder: "Group Innovation & Technology",
    levels: [{ label: "I&T Operating Unit", column: "ITOperatingUnit" }], sortOrder: 4 },
  { key: "projects", label: "Group-led Projects", side: "Project", termSetGuid: "REPLACE-PROJ", stagingFolder: "Group-led Projects",
    levels: [{ label: "Project Name", column: "ProjectName" }, { label: "Department", column: "Department" }, { label: "Unit", column: "Unit" }], sortOrder: 5 },
];
```

- [ ] **Step 4: Rewrite `loadModes`** (lines ~211-238) to read `Side` + `Levels`:

```ts
const loadModes = async (): Promise<UploadMode[]> => {
  const res: SPHttpClientResponse = await context.spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,ModeLabel,Side,TermSetGuid,StagingFolder,Levels,SortOrder&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json" } },
  );
  if (!res.ok) throw new Error("DMS Config list not found");
  const data = await res.json();
  return (data.value ?? []).map(
    (item: {
      Title: string; ModeLabel: string; Side: string; TermSetGuid: string;
      StagingFolder: string; Levels: string; SortOrder: number;
    }) => ({
      key: item.Title,
      label: item.ModeLabel,
      side: (item.Side === "Project" ? "Project" : "BusinessSegment") as "BusinessSegment" | "Project",
      termSetGuid: item.TermSetGuid,
      stagingFolder: item.StagingFolder,
      levels: parseLevels(item.Levels),
      sortOrder: item.SortOrder,
    }),
  );
};
```

Proceed to C3 (do not build yet — interdependent edits).

### Task C3: Group-map reader + generic N-level cascade

**Files:** Modify `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Add a `loadGroupMap` reader** near `loadModes`:

```ts
const loadGroupMap = async (): Promise<GroupMapRow[]> => {
  const res: SPHttpClientResponse = await context.spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('DMS%20Group%20Map')/items?$select=GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) throw new Error("DMS Group Map list not found");
  const data = await res.json();
  return (data.value ?? []).map((r: {
    GroupId: string; GroupName: string; Segment: string; UnitTermGuid: string; Role: string;
  }) => ({
    groupId: r.GroupId ?? "", groupName: r.GroupName ?? "", segment: r.Segment ?? "",
    unitTermGuid: r.UnitTermGuid ?? "", role: r.Role ?? "",
  }));
};
```

- [ ] **Step 2: Add Graph + admin readers** (replace the name-based `detectDepartment` — delete the entire `detectDepartment` function, lines ~326-391, and add):

```ts
const loadUserGroupIds = async (): Promise<string[]> => {
  try {
    const graph = await context.msGraphClientFactory.getClient("3");
    const memberOf = await graph.api("/me/memberOf").select("id").get();
    return (memberOf.value ?? []).map((g: { id?: string }) => g.id ?? "").filter(Boolean);
  } catch {
    return [];
  }
};

const loadIsAdmin = async (): Promise<boolean> => {
  try {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
      SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } });
    if (!res.ok) return false;
    const d = await res.json();
    return d.IsSiteAdmin === true;
  } catch { return false; }
};
```

- [ ] **Step 3: Add a term ancestor walk** near the Term Store helpers (after `loadTermChildren`, ~line 209):

```ts
// Resolve the full ancestor chain [top ... leaf] for a unit term, as {label,id}.
const loadTermPath = async (termSetId: string, leafTermId: string): Promise<TermOption[]> => {
  const chain: TermOption[] = [];
  let currentId: string | null = leafTermId;
  while (currentId) {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${currentId}?$select=id,labels&$expand=parent($select=id)`,
      SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } });
    if (!res.ok) break;
    const t = await res.json();
    chain.unshift({ id: t.id, label: t.labels?.[0]?.name ?? "" });
    currentId = t.parent?.id ?? null;
  }
  return chain;
};
```

- [ ] **Step 4: Replace the single-level cascade state.** Delete these state declarations (lines ~134-148): `detectedDept`, `allDepts`, `adminSelectedDept`, `subTeamChoices`, `subTeamValue`, `modeParents`. Keep `deptLoading`, `isAdmin`, `modes`, `uploadMode`. Add:

```ts
const [levelChoices, setLevelChoices] = useState<TermOption[][]>([]); // options for each level
const [levelValues, setLevelValues] = useState<string[]>([]);         // selected term id per level
const [userPaths, setUserPaths] = useState<UserPath[]>([]);
```

Also delete the now-unused `SubTeamNode` type (lines ~54-57) and the `loadDescendants`/`loadSubTeamTree` functions (lines ~275-318).

- [ ] **Step 5: Add cascade builders** (place after `loadTermPath`):

```ts
const activeMode = (): UploadMode | undefined => modes.find((m) => m.key === uploadMode);

// Load level-0 options (top terms of the mode's set); reset deeper levels.
const initCascade = async (mode: UploadMode): Promise<void> => {
  const tops = await loadTermSet(mode.termSetGuid).catch(() => [] as TermOption[]);
  setLevelChoices([tops]);
  setLevelValues([]);
};

// When level `idx` changes to `termId`, load level idx+1 options and truncate below.
const onLevelChange = async (mode: UploadMode, idx: number, termId: string): Promise<void> => {
  const values = levelValues.slice(0, idx);
  values[idx] = termId;
  setLevelValues(values);
  const choices = levelChoices.slice(0, idx + 1);
  if (idx + 1 < mode.levels.length && termId) {
    const kids = await loadTermChildren(mode.termSetGuid, termId).catch(() => [] as TermOption[]);
    choices[idx + 1] = kids;
  }
  setLevelChoices(choices);
};

// Pre-fill all levels from a detected UserPath (walk the term ancestry).
const prefillFromPath = async (mode: UploadMode, path: UserPath): Promise<void> => {
  const chain = await loadTermPath(mode.termSetGuid, path.unitTermGuid); // [top..leaf]
  const choices: TermOption[][] = [];
  const values: string[] = [];
  choices[0] = await loadTermSet(mode.termSetGuid).catch(() => [] as TermOption[]);
  for (let i = 0; i < chain.length && i < mode.levels.length; i++) {
    values[i] = chain[i].id;
    if (i + 1 < mode.levels.length) {
      choices[i + 1] = await loadTermChildren(mode.termSetGuid, chain[i].id).catch(() => [] as TermOption[]);
    }
  }
  setLevelChoices(choices);
  setLevelValues(values);
};
```

Proceed to C4.

### Task C4: Rewrite `init`, `switchMode`, upload folder routing

**Files:** Modify `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Rewrite the `useEffect` init** (lines ~395-471) to:

```ts
useEffect(() => {
  const init = async (): Promise<void> => {
    const [loadedSettings, loadedModes, groupMap, userGroupIds, admin] = await Promise.all([
      loadSettings().catch(() => DEFAULT_SETTINGS),
      loadModes().catch(() => DEFAULT_MODES),
      loadGroupMap().catch(() => [] as GroupMapRow[]),
      loadUserGroupIds(),
      loadIsAdmin(),
    ]);
    setSettings(loadedSettings);
    setModes(loadedModes);
    setIsAdmin(admin);

    const paths = matchUserPaths(groupMap, userGroupIds);
    setUserPaths(paths);

    const [docTypes, years, confs, vendors] = await Promise.all([
      loadTermSet(loadedSettings.termSets.documentType).catch(() => [] as TermOption[]),
      loadTermSet(loadedSettings.termSets.yearPeriod).catch(() => [] as TermOption[]),
      loadTermSet(loadedSettings.termSets.confidentiality).catch(() => [] as TermOption[]),
      loadTermSet(loadedSettings.termSets.vendor).catch(() => [] as TermOption[]),
    ]);
    setOptions({ documentType: docTypes, yearPeriod: years, confidentiality: confs, vendor: vendors });

    // Default toggle = BusinessSegment; default mode = first matched segment path, else first BS mode.
    const bsModes = loadedModes.filter((m) => m.side === "BusinessSegment");
    const firstPath = paths.find((p) => bsModes.some((m) => m.termSetGuid === p.segment));
    const defaultMode =
      (firstPath && bsModes.find((m) => m.termSetGuid === firstPath.segment)) ?? bsModes[0] ?? loadedModes[0];
    if (defaultMode) {
      setUploadMode(defaultMode.key);
      if (firstPath) await prefillFromPath(defaultMode, firstPath);
      else await initCascade(defaultMode);
    }
    setDeptLoading(false);
  };
  init().catch((err) => {
    console.error("Form init failed:", err);
    showToast("Could not load form data. Please refresh the page.", "error");
    setDeptLoading(false);
  });
}, []);
```

- [ ] **Step 2: Rewrite `switchMode`** (lines ~475-495) and delete `handleAdminDeptChange` (lines ~514-556):

```ts
const switchMode = (modeKey: string): void => {
  setUploadMode(modeKey);
  setStatus("");
  const mode = modes.find((m) => m.key === modeKey);
  if (!mode) return;
  const path = userPaths.find((p) => p.segment === mode.termSetGuid);
  if (path) prefillFromPath(mode, path).catch(() => initCascade(mode));
  else initCascade(mode).catch(() => { setLevelChoices([]); setLevelValues([]); });
};
```

- [ ] **Step 3: Update `resetForm`** (lines ~502-512) — replace `setSubTeamValue("")` with cascade reset:

```ts
const resetForm = (): void => {
  setFile(undefined);
  setDocName("");
  setDocumentType("");
  setLevelValues([]);
  setYearPeriod("");
  setDocumentDate("");
  setConfidentiality("");
  setVendor("");
  if (fileRef.current) fileRef.current.value = "";
};
```

- [ ] **Step 4: Rewrite the folder-resolution block in `handleUpload`.** Replace from `const selectedChoice = subTeamChoices.find(...)` (line ~593) down to `let uploadedServerRelativeUrl = "";` (line ~628) with:

```ts
const mode = activeMode();
if (!mode || mode.levels.length === 0) { showToast("No upload mode configured.", "error"); return; }
// The leaf level's selected term is the permissioned Unit folder.
const leafIdx = mode.levels.length - 1;
const leafTerm = (levelChoices[leafIdx] ?? []).find((o) => o.id === levelValues[leafIdx]);
if (!leafTerm) { showToast("Please choose all folder levels before uploading.", "error"); return; }

setBusy(true);
setStatus("Locating destination folder…");

const mapping = await lookupFolderMapping(context.spHttpClient, siteUrl, leafTerm.id)
  .catch((e: unknown) => { console.error("Folder map lookup error:", e); return null; });
if (!mapping || !mapping.folderUniqueId) {
  showToast(`This folder hasn't been mapped yet. Ask an administrator to run the reconciliation tool. (term ${leafTerm.label})`, "error");
  setStatus(""); setBusy(false); return;
}

// Rename-proof: resolve the Unit folder's CURRENT path from its UniqueId, then
// ensure-create the Year and Document Type subfolders under it (they inherit its ACL).
const unitSru = await resolveFolderServerUrl(context.spHttpClient, siteUrl, mapping.folderUniqueId);
if (!unitSru) {
  showToast("The mapped unit folder no longer exists. Ask an administrator to re-run reconciliation.", "error");
  setStatus(""); setBusy(false); return;
}
const yearLabel = sanitizeFolderSegment(options.yearPeriod.find((o) => o.id === yearPeriod)?.label ?? "");
const docTypeLabel = sanitizeFolderSegment(options.documentType.find((o) => o.id === documentType)?.label ?? "");
if (!yearLabel || !docTypeLabel) { showToast("Year and Document Type are required.", "error"); setStatus(""); setBusy(false); return; }

setStatus("Preparing destination folders…");
const yearFolder = await ensureFolder(context.spHttpClient, siteUrl, unitSru, yearLabel);
if (!yearFolder) { showToast(`Could not create the "${yearLabel}" folder.`, "error"); setStatus(""); setBusy(false); return; }
const destFolder = await ensureFolder(context.spHttpClient, siteUrl, yearFolder.serverRelativeUrl, docTypeLabel);
if (!destFolder) { showToast(`Could not create the "${docTypeLabel}" folder.`, "error"); setStatus(""); setBusy(false); return; }

const folderId = destFolder.uniqueId; // upload target — a fresh, unit-scoped folder
let uploadedServerRelativeUrl = "";
```

The existing duplicate-check + `Files/Add` + item-retrieve code below already targets `folderId` via `GetFolderById(guid'${folderId}')` — leave that mechanism unchanged. Proceed to C5.

### Task C5: Metadata payload with per-level columns; column map

**Files:** Modify `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Replace the `FIELDS` constant** (lines ~11-23) with the trimmed tail-only fields + a `LEVEL_COLUMNS` map. **Paste the real internal names captured in Task A1 Step 3** in place of each `REPLACE_*`. The keys of `LEVEL_COLUMNS` MUST match the `column` tokens used in DMS Config Levels (Task A3):

```ts
const FIELDS = {
  documentType: "Department_x0020_Type",
  yearPeriod: "Year_x002f_Period",
  documentDate: "DocumentDate",
  confidentiality: "Confidentiality_x0020_Level",
  vendor: "Vendor",
  details: "_ExtendedDescription",
};

// logical Levels `column` key -> real Staging internal name (from Task A1 /fields read)
const LEVEL_COLUMNS: Record<string, string> = {
  BusinessSegment: "REPLACE_BusinessSegment_Internal",
  Department: "REPLACE_Department_Internal",
  Unit: "REPLACE_Unit_Internal",
  Region: "REPLACE_Region_Internal",
  EstateMill: "REPLACE_EstateMill_Internal",
  Refinery: "REPLACE_Refinery_Internal",
  ITOperatingUnit: "REPLACE_ITOperatingUnit_Internal",
  ProjectName: "REPLACE_ProjectName_Internal",
};
```

- [ ] **Step 2: Rebuild the `formValues` array** in `handleUpload`. Replace the old `formValues` block + the `if (selectedSubTeam)` block (lines ~705-731) with:

```ts
const mode2 = activeMode();
const selections = (mode2?.levels ?? []).map((lvl, i) => {
  const opt = (levelChoices[i] ?? []).find((o) => o.id === levelValues[i]);
  return { column: lvl.column, label: opt?.label ?? "", id: opt?.id ?? "" };
});
// BusinessSegment column = the mode's segment label (for BS modes; harmless for Project).
if (mode2) selections.unshift({ column: "BusinessSegment", label: mode2.label, id: mode2.termSetGuid });

const formValues: Array<{ FieldName: string; FieldValue: string }> = [
  { FieldName: FIELDS.documentType, FieldValue: toTaxValue(options.documentType, documentType) },
  { FieldName: FIELDS.yearPeriod, FieldValue: toTaxValue(options.yearPeriod, yearPeriod) },
  { FieldName: FIELDS.confidentiality, FieldValue: toTaxValue(options.confidentiality, confidentiality) },
  { FieldName: FIELDS.documentDate, FieldValue: toSpDate(documentDate) },
  ...buildLevelFormValues(LEVEL_COLUMNS, selections),
];
if (vendor) formValues.push({ FieldName: FIELDS.vendor, FieldValue: toTaxValue(options.vendor, vendor) });
```

- [ ] **Step 3: Replace the upload validation block** (lines ~561-576) with a generic level check:

```ts
const missing: string[] = [];
if (!file) missing.push("File");
if (!documentType) missing.push("Document Type");
const m = activeMode();
(m?.levels ?? []).forEach((lvl, i) => { if (!levelValues[i]) missing.push(lvl.label); });
if (!yearPeriod) missing.push("Year / Period");
if (!documentDate) missing.push("Document Date");
if (!confidentiality) missing.push("Confidentiality Level");
if (missing.length > 0) { showToast(`Please complete: ${missing.join(", ")}.`, "error"); return; }
if (!file) return;
```

Proceed to C6.

### Task C6: Render — toggle + generic level dropdowns; clean up + verify

**Files:** Modify `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Replace the dept badge / admin block** (lines ~963-1000) with:

```tsx
{deptLoading ? (
  <p className="dms-dept-loading">Loading your access&hellip;</p>
) : userPaths.length === 0 && !isAdmin ? (
  <div className="dms-dept-error">
    Your account isn't mapped to any unit. Contact your administrator before uploading.
  </div>
) : null}
```

- [ ] **Step 2: Replace the mode radio group** (lines ~1002-1017) with a Business Segment | Project toggle:

```tsx
<div className="dms-radio-group">
  <p>Upload into:</p>
  {(["BusinessSegment", "Project"] as const).map((side) => {
    const first = modes.find((m) => m.side === side);
    if (!first) return null;
    const active = activeMode()?.side === side;
    return (
      <label key={side}>
        <input type="radio" name="sideToggle" checked={active} onChange={() => switchMode(first.key)} />
        {side === "BusinessSegment" ? "Business Segment" : "Project"}
      </label>
    );
  })}
</div>
```

- [ ] **Step 3: Add a segment picker** shown only when the active side has >1 mode (Business Segment has 4). Place it just above the `dms-grid` div (~line 1019):

```tsx
{(() => {
  const side = activeMode()?.side;
  const sideModes = modes.filter((m) => m.side === side);
  if (sideModes.length <= 1) return null;
  return (
    <label className="dms-field">
      <span>Segment</span>
      <select value={uploadMode} onChange={(e) => switchMode(e.target.value)}>
        {sideModes.map((m) => <option key={m.key} value={m.key}>{m.label}</option>)}
      </select>
    </label>
  );
})()}
```

- [ ] **Step 4: Replace the single Sub-team dropdown** (lines ~1053-1063) with one dropdown per level:

```tsx
{(activeMode()?.levels ?? []).map((lvl, i) =>
  renderSelect(
    lvl.label,
    true,
    levelValues[i] ?? "",
    (v) => { const m = activeMode(); if (m) onLevelChange(m, i, v); },
    levelChoices[i] ?? [],
    deptLoading || (i > 0 && !levelValues[i - 1]),
    "--",
    true,
  )
)}
```

- [ ] **Step 5: Fix the Upload button guard** (line ~1097) — replace `disabled={busy || deptLoading || !detectedDept}` with:

```tsx
disabled={busy || deptLoading}
```

- [ ] **Step 6: Remove the two `[DMS DEBUG]` console blocks** (lines ~740-744 and ~759-761).

- [ ] **Step 7: Build until green.**

Run: `npm run build`
Expected: build succeeds. Resolve any lingering references to deleted symbols (`detectedDept`, `subTeamChoices`, `subTeamValue`, `allDepts`, `adminSelectedDept`, `modeParents`, `SubTeamNode`, `handleAdminDeptChange`, `loadDescendants`, `loadSubTeamTree`) until the compiler is clean.

- [ ] **Step 8: Manual workbench verification.**

Run: `npm run start`
Open the workbench (URL in CLAUDE.md). Sign in as a user whose group is in DMS Group Map (Task A4).
Expected:
1. Business Segment toggle selected; the user's segment + its levels pre-filled (editable), e.g. Group Head Office → Department=Group Finance → Unit=Tax.
2. Changing a level reloads the level below.
3. Switching to Project shows the Projects cascade (Project Name → Department → Unit).
4. Fill Year + Document Type + Document Date + Confidentiality, pick a file, Upload.
5. File lands in `Staging/Group Head Office/Group Finance/Tax/<Year>/<Document Type>/` and the label + `_Tid` columns are populated (check the library view).

- [ ] **Step 9: Commit the whole Form.tsx change.**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: multi-segment data-driven form (5 modes, group map, N-level cascade, Year/DocType folders)"
```

---

## Phase D — Reconciliation, fallbacks, ship

### Task D1: Reconciliation 5-mode fallback

**Files:** Modify `src/webparts/reconciliation/components/Reconciliation.tsx`

- [ ] **Step 1: Replace `DEFAULT_MODES`** (lines ~29-32) with the 5 segments (paste GUIDs from A2). Logic that walks the term tree to the leaf is unchanged:

```ts
const DEFAULT_MODES: Mode[] = [
  { key: "gho", termSetGuid: "REPLACE-GHO", stagingFolder: "Group Head Office" },
  { key: "upstream", termSetGuid: "REPLACE-UP", stagingFolder: "Group Upstream Operations" },
  { key: "sdgi", termSetGuid: "REPLACE-SDGI", stagingFolder: "Group SDGI Operations" },
  { key: "it", termSetGuid: "REPLACE-IT", stagingFolder: "Group Innovation & Technology" },
  { key: "projects", termSetGuid: "REPLACE-PROJ", stagingFolder: "Group-led Projects" },
];
```

- [ ] **Step 2: Build.**

Run: `npm run build`
Expected: succeeds.

- [ ] **Step 3: Commit.**

```bash
git add src/webparts/reconciliation/components/Reconciliation.tsx
git commit -m "feat: reconciliation fallback modes for 5 term sets"
```

### Task D2: Drop the single department term set; update CLAUDE.md

**Files:** Modify `src/webparts/form/components/Form.tsx`, `CLAUDE.md`

- [ ] **Step 1: In `DmsSettings` type + `DEFAULT_SETTINGS`** (Form.tsx ~105-127) remove the `department` term-set field entirely (no longer a single set). Keep `documentType`, `yearPeriod`, `confidentiality`, `vendor`.

- [ ] **Step 2: In `loadSettings`** (~255-267) remove the `department: get("termSet_department") ?? ...` line so it no longer reads the retired setting.

- [ ] **Step 3: Build.**

Run: `npm run build`
Expected: succeeds (confirm no remaining reference to `settings.termSets.department`).

- [ ] **Step 4: Update CLAUDE.md** — replace the single `department` line in "Term Set GUIDs" with the 5 new segment sets + GUIDs; in "Staging Library — Column Internal Names" add the 16 new columns (label + `_Tid`); add notes for the `DMS Group Map` list and the on-demand `Year`/`Document Type` folders.

- [ ] **Step 5: Commit.**

```bash
git add src/webparts/form/components/Form.tsx CLAUDE.md
git commit -m "chore: drop single department term set; document 5-segment model"
```

### Task D3: Version bump + package

**Files:** Modify `config/package-solution.json`

- [ ] **Step 1: Bump `solution.version`** `1.0.2.0` → `1.0.3.0`.

- [ ] **Step 2: Build the production package.**

Run: `npm run build`
Expected: `solution/sd-gatrie.sppkg` produced, no errors.

- [ ] **Step 3: Commit.**

```bash
git add config/package-solution.json
git commit -m "chore: bump solution to 1.0.3.0"
```

---

## Self-review notes (spec coverage)

- Spec §1 (mode Levels) → A3, C2, C4. Spec §2 (Group Map) → A4, C3. Spec §3 (flow: toggle, detect, cascade, upload) → C3, C4, C6. Spec §4 (16 text columns, label+_Tid) → A1, B5, C5. Spec §5 (code touch-points) → all Phase C/D. Spec §6 (migration order) → Phase A ordering. Folder-depth addendum (Year/DocType on-demand, inherit ACL) → C1, C4.
- **Known non-green window:** `Form.tsx` does not compile between C2 and C6 (interdependent edits). Tasks C2–C6 are committed once, at C6 Step 9, after the build is green — deliberate, and called out at the top of Phase C and in each task.
- **Intentional placeholders (data, not plan gaps):** every `REPLACE-*` / `REPLACE_*_Internal` token is a real GUID/internal-name the operator captures in Phase A and pastes in Phase C/D. No logic is left unspecified.

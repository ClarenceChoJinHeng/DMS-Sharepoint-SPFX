# Native SharePoint Groups Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace Entra security groups with native SharePoint site groups across the DMS — group search, group creation, member management, membership detection, and folder ACL assignment — removing every Graph group call.

**Architecture:** One new REST module (`src/shared/spGroups.ts`) is the single source of SP-group access for all web parts. `DMS Group Map.GroupId` changes meaning from Entra Object-ID GUID to SP group integer Id as a string (`"27"`); that integer doubles as the folder role-assignment principal id, so the `ensureuser` claim dance disappears. `GroupMapBuilder` gains inline group create + a member editor. Hard cutover — no dual-source support.

**Tech Stack:** SPFx 1.23.0, React, `SPHttpClient` (SharePoint REST only — no Graph), Heft + Jest (`npx heft test`), Node 22 (`nvm use 22`).

**Spec:** `docs/superpowers/specs/2026-07-23-native-sharepoint-groups-design.md`

**Working-tree caution:** the repo has unrelated uncommitted WIP (`src/webparts/bulkUpload/` is untracked, `FolderManager.tsx` has uncommitted bulk-upload-era changes, `config/*.json` modified). Every commit in this plan must `git add` **only the files named in that task** — never `git add -A`/`git add .`.

**Environment note:** line numbers below were verified 2026-07-24; if a hunk doesn't match, locate by the quoted code, not the number.

---

## File structure

| File | Role in this plan |
|---|---|
| `src/shared/groupMapModel.ts` (+`.test.ts`) | add pure `suggestGroupName` (Task 1) |
| `src/shared/spGroups.ts` (+`.test.ts`) — **new** | all SP-group REST + pure `filterDmsGroups` (Task 2) |
| `src/webparts/form/components/Form.tsx` | membership via `currentuser/groups` (Task 3) |
| `src/webparts/bulkUpload/components/BulkUpload.tsx` | same (Task 4) |
| `src/webparts/folderManager/components/FolderManager.tsx` | SP-group search; integer principal id (Task 5) |
| `src/webparts/onboarding/components/Onboarding.tsx` | same + UI copy (Task 6) |
| `src/webparts/folderManager/components/GroupMapBuilder.tsx` | SP search, inline create, permission gate (Task 7); member editor (Task 8) |
| `config/package-solution.json` | drop 3 Graph scopes, bump version (Task 9) |

---

### Task 1: `suggestGroupName` in groupMapModel (TDD)

**Files:**
- Modify: `src/shared/groupMapModel.ts` (append after `validateDraft`, line 76)
- Test: `src/shared/groupMapModel.test.ts` (append at end)

- [ ] **Step 1: Write the failing tests**

Append to `src/shared/groupMapModel.test.ts` (add `suggestGroupName` to the existing import list at the top of the file):

```ts
describe("suggestGroupName", () => {
  it("joins DMS + segment + tier labels + role suffix", () => {
    expect(
      suggestGroupName("Group Head Office", ["Group Finance", "Corporate Reporting"], "UPL"),
    ).toBe("DMS_Group Head Office_Group Finance_Corporate Reporting_UPL");
  });

  it("APR suffix works the same way", () => {
    expect(suggestGroupName("Minamas Head Office", ["Finance"], "APR")).toBe(
      "DMS_Minamas Head Office_Finance_APR",
    );
  });

  it("MEMBER gets no suffix (base/viewer group)", () => {
    expect(suggestGroupName("Group Head Office", ["Group Finance"], "MEMBER")).toBe(
      "DMS_Group Head Office_Group Finance",
    );
  });

  it("GLOBAL is a fixed name regardless of labels", () => {
    expect(suggestGroupName("anything", ["x", "y"], "GLOBAL")).toBe("DMS_GLOBAL");
  });

  it("skips empty labels and trims the rest", () => {
    expect(suggestGroupName(" GHO ", ["", "  Unit A "], "APR")).toBe("DMS_GHO_Unit A_APR");
  });

  it("empty role and empty labels degrade gracefully", () => {
    expect(suggestGroupName("GHO", [], "")).toBe("DMS_GHO");
    expect(suggestGroupName("", [], "")).toBe("DMS");
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx heft test`
Expected: FAIL — `suggestGroupName` is not exported.

- [ ] **Step 3: Implement**

Append to `src/shared/groupMapModel.ts`:

```ts
/**
 * Suggest a site-group title from the builder's current selections, e.g.
 * "DMS_Group Head Office_Group Finance_UPL". Purely a SUGGESTION — the admin
 * edits it freely before create (no abbreviation algorithm by design; see the
 * native-sharepoint-groups spec). GLOBAL is a fixed name; MEMBER has no suffix.
 */
export function suggestGroupName(
  segmentLabel: string,
  tierLabels: string[],
  role: GroupMapRole | "",
): string {
  if (role === "GLOBAL") return "DMS_GLOBAL";
  const parts = ["DMS", norm(segmentLabel), ...tierLabels.map(norm)].filter(Boolean);
  const suffix = role === "UPL" || role === "APR" ? `_${role}` : "";
  return parts.join("_") + suffix;
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx heft test`
Expected: PASS (all existing groupMapModel/formModel/shareGuard/pathEncoding/siteMap tests still green).

- [ ] **Step 5: Commit**

```bash
git add src/shared/groupMapModel.ts src/shared/groupMapModel.test.ts
git commit -m "feat: suggestGroupName for SP-group inline create"
```

---

### Task 2: `src/shared/spGroups.ts` — SP group REST module (pure part TDD)

**Files:**
- Create: `src/shared/spGroups.ts`
- Test: `src/shared/spGroups.test.ts` (pure `filterDmsGroups` only — REST functions are verified live in Task 11, per the spec)

- [ ] **Step 1: Write the failing tests**

Create `src/shared/spGroups.test.ts`:

```ts
import { filterDmsGroups, SpGroup } from "./spGroups";

const g = (id: number, title: string): SpGroup => ({ id, title });

describe("filterDmsGroups", () => {
  const all = [
    g(3, "SPFX Sandbox Owners"),
    g(4, "SPFX Sandbox Members"),
    g(5, "SPFX Sandbox Visitors"),
    g(27, "DMS_GHO_Group Finance_UPL"),
    g(28, "DMS_GHO_Group Finance_APR"),
    g(29, "dms_gho_group finance"),
  ];

  it("keeps only DMS_-prefixed groups (built-ins excluded)", () => {
    expect(filterDmsGroups(all, "").map((x) => x.id)).toEqual([27, 28, 29]);
  });

  it("prefix match is case-insensitive", () => {
    expect(filterDmsGroups(all, "").some((x) => x.id === 29)).toBe(true);
  });

  it("query filters by substring, case-insensitive", () => {
    expect(filterDmsGroups(all, "_upl").map((x) => x.id)).toEqual([27]);
    expect(filterDmsGroups(all, "FINANCE").map((x) => x.id)).toEqual([27, 28, 29]);
  });

  it("no match returns empty, never throws", () => {
    expect(filterDmsGroups(all, "zzz")).toEqual([]);
    expect(filterDmsGroups([], "x")).toEqual([]);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx heft test`
Expected: FAIL — cannot find module `./spGroups`.

- [ ] **Step 3: Implement the module**

Create `src/shared/spGroups.ts` (complete file):

```ts
// Native SharePoint site-group access — the ONLY module that talks to
// /_api/web/sitegroups. Replaces Graph group search/membership everywhere
// (see docs/superpowers/specs/2026-07-23-native-sharepoint-groups-design.md).
// Site-collection scoped by nature: these groups exist only on this site.
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export type SpGroup = { id: number; title: string };
export type SpGroupMember = { id: number; title: string; email: string; loginName: string };
export type PersonPick = { loginName: string; displayName: string; email: string };

const GET_HEADERS = { Accept: "application/json;odata=nometadata" };
const POST_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
};

/** Thrown by createSiteGroup when the title is already taken on the site collection. */
export const DUPLICATE_GROUP = "DUPLICATE_GROUP";

const fail = async (label: string, res: SPHttpClientResponse): Promise<never> => {
  const body = await res.text().catch(() => "");
  throw new Error(`${label} HTTP ${res.status} ${body.slice(0, 200)}`);
};

/** Pure: DMS_-prefixed site groups matching q (both case-insensitive). Exported for tests. */
export function filterDmsGroups(all: SpGroup[], q: string): SpGroup[] {
  const query = (q ?? "").trim().toLowerCase();
  return all
    .filter((g) => g.title.toUpperCase().indexOf("DMS_") === 0)
    .filter((g) => !query || g.title.toLowerCase().indexOf(query) !== -1);
}

export async function fetchAllSiteGroups(sp: SPHttpClient, siteUrl: string): Promise<SpGroup[]> {
  const res = await sp.get(
    `${siteUrl}/_api/web/sitegroups?$select=Id,Title&$top=500`,
    SPHttpClient.configurations.v1,
    { headers: GET_HEADERS },
  );
  if (!res.ok) return fail("sitegroups", res);
  const data = await res.json();
  return ((data.value ?? []) as Array<{ Id: number; Title: string }>).map((g) => ({
    id: g.Id,
    title: g.Title,
  }));
}

export async function searchSiteGroups(
  sp: SPHttpClient,
  siteUrl: string,
  q: string,
): Promise<SpGroup[]> {
  return filterDmsGroups(await fetchAllSiteGroups(sp, siteUrl), q);
}

/**
 * Create a site group with NO permissions (Reconciliation grants folder access
 * later, so creating a group is a safe, reversible act). Throws Error(DUPLICATE_GROUP)
 * if the title already exists in the site collection.
 */
export async function createSiteGroup(
  sp: SPHttpClient,
  siteUrl: string,
  title: string,
): Promise<SpGroup> {
  const res = await sp.post(`${siteUrl}/_api/web/sitegroups`, SPHttpClient.configurations.v1, {
    headers: POST_HEADERS,
    body: JSON.stringify({ Title: title.trim() }),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    if (body.toLowerCase().indexOf("already in use") !== -1) throw new Error(DUPLICATE_GROUP);
    throw new Error(`create group HTTP ${res.status} ${body.slice(0, 200)}`);
  }
  const data = await res.json();
  return { id: data.Id as number, title: data.Title as string };
}

export async function getGroupMembers(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
): Promise<SpGroupMember[]> {
  const res = await sp.get(
    `${siteUrl}/_api/web/sitegroups(${groupId})/users?$select=Id,Title,Email,LoginName`,
    SPHttpClient.configurations.v1,
    { headers: GET_HEADERS },
  );
  if (!res.ok) return fail("group members", res);
  const data = await res.json();
  return ((data.value ?? []) as Array<{ Id: number; Title: string; Email?: string; LoginName: string }>).map(
    (u) => ({ id: u.Id, title: u.Title, email: u.Email ?? "", loginName: u.LoginName }),
  );
}

/** Add by login name (claim). SharePoint ensures the user on the site automatically; guests work. */
export async function addGroupMember(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
  loginName: string,
): Promise<void> {
  const res = await sp.post(
    `${siteUrl}/_api/web/sitegroups(${groupId})/users`,
    SPHttpClient.configurations.v1,
    { headers: POST_HEADERS, body: JSON.stringify({ LoginName: loginName }) },
  );
  if (!res.ok) return fail("add member", res);
}

export async function removeGroupMember(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
  userId: number,
): Promise<void> {
  const res = await sp.post(
    `${siteUrl}/_api/web/sitegroups(${groupId})/users/removebyid(${userId})`,
    SPHttpClient.configurations.v1,
    { headers: POST_HEADERS },
  );
  if (!res.ok) return fail("remove member", res);
}

/** Tenant-wide people search via SharePoint's own picker service — no Graph. */
export async function searchTenantPeople(
  sp: SPHttpClient,
  siteUrl: string,
  q: string,
): Promise<PersonPick[]> {
  const res = await sp.post(
    `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`,
    SPHttpClient.configurations.v1,
    {
      headers: POST_HEADERS,
      body: JSON.stringify({
        queryParams: {
          QueryString: q,
          MaximumEntitySuggestions: 10,
          AllowEmailAddresses: true,
          AllowOnlyEmailAddresses: false,
          PrincipalType: 1, // users only
          PrincipalSource: 15, // all sources
        },
      }),
    },
  );
  if (!res.ok) return fail("people search", res);
  const data = await res.json();
  // The endpoint returns its result as a JSON *string*.
  const raw = (data.value ?? data.ClientPeoplePickerSearchUser ?? "[]") as string;
  const entries = JSON.parse(raw) as Array<{
    Key: string;
    DisplayText: string;
    EntityData?: { Email?: string };
  }>;
  return entries.map((e) => ({
    loginName: e.Key,
    displayName: e.DisplayText,
    email: e.EntityData?.Email ?? "",
  }));
}
```

> If `ClientPeoplePickerSearchUser` returns HTTP 400 in the sandbox, add `"odata-version": "3.0"` to `POST_HEADERS` for that one call — some tenants require it on ApplicationPages endpoints.

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx heft test`
Expected: PASS (filterDmsGroups suite green; TypeScript compiles the module).

- [ ] **Step 5: Commit**

```bash
git add src/shared/spGroups.ts src/shared/spGroups.test.ts
git commit -m "feat: spGroups.ts — native SP site-group REST module"
```

---

### Task 3: Form.tsx — membership via `currentuser/groups`

**Files:**
- Modify: `src/webparts/form/components/Form.tsx:430-444` (`loadUserGroupIds`)

- [ ] **Step 1: Replace the Graph call**

Replace this block (lines 430–444):

```ts
  /* ---------- Identity: group ids + admin --------------------------------- */

  // Detect the user's M365 (Entra) group Object IDs via MS Graph /me/memberOf.
  // These are matched against DMS Group Map GroupId (Object ID) — names are cosmetic.
  const loadUserGroupIds = async (): Promise<string[]> => {
    try {
      const graph = await context.msGraphClientFactory.getClient("3");
      const memberOf = await graph.api("/me/memberOf").select("id").get();
      return (memberOf.value ?? [])
        .map((g: { id?: string }) => g.id ?? "")
        .filter(Boolean);
    } catch {
      return [];
    }
  };
```

with:

```ts
  /* ---------- Identity: group ids + admin --------------------------------- */

  // Detect the user's native SharePoint site-group Ids (direct membership only).
  // Matched as strings against DMS Group Map GroupId (SP group integer Id, e.g. "27").
  // No Graph, no admin consent. Nested Entra groups inside an SP group are NOT
  // returned — the DMS model adds users to site groups directly by design.
  const loadUserGroupIds = async (): Promise<string[]> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser/groups?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return [];
      const d = await res.json();
      return ((d.value ?? []) as Array<{ Id?: number }>)
        .map((g) => (g.Id != null ? String(g.Id) : ""))
        .filter(Boolean);
    } catch {
      return [];
    }
  };
```

`collectMembership` in `formModel.ts` is untouched — it compares strings and does not care whether they are GUIDs or integers.

- [ ] **Step 2: Verify no Graph reference remains in Form.tsx**

Run: `grep -n "msGraphClientFactory" src/webparts/form/components/Form.tsx`
Expected: no output. (`Form-Copy.txt` still matches repo-wide — it is a non-compiled backup, leave it.)

- [ ] **Step 3: Run tests + typecheck**

Run: `npx heft test`
Expected: PASS, no compile errors.

- [ ] **Step 4: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: Form membership detection via SP currentuser/groups (no Graph)"
```

---

### Task 4: BulkUpload.tsx — identical membership change

**Files:**
- Modify: `src/webparts/bulkUpload/components/BulkUpload.tsx:446-456` (`loadUserGroupIds`)

- [ ] **Step 1: Replace the Graph call**

Replace (lines 446–456):

```ts
  const loadUserGroupIds = async (): Promise<string[]> => {
    try {
      const graph = await context.msGraphClientFactory.getClient("3");
      const memberOf = await graph.api("/me/memberOf").select("id").get();
      return (memberOf.value ?? [])
        .map((g: { id?: string }) => g.id ?? "")
        .filter(Boolean);
    } catch {
      return [];
    }
  };
```

with (repeated in full so this task stands alone):

```ts
  // Native SP site-group Ids as strings (matches DMS Group Map GroupId, e.g. "27").
  const loadUserGroupIds = async (): Promise<string[]> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser/groups?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return [];
      const d = await res.json();
      return ((d.value ?? []) as Array<{ Id?: number }>)
        .map((g) => (g.Id != null ? String(g.Id) : ""))
        .filter(Boolean);
    } catch {
      return [];
    }
  };
```

- [ ] **Step 2: Verify**

Run: `grep -n "msGraphClientFactory" src/webparts/bulkUpload/components/BulkUpload.tsx`
Expected: no output.

Run: `npx heft test`
Expected: PASS.

- [ ] **Step 3: Commit**

BulkUpload.tsx is untracked WIP (`src/webparts/bulkUpload/` is a new directory with other in-progress changes). **Do not commit it in this plan** — record in the tracker that the membership change rides along when the user commits the bulk-upload web part. Skip `git add` here.

---

### Task 5: FolderManager.tsx — SP search + integer principal id

**Files:**
- Modify: `src/webparts/folderManager/components/FolderManager.tsx` — lines 3 (import), 101 (`GroupPick`), 228 (mail render), 386-398 (`ensureGroupPrincipal`), 409-417 (`searchGroups`), 655-663 (`ensurePrincipal`), 1004 (reconcile principal)

- [ ] **Step 1: Swap the type, search, and principal resolution**

1. Line 3 — remove `MSGraphClientV3` from the import:
```ts
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
```
2. Add below the existing imports:
```ts
import { searchSiteGroups } from "../../../shared/spGroups";
```
3. Line 101 — `GroupPick` loses the Graph-only fields:
```ts
type GroupPick      = { id: string; displayName: string };
```
4. Line 228 — delete this line inside the GroupSearch dropdown item:
```ts
              {g.mail && <div style={{ fontSize: 11, color: "#888" }}>{g.mail}</div>}
```
5. Lines 386–398 — replace `ensureGroupPrincipal` entirely with a synchronous converter:
```ts
  // SP site group: the group's integer Id IS the role-assignment principal id.
  // No ensureuser, no federateddirectoryclaimprovider claim. A GUID here means a
  // legacy Entra row that must be recreated via the Group Map tab.
  const spGroupPrincipalId = (groupId: string): number => {
    const n = Number((groupId ?? "").trim());
    if (!(n > 0) || n % 1 !== 0) {
      throw new Error(
        `"${groupId}" is not a SharePoint site-group id — recreate this Group Map row with the Group Map tab`,
      );
    }
    return n;
  };
```
6. Lines 409–417 — replace `searchGroups`:
```ts
  const searchGroups = async (query: string): Promise<GroupPick[]> => {
    const groups = await searchSiteGroups(context.spHttpClient, siteUrl, query);
    return groups.map((g) => ({ id: String(g.id), displayName: g.title }));
  };
```
7. Lines 655–663 — the `ensurePrincipal` wrapper and its cache become a direct call. Replace:
```ts
    const principalCache = new Map<string, number>();
    const ensurePrincipal = async (group: GroupPick): Promise<number> => {
      const cached = principalCache.get(group.id);
      if (cached !== undefined) return cached;
      const pid = await ensureGroupPrincipal(group);
      principalCache.set(group.id, pid);
      return pid;
    };
```
with:
```ts
    const ensurePrincipal = async (group: GroupPick): Promise<number> =>
      spGroupPrincipalId(group.id);
```
(keeping it async avoids touching its call sites at line ~689.)
8. Line 1004 — replace:
```ts
                const pid = await ensureGroupPrincipal({ id: g.groupId, displayName: g.groupName, isUnified: false });
```
with:
```ts
                const pid = spGroupPrincipalId(g.groupId);
```

- [ ] **Step 2: Compile-sweep the leftovers**

Run: `grep -n "isUnified\|\.mail\|ensureGroupPrincipal\|MSGraphClientV3" src/webparts/folderManager/components/FolderManager.tsx`
Expected: no output. If the compiler or grep finds another `isUnified`/`mail` usage, delete that usage — the fields no longer exist.

Run: `npx heft test`
Expected: PASS, no compile errors.

- [ ] **Step 3: Commit**

FolderManager.tsx carries unrelated uncommitted WIP (~47 lines from the bulk-upload work plus the RECON_MODES head-office entries). **Default for the executor: skip the commit and record "pending commit with bulk-upload WIP"** — commit only if the user confirms the WIP should ride along.

---

### Task 6: Onboarding.tsx — same swaps + UI copy

**Files:**
- Modify: `src/webparts/onboarding/components/Onboarding.tsx` — lines 3, 13-16 (`GroupPick`), 86, 137, 206-234, 277-292, 353, 462, 502-510

- [ ] **Step 1: Apply the changes**

1. Line 3 — remove `MSGraphClientV3`:
```ts
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
```
and add:
```ts
import { searchSiteGroups } from "../../../shared/spGroups";
```
2. `GroupPick` (around lines 13–16) — remove `mail`/`isUnified`:
```ts
type GroupPick = { id: string; displayName: string };
```
3. Line 86 comment — change `/* ── Group search field (live Microsoft Graph search) ── */` to `/* ── Group search field (native SP site groups) ── */`.
4. Line 137 — delete the `{g.mail && …}` line in the dropdown render.
5. Lines 206–224 — replace `searchGroups` with:
```ts
  const searchGroups = async (query: string): Promise<GroupPick[]> => {
    const groups = await searchSiteGroups(context.spHttpClient, siteUrl, query);
    return groups.map((g) => ({ id: String(g.id), displayName: g.title }));
  };
```
6. Lines 226–234 — replace `groupExists` with:
```ts
  const groupExists = async (id: string): Promise<boolean> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/sitegroups(${Number(id)})?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    return res.ok;
  };
```
7. Lines 277–292 — replace `ensureGroupPrincipal` with the same `spGroupPrincipalId` converter shown in Task 5 Step 1.5 (identical code):
```ts
  const spGroupPrincipalId = (groupId: string): number => {
    const n = Number((groupId ?? "").trim());
    if (!(n > 0) || n % 1 !== 0) {
      throw new Error(
        `"${groupId}" is not a SharePoint site-group id — pick the group from the search box`,
      );
    }
    return n;
  };
```
8. Line 353 — replace `const pid = await ensureGroupPrincipal(g);` with `const pid = spGroupPrincipalId(g.id);`.
9. Line 462 — `title={a.group.mail ?? a.group.displayName}` → `title={a.group.displayName}`.
10. Lines 502–505 subtitle — replace "assign one or more M365 groups to each" with "assign one or more SharePoint site groups to each".
11. Lines 507–510 — replace the Graph notice with:
```tsx
      <div style={styles.note}>
        Groups here are <strong>native SharePoint site groups</strong> on this site (names
        starting <strong>DMS_</strong>). Create and manage them in the DMS Admin Tool's
        Group Map tab — no Microsoft Graph approval is involved.
      </div>
```

- [ ] **Step 2: Verify + test**

Run: `grep -n "msGraphClientFactory\|MSGraphClientV3\|isUnified\|\.mail" src/webparts/onboarding/components/Onboarding.tsx`
Expected: no output.

Run: `npx heft test`
Expected: PASS.

- [ ] **Step 3: Commit**

```bash
git add src/webparts/onboarding/components/Onboarding.tsx
git commit -m "feat: Onboarding uses native SP site groups (no Graph)"
```

---

### Task 7: GroupMapBuilder — SP search, inline create, permission gate

**Files:**
- Modify: `src/webparts/folderManager/components/GroupMapBuilder.tsx` — lines 4 (import), 83-95 (`searchGroups`), 68-74 (state), 175-178 (mount), 344-380 (render: intro + group field)

- [ ] **Step 1: Imports and search**

1. Line 4 — remove `MSGraphClientV3`:
```ts
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
```
2. Extend the `groupMapModel` import (lines 5–13) with `suggestGroupName`.
3. Add:
```ts
import {
  searchSiteGroups,
  createSiteGroup,
  getGroupMembers,
  addGroupMember,
  removeGroupMember,
  searchTenantPeople,
  DUPLICATE_GROUP,
  SpGroupMember,
  PersonPick,
} from "../../../shared/spGroups";
```
4. Replace `searchGroups` (lines 83–95) with:
```ts
  const searchGroups = async (q: string): Promise<GroupPick[]> => {
    const groups = await searchSiteGroups(context.spHttpClient, siteUrl, q);
    return groups.map((g) => ({ id: String(g.id), displayName: g.title }));
  };
```
`GroupPick` stays `{ id: string; displayName: string }` — the id is now the integer as a string.

- [ ] **Step 2: Permission gate state + check**

Add state next to the existing `busy` state (after line 57):
```ts
  const [canManage, setCanManage] = useState<boolean | undefined>(undefined); // undefined = still checking
```
Add the loader with the other data-access functions:
```ts
  // Create/add/remove need Full Control. Detect up front so the controls are
  // disabled with an explanation instead of failing with a 403 on click.
  const loadCanManage = async (): Promise<boolean> => {
    const meRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/currentuser?$select=Id,IsSiteAdmin`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!meRes.ok) return false;
    const me = await meRes.json();
    if (me.IsSiteAdmin === true) return true;
    const ownRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/AssociatedOwnerGroup/Users?$filter=Id eq ${me.Id}&$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!ownRes.ok) return false;
    const own = await ownRes.json();
    return ((own.value ?? []) as unknown[]).length > 0;
  };
```
And in the mount effect (lines 175–178) add:
```ts
    loadCanManage().then(setCanManage).catch(() => setCanManage(false));
```

- [ ] **Step 3: Inline-create state + handlers**

Add state below the search-box state (after line 70):
```ts
  const [creating, setCreating]   = useState(false);
  const [newName, setNewName]     = useState("");
```
Add handlers after `clearGroup` (line 211):
```ts
  const startCreate = (): void => {
    setCreating(true);
    setNewName(
      query.trim() ||
        suggestGroupName(mode?.label ?? "", chosen.map((t) => t.label), (role || "") as GroupMapRole | ""),
    );
  };

  const cancelCreate = (): void => { setCreating(false); setNewName(""); };

  // Warn (never block) when the typed name's suffix disagrees with the selected Role.
  const nameRoleMismatch = (): boolean => {
    if (!newName.trim() || !role || role === "GLOBAL") return false;
    return roleFromGroupName(newName) !== role;
  };

  const onCreateGroup = async (): Promise<void> => {
    const title = newName.trim();
    if (!title) return;
    setBusy(true);
    try {
      const created = await createSiteGroup(context.spHttpClient, siteUrl, title);
      pickGroup({ id: String(created.id), displayName: created.title });
      setCreating(false);
      setNewName("");
      showToast(`Group "${created.title}" created (no permissions yet — Reconciliation grants folder access).`, false);
    } catch (e) {
      if ((e as Error).message === DUPLICATE_GROUP) {
        showToast("A group with that name already exists — search for it and select it instead.", true);
        setCreating(false);
        setQuery(title); // re-runs the debounced search so the existing group appears
      } else {
        showToast(`Create failed: ${(e as Error).message}`, true);
      }
    } finally {
      setBusy(false);
    }
  };
```

- [ ] **Step 4: Render changes**

1. Intro paragraph (lines 346–351) — replace with:
```tsx
      <p style={s.intro}>
        Map a <strong>native SharePoint site group</strong> to a segment, tier, and role — or create
        the group right here and add its members. This writes a clean row into the{" "}
        <strong>DMS Group Map</strong> list. Member changes take effect <strong>immediately</strong>;
        new/deleted <em>rows</em> need a <strong>Folder Reconciliation</strong> run to apply folder
        permissions (MEMBER → Read, UPL → Contribute, APR → Design).
      </p>
```
2. Add a read-only banner right after the intro:
```tsx
      {canManage === false && (
        <div style={{ ...s.card, borderColor: "#f0c000", background: "#fff8e1" }}>
          Read-only: creating groups and editing members needs <strong>Full Control (site owner)</strong>{" "}
          on this site. You can still view mappings.
        </div>
      )}
```
3. Search placeholder (line 365) — `"Type to search Entra groups…"` → `"Type to search this site's DMS_ groups…"`.
4. Dropdown (lines 370–378) — replace with a version that always offers create:
```tsx
            {(searching || results.length > 0 || query.trim()) && (
              <div style={s.dd}>
                {searching && <div style={s.ddItem}>Searching…</div>}
                {!searching && results.map((g) => (
                  <div key={g.id} style={s.ddItem} onClick={() => pickGroup(g)}>{g.displayName}</div>
                ))}
                {!searching && results.length === 0 && query.trim() && (
                  <div style={{ ...s.ddItem, color: "#666" }}>No matching group.</div>
                )}
                {!searching && canManage === true && (
                  <div
                    style={{ ...s.ddItem, color: "#0f6c3f", fontWeight: 600 }}
                    onClick={startCreate}
                  >
                    ➕ Create a new group…
                  </div>
                )}
              </div>
            )}
```
5. Below the search box (still inside the `group ? … : …` else-branch), render the create editor when `creating`:
```tsx
        {creating && !group && (
          <div style={{ marginTop: 8 }}>
            <input
              style={s.input}
              value={newName}
              disabled={busy}
              onChange={(e) => setNewName(e.target.value)}
              placeholder="New group name"
            />
            {nameRoleMismatch() && (
              <div style={{ fontSize: 12, color: "#a4262c", marginTop: 4 }}>
                Warning: the name suffix doesn't match the selected role ({role}). You can still create it.
              </div>
            )}
            <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
              <button
                style={newName.trim() && !busy ? s.addBtn : s.addBtnOff}
                disabled={!newName.trim() || busy}
                onClick={() => { onCreateGroup().catch(() => undefined); }}
              >
                Create group
              </button>
              <button style={s.ghost} disabled={busy} onClick={cancelCreate}>Cancel</button>
            </div>
          </div>
        )}
```
(Both `s.addBtn`/`s.addBtnOff` carry `marginTop: 12` from the style map — acceptable here.)

Picking an existing group still runs `roleFromGroupName` via the unchanged `pickGroup` (line 199) — pre-selecting the Role exactly as today.

- [ ] **Step 5: Test + typecheck**

Run: `npx heft test`
Expected: PASS.

Run: `grep -n "msGraphClientFactory\|MSGraphClientV3" src/webparts/folderManager/components/GroupMapBuilder.tsx`
Expected: no output.

- [ ] **Step 6: Commit**

```bash
git add src/webparts/folderManager/components/GroupMapBuilder.tsx
git commit -m "feat: Group Map builder searches/creates native SP site groups"
```

---

### Task 8: GroupMapBuilder — member editor panel

**Files:**
- Modify: `src/webparts/folderManager/components/GroupMapBuilder.tsx` (builds on Task 7)

- [ ] **Step 1: State + handlers**

Add state below the create-flow state:
```ts
  const [membersOpen, setMembersOpen]   = useState(false);
  const [members, setMembers]           = useState<SpGroupMember[] | undefined>(undefined);
  const [memberBusy, setMemberBusy]     = useState(false);
  const [peopleQuery, setPeopleQuery]   = useState("");
  const [peopleResults, setPeopleResults] = useState<PersonPick[]>([]);
  const [peopleSearching, setPeopleSearching] = useState(false);
  const [confirmRemove, setConfirmRemove] = useState<number | undefined>(undefined);
```
Handlers (place after the create handlers):
```ts
  const groupIdNum = (): number => Number(group?.id ?? 0);

  const reloadMembers = async (): Promise<void> => {
    setMembers(await getGroupMembers(context.spHttpClient, siteUrl, groupIdNum()));
  };

  const toggleMembers = (): void => {
    const opening = !membersOpen;
    setMembersOpen(opening);
    if (opening && members === undefined) {
      reloadMembers().catch(() => { setMembers([]); showToast("Could not load members.", true); });
    }
  };

  const onAddMember = async (p: PersonPick): Promise<void> => {
    setMemberBusy(true);
    try {
      await addGroupMember(context.spHttpClient, siteUrl, groupIdNum(), p.loginName);
      await reloadMembers();
      setPeopleQuery("");
      setPeopleResults([]);
      showToast(`${p.displayName} added — access is immediate.`, false);
    } catch (e) {
      showToast(`Add member failed: ${(e as Error).message}`, true);
    } finally {
      setMemberBusy(false);
    }
  };

  const onRemoveMember = async (userId: number): Promise<void> => {
    setMemberBusy(true);
    try {
      await removeGroupMember(context.spHttpClient, siteUrl, groupIdNum(), userId);
      await reloadMembers();
      setConfirmRemove(undefined);
      showToast("Member removed — access revoked immediately.", false);
    } catch (e) {
      showToast(`Remove failed: ${(e as Error).message}`, true);
    } finally {
      setMemberBusy(false);
    }
  };
```
Reset the member state inside the existing `clearGroup` **and** `pickGroup` (so a newly picked group never shows the previous group's members) — add to both:
```ts
    setMembersOpen(false);
    setMembers(undefined);
    setPeopleQuery("");
    setPeopleResults([]);
    setConfirmRemove(undefined);
```
Debounced people search — add alongside the group-search effect:
```ts
  useEffect(() => {
    const q = peopleQuery.trim();
    if (!membersOpen || q.length < 2) { setPeopleResults([]); return; }
    let cancelled = false;
    setPeopleSearching(true);
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then((r) => { if (!cancelled) { setPeopleResults(r); setPeopleSearching(false); } })
        .catch(() => { if (!cancelled) { setPeopleResults([]); setPeopleSearching(false); } });
    }, 350);
    return () => { cancelled = true; clearTimeout(t); };
  }, [peopleQuery, membersOpen]);
```

- [ ] **Step 2: Render the panel**

Directly under the picked-group chip (inside the `group ? (…)` branch, after the `</span>` closing the chip — the branch must become a fragment `<>…</>` since it now renders more than the chip), add:
```tsx
            <button style={s.seglvl} disabled={busy} onClick={toggleMembers}>
              {members === undefined ? "members" : `${members.length} member(s)`} {membersOpen ? "▴" : "▾"}
            </button>
            {membersOpen && (
              <div style={{ ...s.preview, borderStyle: "solid", marginTop: 8 }}>
                {members === undefined && <div>Loading members…</div>}
                {members !== undefined && members.length === 0 && <div>No members yet.</div>}
                {(members ?? []).map((m) => (
                  <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "3px 0" }}>
                    <span style={{ flex: 1 }}>{m.title}</span>
                    <span style={s.mono}>{m.email}</span>
                    {canManage === true && (confirmRemove === m.id ? (
                      <span style={{ display: "inline-flex", gap: 6 }}>
                        <button style={s.delBtn} disabled={memberBusy} onClick={() => { onRemoveMember(m.id).catch(() => undefined); }}>Remove</button>
                        <button style={s.ghost} disabled={memberBusy} onClick={() => setConfirmRemove(undefined)}>Cancel</button>
                      </span>
                    ) : (
                      <button style={s.chipX} disabled={memberBusy} title="Remove from group" onClick={() => setConfirmRemove(m.id)}>✕</button>
                    ))}
                  </div>
                ))}
                {canManage === true && (
                  <div style={{ ...s.ddwrap, marginTop: 8 }}>
                    <input
                      style={s.input}
                      placeholder="Search people in the tenant to add…"
                      value={peopleQuery}
                      disabled={memberBusy}
                      onChange={(e) => setPeopleQuery(e.target.value)}
                    />
                    {(peopleSearching || peopleResults.length > 0) && (
                      <div style={s.dd}>
                        {peopleSearching && <div style={s.ddItem}>Searching…</div>}
                        {!peopleSearching && peopleResults.map((p) => (
                          <div key={p.loginName} style={s.ddItem} onClick={() => { onAddMember(p).catch(() => undefined); }}>
                            {p.displayName} <span style={s.mono}>{p.email}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
```

- [ ] **Step 3: Test + typecheck**

Run: `npx heft test`
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add src/webparts/folderManager/components/GroupMapBuilder.tsx
git commit -m "feat: Group Map member editor — add/remove SP group members inline"
```

---

### Task 9: Drop Graph scopes from package-solution.json

**Files:**
- Modify: `config/package-solution.json:10-27` (`webApiPermissionRequests`), line 6 (version)

- [ ] **Step 1: Edit the scopes**

Replace the `webApiPermissionRequests` array (lines 10–27) with:
```json
    "webApiPermissionRequests": [
      {
        "resource": "Microsoft Graph",
        "scope": "User.Read.All"
      }
    ],
```
Removed: `Group.Read.All`, `GroupMember.Read.All` (this spec) and `Sites.Read.All` (dead — nothing calls it). **`User.Read.All` stays** — `RequestShare.tsx`'s PnP PeoplePicker still needs it until the Share Guard retirement (`2026-07-23-share-guard-retirement.md`) is implemented. Do not tell the client "zero Graph" until that lands.

Bump `"version": "1.0.28.0"` → `"1.0.29.0"` (line 6).

- [ ] **Step 2: Full verification**

Run: `grep -rn "msGraphClientFactory" src/ --include=*.tsx --include=*.ts`
Expected: exactly one hit — `src/webparts/requestShare/components/RequestShare.tsx:21` (Share Guard, retired separately).

Run: `npm run build`
Expected: `ALL DONE!` — package builds with the reduced scopes.

- [ ] **Step 3: Commit**

`config/package-solution.json` carries unrelated uncommitted WIP (bulk-upload component ids). Same rule as Task 5: **default skip + record "pending commit"**; commit only if the user confirms the WIP rides along.

---

### Task 10: Full test + build gate

- [ ] **Step 1:** Run: `npx heft test` — Expected: all suites pass (groupMapModel incl. suggestGroupName, spGroups, formModel, shareGuard, pathEncoding, siteMap).
- [ ] **Step 2:** Run: `npm run build` — Expected: `ALL DONE!`, `sharepoint/solution/sd-gatrie.sppkg` produced.
- [ ] **Step 3:** Run: `grep -rn "memberOf" src/webparts --include=*.tsx` — Expected: no Graph membership endpoints left (SP `currentuser/groups` hits are fine).

---

### Task 11: Live sandbox verification (manual — Clarence)

No code. From the spec's Testing section; do these on the sandbox site after deploying 1.0.29.0.

- [ ] Group Map tab: search shows **only** `DMS_*` site groups; site Owners/Members/Visitors never appear.
- [ ] Create a group from the picker (edit the suggested name); confirm it appears in Site Settings → People and groups with **no permissions**.
- [ ] Create with a duplicate name → friendly "already exists" toast, search re-runs.
- [ ] Add a member (tenant user), confirm they appear; remove them; both toasts state the immediate-effect rule.
- [ ] Add a Group Map row with the new group → `GroupId` in the list is the integer string (e.g. `"27"`).
- [ ] Run Folder Reconciliation → log shows `↳ <group> → <perm>` lines; folder ACL shows the SP group directly (no `c:0o.c|federateddirectoryclaimprovider|…` entries).
- [ ] **Non-admin account** (NOT the SCA dev account — `IsSiteAdmin` must be `false`): member of one `DMS_*_UPL` group only → upload form auto-detects the path, upload lands in the unit folder, other units' folders are invisible.
- [ ] Non-admin in **no** DMS group → form shows the existing "no access" path.
- [ ] Non-owner opens the Group Map tab → read-only banner, create/member controls hidden/disabled.

**Cutover order (production, from the spec — out of order locks users out):**
1. Deploy the new package. 2. Create SP groups + Group Map rows. 3. Add members. 4. Delete old Entra rows from DMS Group Map. 5. Run Folder Reconciliation (its `copyRoleAssignments=false` wipe self-cleans stale Entra ACLs).

---

## Out of scope (this plan)

- Share Guard deletion (`2026-07-23-share-guard-retirement.md`) — separate small plan; until it lands, `User.Read.All` remains requested.
- Reconciliation global-reader principal (Phase 2).
- Any change to term store, DMS Config, DMS Folder Map, or the upload/approval pipeline.

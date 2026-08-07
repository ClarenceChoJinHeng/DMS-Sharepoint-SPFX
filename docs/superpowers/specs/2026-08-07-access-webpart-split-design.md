# Splitting Folder & Group Manager into two web parts

**Date:** 2026-08-07
**Status:** BUILT 2026-08-07 (`2baf48f`) — but **one page per surface**, not the single tabbed
page proposed in §2. The client chose four pages when asked. §2 and §4 below are amended to
match what was built; the reasoning in §1 and §3 is unchanged and is what justified either shape.
**Trigger:** client feedback that the Folder & Group Manager is "too complicated".

---

## 1. The problem

One web part currently carries two unrelated jobs behind four top-level tabs:

| Tab | Job | Who does it | How often |
|---|---|---|---|
| Staging | browse/rename/create folders | admin | rarely |
| Documents | same, other library | admin | rarely |
| Folder Reconciliation | provision the whole tree from the term store | admin | after any term/group change |
| **User Access** (4 sub-tabs) | map SharePoint groups to segments, tiers and roles | **access owner** | **whenever a person joins or moves** |

Two tab bars stacked on one page, and the most frequent task is buried one level down
inside the least frequent surface. The nesting *is* the complexity the client is reporting.

## 2. The split

**Page A — "Folder Manager"** (existing web part, retitled)
Staging · Documents · Folder Reconciliation. Structure and provisioning.

**Pages B–E — one per access surface**, each its own web part:

| Page | Web part | Component | Id |
|---|---|---|---|
| Folder Access | `FolderAccessWebPart` | `GroupMapBuilder` | `a5955b3e-…` |
| Site Access | `SiteAccessWebPart` | `SiteAccess` | `695640a1-…` |
| Approval Library Access | `ApprovalLibraryAccessWebPart` | `StagingAccess` | `591eb30f-…` |
| Page Access | `PageAccessWebPart` | `PageAccess` | `2d6d7010-…` |

All four share one directory (`src/webparts/userAccess/`), one bundle
(`user-access-web-parts`) and one loc file — only the entry point and manifest differ.
`AccessShell.tsx` gives them common chrome so they cannot drift apart visually.

Four pages rather than one tabbed page means **four Page Access rows**, which is the point:
the person who manages folder mappings need not be the person who manages site entry.

## 3. Why this is low risk

All four access components are already self-contained:

- `GroupMapBuilder.tsx` (1,975 lines) — Folder Access
- `SiteAccess.tsx` (471)
- `StagingAccess.tsx` (709)
- `PageAccess.tsx` (889)

Verified 2026-08-07: **no back-imports into `FolderManager.tsx`.** Every import is from
`shared/`, and each component already calls `primeNames()` itself. Props are uniform —
`{ context, siteUrl }`, plus `library` on `StagingAccess`.

So this is a **move**, not a rewrite. No logic changes to the four components.

## 4. Work

1. New SPFx web part `userAccess` — manifest (new GUID), `UserAccessWebPart.ts`,
   `IUserAccessProps.ts`, loc strings.
2. `components/UserAccess.tsx` — a tab shell reusing the existing tab styling, rendering
   the four components. Roughly the shape of `FolderManager`'s `accessTab` block.
3. Move the four component files into `src/webparts/userAccess/components/`. Import paths
   stay at the same depth (`../../../shared/`) — unchanged.
4. `FolderManager.tsx` — drop the `"GroupMap"` member of `Tab`, the `accessTab` state,
   the four imports, and the `AccessTab` table. Retitle the heading to **Folder Manager**.
5. Add the new web part to a new page under **CRS Settings**.

## 5. Naming cleanups to fold in

The current help text is stale and contributes to the "complicated" impression:

- Body text says **"DMS Group Map"**, **"DMS Upload / DMS Approve / DMS Delete"**. These
  must resolve through `cachedListTitle()` / the write-prefix helpers, not be hardcoded —
  on this site they are already `CRS *`.
- The tab reading **"Staging Library Access"** should read from `libraryTitle()`; the
  library is titled **Approval Document** now. "Staging" survives only as a LOGICAL key.
- The role blurb is four dense paragraphs above the form. Move the DEL/DELS distinction
  and the custom-permission-level prerequisites behind a collapsible "How roles work",
  leaving the form first.

## 6. The seam this creates — and the mitigation

Splitting the pages splits a workflow: **mapping a group on Folder Access does nothing to
folder ACLs until Folder Reconciliation runs on Folder Manager.** That was already true, but
the two used to be one click apart.

The note therefore sits on **Folder Access alone**. Verified 2026-08-07 by call-site count:
`SiteAccess`, `StagingAccess` and `PageAccess` each call `addroleassignment` directly and take
effect immediately; only `GroupMapBuilder` writes rows and defers to reconciliation. The
single-page version showed that caveat above all four surfaces, which was false for three.

Do **not** trigger reconciliation from an access page: it is a long, throttled, page-bound run
(it warns before unload), and burying it behind a routine mapping is how it gets interrupted
halfway.

## 6a. Packaging trap — new web parts are not registered automatically

`config/package-solution.json` pins an explicit `features[0].componentIds` array. A web part
absent from it is **compiled, bundled, and then silently dropped from the `.sppkg`** — the
build reports success, and the only symptom is that the web part never appears in the toolbox.
The first attempt at this split lost its web part exactly this way.

**Verify the package, not the build log:**

```
node -e "const s=require('fs').readFileSync('sharepoint/solution/sd-gatrie.sppkg').toString('latin1');
const m=new Set([...s.matchAll(/WebPart_([0-9a-f-]{36})\.xml/g)].map(x=>x[1]));console.log([...m])"
```

Bump `solution.version` on every packaging change so the tenant treats it as an upgrade.

## 7. Out of scope

- No change to reconciliation logic, the Group Map schema, or any role semantics.
- No change to who can see which page — that is `PageAccess`'s own job, and the new page
  needs its own row there once it exists.

## 8. Related

- Staging tab showed "No top-level folders" because `loadTree()` raced `primeNames()`;
  fixed separately the same day (`namesReady` gate). Not caused by this split.
- `docs/2026-08-07-project-state.md` — current site state.

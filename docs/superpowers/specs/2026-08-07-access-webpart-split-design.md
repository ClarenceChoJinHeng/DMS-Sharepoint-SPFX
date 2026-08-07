# Splitting Folder & Group Manager into two web parts

**Date:** 2026-08-07
**Status:** proposed — awaiting go-ahead
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

**Page B — "User Access"** (new web part)
Folder Access · Site Access · Approval Library Access · Page Access. People and roles.
The sub-tabs become the top-level tabs, so nothing is nested.

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

Splitting the pages splits a workflow: **mapping a group on Page B does nothing to folder
ACLs until Folder Reconciliation runs on Page A.** That is already true today, but today
the two are one click apart.

Mitigation, cheapest first:

1. After a successful mapping write, the confirmation states that folder permissions apply
   on the next reconciliation, **with a link to the Folder Manager page**.
2. Reconciliation already reports newly mapped units, so the reverse direction needs nothing.

Do **not** trigger reconciliation from the access page: it is a long, throttled,
page-bound run (it warns before unload), and burying it behind a routine mapping is how it
gets interrupted halfway.

## 7. Out of scope

- No change to reconciliation logic, the Group Map schema, or any role semantics.
- No change to who can see which page — that is `PageAccess`'s own job, and the new page
  needs its own row there once it exists.

## 8. Related

- Staging tab showed "No top-level folders" because `loadTree()` raced `primeNames()`;
  fixed separately the same day (`namesReady` gate). Not caused by this split.
- `docs/2026-08-07-project-state.md` — current site state.

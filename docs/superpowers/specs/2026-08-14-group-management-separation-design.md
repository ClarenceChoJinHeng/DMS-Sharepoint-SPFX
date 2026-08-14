# Separating group creation from Folder Access

**Date:** 2026-08-14
**Status:** Agreed, not yet built
**Supersedes:** the inline create-group flow in `GroupMapBuilder.tsx` (`onCreateAndMap`), added with
the persona work of `2026-08-09-persona-driven-folder-access-design.md`
**Related:** `2026-08-07-access-webpart-split-design.md`, `2026-07-22-group-map-builder-design.md`

---

## 1. The problem, in the client's words

> *"For Folder Access, for User experience wise it is confusing to client… they want the group
> creation to be separated from folder access. Think about it the group creation is done in folder
> creation and its confusing."*

They are right, and the confusion is not cosmetic. It is a **missing state in the model**.

Today the only way to create a SharePoint group from the DMS is the "➕ Create a new group" row in
Folder Access's group search, which opens a panel whose single button is **"Create group & add
mapping"** (`GroupMapBuilder.tsx:764`, `onCreateAndMap`). That action is indivisible:

1. `createSiteGroup` — creates the SP group
2. `draftRows(...)` → `postRow` per row — writes the persona's Group Map rows
3. on ANY row failure: `deleteSiteGroup` — **destroys the group again** and rethrows
4. adds the staged members, and each of them to the site-entry group
5. writes one `GroupMapChanged` audit row

So an admin who wants a group must first decide a segment, a tier and a persona. There is no way to
create a group and stop. And the reverse is enforced too: `onDelete` (`GroupMapBuilder.tsx:1110`)
deletes the SP group when its **last** mapping row is deleted.

The system therefore cannot represent *"this group exists but is not assigned to a folder yet"* —
which is precisely the state the client is asking to be able to work in.

### What group creation does NOT do

Worth stating, because it is what makes the split safe: `createSiteGroup`
(`src/shared/spGroups.ts:106`) POSTs `{ Title }` to `/_api/web/sitegroups` and **assigns no
permission level at all**. Folder ACLs are applied later, by Folder Reconciliation, from the Group
Map rows. A group with no mapping is inert — it grants nothing, anywhere. Creating one early is
harmless, and that is what makes "create now, assign later" a legitimate workflow rather than a
window in which something is half-granted.

---

## 2. Decisions

| # | Decision | Chosen |
|---|---|---|
| D1 | What moves | **The whole group lifecycle** — create, members, delete |
| D2 | Where it lives | **Its own web part and page**, alongside the four existing access web parts |
| D3 | Naming convention | Kept, via an **optional name builder** that writes nothing |
| D4 | Auto-delete on last mapping row | **Removed** |
| D5 | Site-entry group logic | **Extracted to one shared module** |
| D6 | Audit | Three new event types |

### D1 — the whole lifecycle, not just creation

The alternative was to move only creation and leave membership editable from both screens. Rejected:
membership in two places is a milder version of the same confusion. One page owns **groups and the
people in them**; the other owns **which folder a group can reach**. That sentence is the whole
mental model, and it only works if the split is complete.

### D2 — its own web part

A tab bar over all five access screens (the Folder Administration pattern of
`2026-08-12-term-abbreviation-page-design.md` §6) was offered and not chosen. Consequence, accepted
and mitigated in §4.4: with five separate pages, nothing on screen tells an admin to visit Groups
*before* Folder Access.

### D3 — the naming convention survives, decoupled

`suggestGroupName` (`groupMapModel.ts:649`) derives `GHO_GF_CORU_UPLOADER` from segment + tier +
role, and Folder Access reads that suffix back through `roleFromGroupName` to pre-select the role
when a group is picked. Create the group first and none of that context exists yet.

So the Group Management page offers **the same cascade purely as a name builder**: it fills the name
box and writes nothing else. The admin may ignore it entirely and type any name. The convention is
preserved as a convenience, not as a coupling — which is what it always was (`suggestGroupName` is
already only a suggestion, and `nameRoleMismatch` warns without blocking).

### D4 — removing the auto-delete is not a side effect, it is required

`onDelete` currently deletes the SP group along with its last mapping row. Once "created but not yet
assigned" is legitimate, that rule **silently destroys a group the admin created minutes earlier**.
Group deletion becomes explicit, on the Group Management page, and nowhere else.

### D5 — one copy of the site-entry logic

"Find or create the site-entry group, and add this person to it" exists three times today:
`GroupMapBuilder.tsx:624` (`resolveEntryGroupId`), `SiteAccess.tsx:230`, and
`FolderManager.tsx:2047`. This change would make it four. It is load-bearing and invisible when
wrong — a member outside the entry group holds a folder grant they cannot navigate to, which
presents as a permissions bug and is not one — so it moves to `src/shared/siteEntryGroup.ts` and the
three existing call sites are re-pointed at it.

### D6 — audit

`GroupCreated`, `GroupDeleted`, `MembersChanged` are added to `EVENT` in `src/shared/auditLog.ts`.
Safe by construction: `EventType` is a **Text** column precisely so that adding a type cannot fail a
write (a Choice column rejects an unlisted value and loses the whole row —
`2026-08-13-audit-log-design.md`).

`GroupMapChanged` stays as it is, and stays on Folder Access — it describes a *mapping*, which is
still that screen's business.

---

## 3. Scope

### In

- A new web part, `CRS Group Management`, and its page component
- Create a group, with an optional name builder
- Add and remove members, on any site group
- Delete a group, explicitly, with its Group Map rows
- Removing the create panel, the per-row **Members** button, the **People** tab and the auto-delete
  rule from `GroupMapBuilder.tsx`
- `src/shared/siteEntryGroup.ts`, and re-pointing the three existing call sites
- Three audit event types and their writers
- The "no match" signpost on Folder Access (§4.4)

### Out

- Any change to the `CRS Group Map` list schema, or to any existing row. **No migration.**
- Any change to Folder Reconciliation, or to how ACLs are applied
- Merging the access web parts into a tab bar (D2)
- Creating M365 groups, or anything outside site groups
- Renaming an existing group

---

## 4. The Group Management page

Admin-only, behind the same check Folder Access uses (`loadCanManage`,
`GroupMapBuilder.tsx:411`). A non-admin sees the reason, not a broken screen.

### 4.1 Create a group

- **Name** — free text, required. Validated for the characters SharePoint refuses, and for a name
  already in use (`createSiteGroup` throws `DUPLICATE_GROUP`; the message must name the existing
  group rather than say "failed").
- **Name builder (optional, collapsed by default)** — segment `<select>` from the Config `mode`
  rows, then the term cascade, then a role. Fills the name box via `suggestGroupName`. Writes
  nothing. Changing it after the admin has hand-edited the name must **ask before overwriting**.
- **Members (optional)** — the same tenant people search used today, staged before creation.
- One button: **Create group**.

On success: the group exists, it holds no permission level, and it maps to nothing. The page says so
and points at Folder Access as the next step — the mirror of §4.4's signpost, so the pair of screens
describes its own order from both ends.

**Rollback is no longer needed.** `onCreateAndMap` deletes the group when a mapping row fails,
because a group created for a mapping that does not exist was garbage. Here the group is the
deliverable. A member who fails to add is reported and the group is kept.

### 4.2 Members

A group list (all site groups, searchable — `searchSiteGroups`, which already excludes the three
built-in association groups), each expanding to its members. Add via tenant people search, remove
with the existing two-step confirm.

**Every added member is also added to the site-entry group** (D5), except when the group being
edited *is* the entry group. A failure there is reported explicitly and is never silent: the person
holds a folder grant they cannot reach the site to use.

This absorbs both of the membership surfaces Folder Access has today — the per-row **Members** modal
(`GroupMapBuilder.tsx:599`) and the read-only **People** tab (`:1578`) — including the People tab's
worthwhile details: lazily loaded members, a sequential "load all" that avoids 429s, and the
explicit note naming groups not yet read, so "no match" is never mistaken for "nobody there".

### 4.3 Delete a group

Deleting a group is the one destructive action here, and its danger is not the group.

Before confirming, the dialog lists **every Group Map row the group holds**, by segment/tier/role,
and states plainly:

> Deleting this group removes its mappings. **The folder permissions it was granted stay in place
> until Folder Reconciliation runs.**

That sentence is the point. It is the same distinction the audit log already draws — *a revoke can
remove the mapping while the access REMAINS* — and without it an admin believes access is gone when
it is not.

A group with no rows is deleted with an ordinary confirm.

Audit: `GroupDeleted`, recording the group, the number of mapping rows removed, and — when rows were
removed — that reconciliation is still outstanding.

### 4.4 What Folder Access says instead

Today, a search with no match offers "➕ Create a new group". That row is removed. In its place, when
nothing matches, the field says:

> No group by that name. Groups are created on the **Group Management** page.

The dead end becomes a signpost. This is the entire mitigation for D2, and it is the reason the
separation does not simply relocate the confusion.

---

## 5. What Folder Access becomes

Pick an existing group → pick a persona → pick segment/tier → **Add mapping**. Plus the mappings
table, with **Delete** per row and the bulk delete.

Removed: `startCreate` (`:689`), `stageMember` (`:709`), `createErrors` (`:726`),
`onCreateAndMap` (`:764`), the create panel JSX (`:1915–1998`), the member modal and its handlers
(`:599–685`, `:2282`), the People view (`:1495–1708`), `onDeleteGroup` (`:575`), and the SP-group
deletion branch inside `onDelete` (`:1110`).

`draftRows` (`:757`) and `postRow` (`:381`) stay — they are the mapping write, which is this
screen's job.

This should also take `GroupMapBuilder.tsx` under the 2000-line lint ceiling it currently exceeds
(2345).

---

## 6. Risks

| Risk | Handling |
|---|---|
| Admin creates a group and never maps it — it grants nothing and looks fine | The create confirmation names Folder Access as the next step; the group list shows a **"not mapped"** badge |
| Admin deletes a group believing access is revoked | §4.3's dialog states the reconciliation caveat before confirming, and the audit row repeats it |
| A member is added but not to the site-entry group | Reported explicitly, never silently; single shared implementation (D5) |
| The name builder overwrites a hand-typed name | Asks first |
| An existing page hosts the Folder Access web part and expects the create panel | The web part id is unchanged; only its contents change. No page needs re-authoring |
| The new web part is not registered in the package | `2026-08-07-access-webpart-split-design.md` §6a — add the bundle to `config/config.json` AND the componentId to `package-solution.json`, or it ships silently absent |

---

## 7. Test plan

Pure logic (unit tests, `src/shared/`):

- name validation: SharePoint-illegal characters, blank, whitespace-only, duplicate
- `suggestGroupName` unchanged by this work — pinned so the builder cannot drift from what
  `roleFromGroupName` reads back
- site-entry resolution: entry group present / absent / unreadable, and the "group IS the entry
  group" short-circuit

Site verification, in order:

1. Create a group with the builder; confirm the name matches the old convention exactly
2. Create a group with a hand-typed name and no builder
3. Add two members; confirm both appear in the site-entry group
4. Map the group on Folder Access; run reconciliation; confirm the folder ACL
5. Delete **one** of several mapping rows on Folder Access — confirm the **group still exists**
   (D4's regression)
6. Delete the group from Group Management with rows present — confirm the dialog lists them, and
   that folder permissions persist until reconciliation
7. Search a nonexistent group on Folder Access — confirm the signpost, and that no create option is
   offered
8. Confirm three audit rows: `GroupCreated`, `MembersChanged`, `GroupDeleted`

---

## 8. Implementation order

1. `src/shared/siteEntryGroup.ts` + tests; re-point the three existing call sites (no behaviour
   change — verify this in isolation first)
2. Audit event types + labels
3. The new web part: create → members → delete
4. Strip `GroupMapBuilder.tsx`, add the signpost
5. Register the bundle and componentId, bump the package version, deploy, run §7

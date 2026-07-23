# Native SharePoint Groups — Design

**Date:** 2026-07-23
**Status:** approved, not yet implemented
**Supersedes:** `2026-07-22-group-map-builder-design.md` (Entra group search).
**Amends:** `2026-07-16-tiered-group-detection-design.md` — the tiered detection *logic* and the
`collectMembership` hard-check are unchanged; only the source of the group ids changes.

## Problem

The DMS access model is built on **Entra (Azure AD) security groups**. Every group has to be
created by the client's security department, which is slow and outside the DMS admin's control.
The client has stated they do not want to go through that process.

They want instead:

1. Use **native SharePoint site groups** on the DMS site.
2. The Group Map picker searches **only** those site groups — no tenant Entra groups.
3. The admin can **create** groups and **assign users** to them from the same screen.

## Decision

Replace Entra security groups with **native SharePoint site groups**, site-collection-scoped
to the DMS site. Hard cutover — no dual-source support.

### Why this is simpler than what it replaces

| Concern | Entra group (today) | SP site group (new) |
|---|---|---|
| Search | Graph `/groups` + `Group.Read.All` admin consent | `/_api/web/sitegroups` — no Graph, no consent |
| Membership detection | Graph `/me/memberOf` | `/_api/web/currentuser/groups` — no Graph, no consent |
| Grant on folder | `ensureuser` with `c:0o.c\|federateddirectoryclaimprovider\|{id}` claim to obtain a principal id | group's integer `Id` **is** the principal id — assign directly |
| Add a user | blocked on the security department | `POST /_api/web/sitegroups(N)/users` by the site admin |

### Permission impact

| Scope in `package-solution.json` | Used by | After |
|---|---|---|
| `Group.Read.All` | group search (GroupMapBuilder, Onboarding, FolderManager) | **remove** — this spec |
| `GroupMember.Read.All` | `/me/memberOf` (Form, BulkUpload) | **remove** — this spec |
| `User.Read.All` | PnP `PeoplePicker` in `RequestShare.tsx` | **remove** — with Share Guard retirement (`2026-07-23-share-guard-retirement.md`) |
| `Sites.Read.All` | nothing — no Graph call in the codebase uses it | **remove** — already dead |

**`webApiPermissionRequests` ends up empty.** Once this spec and the Share Guard retirement are
both implemented, no code calls `msGraphClientFactory`, so the solution requests **zero Graph
permissions** and needs no tenant API-access approval at all.

The new member-assignment picker uses SharePoint's `ClientPeoplePickerSearchUser` endpoint, so
it adds no Graph dependency.

**Sequencing:** this spec removes the first two scopes. Emptying the array entirely depends on
the Share Guard retirement landing. If Share Guard is kept for any reason, `User.Read.All` stays
and the "zero Graph" claim does not hold — do not tell the client otherwise until both ship.

**No Entra directory role is required.** Creating and populating SP site groups needs only
**Full Control (Site Owner)** on the DMS site — a SharePoint grant. The Entra **Group
Administrator** role and the tenant "approve API access" step in SharePoint Admin Center are
both no longer needed for group work.

### Accepted constraints

1. **Site-collection scoped.** An SP group on the DMS site cannot be assigned on another site.
   Two groups sharing a name on two sites are unrelated objects; adding a person to one does
   nothing to the other. Phase 2 (`2026-07-20-multi-site-storage-phase2-design.md`) therefore
   needs a parallel group set **and parallel membership** per segment site — an ongoing
   duplication, not a one-time setup. Client has confirmed Phase 2 **will** happen.

   Scope of the impact, which is narrower than it first appears:
   - **Upload routing is unaffected.** The form runs only on the DMS site and Phase 2 keeps
     Staging central, so `DMS Group Map`, the tiered hard-check, and the `GLOBAL` role
     (`isGlobalUploader`, an *upload* flag — `formModel.ts:139`) stay single-site.
   - **Cross-site *read* needs a per-site grant regardless of group technology.** Entra groups
     would also have to be assigned on each segment site; the only saving would be maintaining
     membership once instead of N times.
   - **A global reader (e.g. an executive) cannot be served by a site-level grant**, because
     inheritance is broken per unit folder. They need Read on each unit folder, which means
     Reconciliation must learn to write a **global-reader principal** into the folder ACL
     template alongside MEMBER/UPL/APR. **Not in this spec — required for Phase 2.**
   - Site Collection Administrator would bypass all of this, but grants full destructive
     control and cannot be scoped read-only. Not an acceptable substitute.
2. **No nesting, no mailbox.** SP groups cannot contain other SP groups and have no group email.
3. **Direct membership only.** `currentuser/groups` returns groups the user is a *direct* member
   of. Users must be added individually — which is the requested workflow.
4. **No joiner/mover/leaver automation — the significant one.** Entra groups inherit the
   directory's identity lifecycle: a leaver loses group membership, and therefore DMS access,
   automatically. SP site groups do not. When someone leaves SDG or transfers department, a DMS
   admin must remove them from the SP group **by hand**, or their access persists. Membership
   also now lives in SharePoint rather than the directory where IT manages people, so the two
   can drift.

   **Mitigation:** none technical in this scope. This is a governance change the client must
   accept explicitly, and it should be written into the DMS admin handover as a standing duty
   (periodic membership review, plus offboarding checklist entry). Flag it to the client
   before implementation.

## Data model

**`DMS Group Map` — no schema change.**

| Column | Before | After |
|---|---|---|
| `GroupId` | Entra Object ID GUID | SP site group **integer Id as a string** (`"27"`) |
| `GroupName` | Entra displayName | SP group Title |
| `Segment`, `UnitTermGuid`, `Role` | unchanged | unchanged |

The integer Id doubles as the folder role-assignment principal id, so
`addroleassignment(principalid=27,...)` works with no `ensureuser` step.

**Group name is cosmetic.** `roleFromGroupName` is used in exactly one place
(`GroupMapBuilder.tsx:204`) and only to *pre-select* the Role button. Reconciliation reads the
`Role` column, never the name. A misnamed group therefore grants nothing incorrect.

## Components

### 1. `src/shared/spGroups.ts` (new) — SP group REST access

One module, no React, so all three consuming tabs share one implementation.

| Function | Endpoint |
|---|---|
| `searchSiteGroups(q)` | `GET /_api/web/sitegroups?$select=Id,Title,LoginName` — filtered client-side to `Title` starting `DMS_`, then by `q` |
| `createSiteGroup(title)` | `POST /_api/web/sitegroups` |
| `getGroupMembers(id)` | `GET /_api/web/sitegroups(id)/users?$select=Id,Title,Email` |
| `addGroupMember(id, loginName)` | `POST /_api/web/ensureuser` → `POST /_api/web/sitegroups(id)/users` |
| `removeGroupMember(id, userId)` | `POST /_api/web/sitegroups(id)/users/removebyid(userId)` |
| `searchTenantPeople(q)` | `POST .../ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser` |

Built-in groups (site Owners/Members/Visitors) are excluded by the `DMS_` prefix filter so an
admin cannot accidentally map them to a unit folder.

The people picker resolves **any tenant user**, including guests. `ensureuser` handles guests;
the group grants only what its folder role assignment gives.

### 2. `groupMapModel.ts` — one addition

`suggestGroupName(segmentLabel, tierLabels[], role)` → a suggested title such as
`DMS_GHO_GF_CORU_UPL`. Pure, unit-tested. Purely a **suggestion**: the admin edits it freely.
No abbreviation algorithm, no term-store `Code` property, no `DMS Config` mapping — labels are
joined as-is and the admin corrects them.

### 3. `GroupMapBuilder.tsx` — picker + member editor

The Role → Segment → Tier cascade and the Add-mapping flow are unchanged. Two changes:

**Group picker with inline create.** Same debounced search box, now against `sitegroups`. When
the query matches nothing, the dropdown offers a create action with an **editable** name field
pre-filled from `suggestGroupName()` using the current Role/Segment/Tier selections:

```
Group
[ DMS_GHO_GF_CORU_UPL                            ]
┌─────────────────────────────────────────────────┐
│ No matching group.                              │
│ ➕ Create "DMS_GHO_GF_CORU_UPL"        [edit]   │
└─────────────────────────────────────────────────┘
```

Create issues one `POST /_api/web/sitegroups` and selects the result. **The new group is created
with no permissions** — Reconciliation is what grants folder access — so creating a group is a
safe, reversible act.

Both directions work: picking an *existing* group still runs `roleFromGroupName` to pre-select
the Role, exactly as today.

Validation:
- **Block** on empty title, or a title already in use (SharePoint enforces uniqueness per site
  collection — catch the error and offer to select the existing group instead).
- **Warn, do not block**, if the typed name's suffix disagrees with the selected Role
  (e.g. name ends `_APR` but Role is `UPL`).

**Member editor.** Once a group is picked, an expandable panel below the chip:

```
DMS_GHO_GF_CORU_UPL ✕                    3 members ▾
┌─────────────────────────────────────────────────┐
│ Aisha Rahman    aisha@example.com          ✕    │
│ Chen Wei        chen@example.com           ✕    │
│ [ search people in the tenant…            ]     │
└─────────────────────────────────────────────────┘
```

Removal confirms inline, matching the existing row-delete pattern.

### 4. Consuming web parts

| File | Change |
|---|---|
| `Form.tsx:398` | `loadUserGroupIds` — Graph `/me/memberOf` → `/_api/web/currentuser/groups`, returning ids as strings |
| `BulkUpload.tsx:410` | identical change |
| `FolderManager.tsx:385-407` | drop the `ensureuser` federated-claim path; use the group Id as principal id. Group search → `spGroups.ts` |
| `Onboarding.tsx:207,228,279` | same two changes; remove the `Group.Read.All` notice at line 508 |

`collectMembership` in `formModel.ts` is unchanged — it compares strings and does not care
whether they are GUIDs or integers.

## Data flow

```
Admin: search/create SP group ──> add members ──────────────> immediate effect
                    │                                        (group already holds the ACL)
                    └─> Group Map row (GroupId = "27") ──> Reconciliation ──> folder ACL
                                                            (deferred effect)
```

**Two different "when does it take effect" rules**, surfaced in the UI because they are easy to
confuse:

- **Member add/remove** — immediate. The group already holds the role assignment.
- **Group Map row add/delete** — requires Reconciliation. Existing toast wording already says so.

## Cutover

Order matters — out of order locks users out mid-way.

1. Deploy the new package.
2. Create SP groups and Group Map rows (folders still on old Entra ACLs; users unaffected).
3. Add members to each new group.
4. Delete the old Entra rows from `DMS Group Map`.
5. Run Folder Reconciliation.

**Stale Entra assignments clean themselves up.** Reconciliation breaks inheritance with
`copyRoleAssignments=false` (`FolderManager.tsx:365`), wiping the folder ACL before re-adding
from the Group Map. No manual ACL cleanup is needed.

Site access follows automatically: an SP group holding a folder role assignment gives its
members Limited Access that bubbles up to the web, so members can enter the site without a
separate Visitors membership (see `dms-two-layer-access-site-plus-folder`).

## Error handling

| Case | Response |
|---|---|
| Non-owner opens the tab | Create/add/remove return 403. Detect up front and disable the controls with "needs Full Control on this site" rather than failing on click |
| Duplicate group title | Catch, show "a group with that name already exists", offer to select it |
| People picker resolves a guest | Allowed — `ensureuser` handles it |
| `currentuser/groups` returns nothing | User has no DMS group — the existing "no access" path in `Form.tsx` covers it |
| Term/`sitegroups` call fails | Existing pattern: empty list + toast, never a silent success |

## Testing

- `suggestGroupName` — new unit tests in `groupMapModel.test.ts`.
- `roleFromGroupName`, `buildGroupMapRow`, `isDuplicateRow`, `validateDraft`, `collectMembership`
  — unchanged, existing tests must still pass.
- `src/shared/spGroups.ts` REST calls — verified live against the sandbox site, not mocked.
- End-to-end: create group → add member → map row → reconcile → upload via the form → confirm the
  file lands in the unit folder and a non-member cannot see it.

### Testing must not use a Site Collection Administrator account

**SCA bypasses every ACL, including broken inheritance.** A unit folder with completely wrong
permissions looks correct from an SCA account, so testing with one proves nothing about the
permission model.

The developer account is SCA on the sandbox site — verified 2026-07-23 via
`/_api/web/currentuser?$select=IsSiteAdmin` → `true`. Check any account before trusting what it
shows you.

Every permission assertion must be verified with a **non-admin test account** holding only the
group under test:

- A user in only `DMS_..._UPL` sees exactly their unit folder in Staging — and no other unit's.
- A user in no DMS group sees nothing and gets the existing "no access" path.
- A user in only the global-reader group (when that exists) can read across units but not upload.

"I can see it" from the developer account is not evidence. Only a non-admin account is.

## Out of scope

- Bulk provisioning all three groups per unit from the term store (considered, deferred).
- Phase 2 multi-site group replication.
- Any change to the term store, `DMS Folder Map`, `DMS Config`, or the upload/approval pipeline.

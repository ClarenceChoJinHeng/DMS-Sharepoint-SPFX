# Highly Confidential Libraries & Role Expansion — Design

**Status:** agreed 2026-07-31, ready to implement
**Supersedes:** the 2026-07-16 version (elevated Power Automate flow) and the
first 2026-07-31 revision (HC folder under Document Type). See §13.
**Client-facing note:** `docs/client/highly-confidential-folder-impact.md`

---

## 1. Goal

Highly Confidential (HC) documents are physically separated into **their own
libraries**, so access is granted by library rather than carved out of the
ordinary document tree.

Confidential and Restricted remain plain metadata labels with no access effect
and no folders.

The client's PIC type 2 — *"Upload and View Staging, only HC data will show no
Confidential and Restricted showing"* — is what drives the separation. HC is a
**compartment, not a clearance tier**: HC access does not imply ordinary access,
it replaces it. A separate library expresses that directly; a subfolder inside
the ordinary tree cannot, because reaching it requires passing through folders
full of Confidential files (§13).

---

## 2. Confirmed requirements

| # | Requirement | Source |
| - | ----------- | ------ |
| 1 | HC documents live in dedicated libraries, not in Staging/Documents | client, 2026-07-31 |
| 2 | Only members of the unit's `_HC` group may upload an HC document | client, 2026-07-31 |
| 3 | Non-members must not see the Highly Confidential option at all | client, 2026-07-31 |
| 4 | PIC #2 sees **only** HC; no Confidential or Restricted content | client, 2026-07-31 |
| 5 | The unit's normal approver approves HC; no separate HC approver | client, 2026-07-31 |
| 6 | No exposure window — no peer may see the file, ever | client, 2026-07-16 |
| 7 | Confidential and Restricted create no folders; metadata only | client, 2026-07-31 |
| 8 | Single file per upload | client, 2026-07-16 |

Requirement 2 removes the elevated service the 2026-07-16 design needed. Once
the uploader belongs to the group the folder grants, reconciliation can
pre-secure the folder and the form simply writes into it.

Requirement 1 removes the corridor leak. There is no longer a path from an HC
file to a Confidential one, so requirement 4 holds by construction.

---

## 3. Library structure

Four sibling libraries. SharePoint libraries cannot nest; these are all at site
root.

| Display name | Internal / URL name | Content approval | Holds |
| ------------ | ------------------- | ---------------- | ----- |
| Documents | `Documents` | off | approved, ordinary |
| Staging | `Staging` | **on** | pending, ordinary |
| **HC Library** | `HCLibrary` | off | approved, HC |
| **HC Approval** | `HCApproval` | **on** | pending, HC |

**Create each with a no-space name, then rename the display name.** SharePoint
freezes the URL from the creation name. This keeps `%20` out of paths — relevant
because CLAUDE.md gotcha #9 records a real HTTP 400 caused by too many encoded
characters in one folder path.

**`HC Approval` must not contain the word "Staging".** Auto-route splits the
path on `Staging/` (memory `dms-autoroute-path-based-routing`); a library named
`HC Staging` would match that split and route HC files into the ordinary
Documents library. The agreed names avoid this.

### Folder tree inside the HC libraries

```
HC Approval/
├── GHO/                             ← inherits library; Read via ancestor fan-up
│   ├── GCA/                         ← inherits; Read via ancestor fan-up
│   │   ├── GMB_STRATCOMMS/          ← BREAKS INHERITANCE — the only ACL
│   │   ├── TREASURY/                ← BREAKS
│   │   └── TAX/                     ← BREAKS
│   └── GLRC/…
├── MHO/…
└── NBPOLHO/…

HC Library/                          ← identical tree
```

Files sit **directly in the unit folder**. There is no Year × Document Type
grid — Year and Document Type are columns.

Segment and department folders do not break inheritance. They receive Read
through the existing ancestor fan-up (`FolderManager.tsx:1444-1456`), which is
safe here because those folders contain only subfolders, never files.

Folder names reuse the abbreviations from `DMS Term Abbreviation`, so
`GHO/GCA/TREASURY` matches the ordinary tree exactly.

### Scope count

| | Per library | Both HC libraries |
| --- | --- | --- |
| Unit folders (break) | 133 | 266 |
| Library-level break | 1 | 2 |
| **Total** | **134** | **268** |

Microsoft recommends under 5,000 unique permission scopes per list. This is
comfortably inside it, and adding a Year or Document Type term creates **zero**
new folders.

Staging and Documents are unchanged: they keep their Year × Document Type grid,
which does **not** break inheritance (Year inherits from Unit, Document Type
from Year), so it costs creation time but no scopes.

---

## 4. Role model

### 4.1 The four Head-of bundles are a 2×2

Head of Department and Head of Unit have identical group structures, differing
only in scope (§4.4). Both reduce to two switches over a common approve
capability:

| Bundle | Approve | Upload | Delete approved |
| ------ | ------- | ------ | --------------- |
| 1      | ✓       | ✓      | —               |
| 2      | ✓       | —      | —               |
| 3      | ✓       | —      | ✓               |
| 4      | ✓       | ✓      | ✓               |

Group 4 retains delete (confirmed 2026-07-31); without it Group 4 would equal
Group 1 and the 2×2 would collapse to three.

Two new capabilities are required — **approve-without-upload** (§5) and
**delete-in-Documents** — not eight new things.

### 4.2 Atomic groups per unit

One group per (role, library). `FolderManager.tsx:1410-1413` enforces that
Staging receives only UPL/APR and Documents only MEMBER, keeping viewers off
pending documents. That rule extends to the HC pair.

| Group        | Library         | Permission level | Status                |
| ------------ | --------------- | ---------------- | --------------------- |
| `{Unit}`     | Documents       | Read             | exists (`MEMBER`)     |
| `{Unit}_UPL` | Staging         | Contribute       | exists                |
| `{Unit}_APR` | Staging         | **DMS Approve**  | exists, level changes |
| `{Unit}_DEL` | Documents       | **DMS Delete**   | new                   |
| `{Unit}_HC`  | **HC Approval** | Contribute       | new                   |
| `{Unit}_HC`  | **HC Library**  | Read             | new — same group      |

`_HC` is a single group appearing in both HC libraries at different levels,
mirroring how `MEMBER` and `UPL` split across Documents and Staging.

### 4.3 Personas are membership combinations

No bundle groups are created:

| Persona      | Groups                              |
| ------------ | ----------------------------------- |
| Head-of #1   | `{Unit}` + `_UPL` + `_APR`          |
| Head-of #2   | `{Unit}` + `_APR`                   |
| Head-of #3   | `{Unit}` + `_APR` + `_DEL`          |
| Head-of #4   | `{Unit}` + `_UPL` + `_APR` + `_DEL` |
| PIC #1       | `{Unit}` + `_UPL`                   |
| PIC #2       | `_HC` only                          |
| PIC #3       | `_UPL` only                         |
| SDG Employee | `{Unit}` only                       |

PIC #3 ("upload only, sees only their own file") needs no new mechanism:
omitting the base group removes Documents access, and Staging's item-level
permissions already restrict them to their own pending items.

PIC #2 holds `_HC` only, so they reach the HC libraries and nothing else.
Requirement 4 is satisfied without special handling.

### 4.4 Head of Department vs Head of Unit

- **Head of Unit** — their own unit. A unit-tier Group Map row already does
  this; nothing new required.
- **Head of Department** — every unit under the department. Does not work
  today; see §6.

The client's document repeated the Head of Department wording verbatim under
Head of Unit; confirmed 2026-07-31 as a copy-paste error.

---

## 5. Custom permission levels

`APR` maps to **Design** today (`FolderManager.tsx:99`). Design includes Add
Items and Delete Items, so **current approvers can already upload and delete** —
existing behaviour, not a regression, but it means "approve only" has never
actually been enforced.

Two custom levels, created once per site at Site Settings → Site Permissions →
Permission Levels. Copy an existing level rather than building from scratch, so
supporting permissions come along:

| Level | Copy from | Then |
| ----- | --------- | ---- |
| **DMS Approve** | Contribute | tick **Approve Items**; untick **Add Items**, **Delete Items** |
| **DMS Delete** | Read | tick **Delete Items** |

Copying Contribute rather than Design also avoids granting Manage Lists.

Approve Items depends on Edit Items, not Add Items, so approve-without-upload is
a legal combination.

Code change is confined to `ROLE_TO_PERMISSION` (`FolderManager.tsx:96`):

```ts
const ROLE_TO_PERMISSION: Record<string, string> = {
  MEMBER: "Read",
  UPL: "Contribute",
  APR: "DMS Approve",   // was "Design" — Design allowed approvers to upload
  DEL: "DMS Delete",
  HC: "Contribute",     // HC Approval; Read on HC Library — see §7.1
};
```

Reconciliation already logs `no "<name>" role definition on site`
(`FolderManager.tsx:1419-1423`) and skips the assignment, so deploying before
the levels exist fails safely and visibly.

---

## 6. Departmental fan-out

Head of Department must reach every unit under the department. Unit folders have
unique permissions, so a department-tier assignment stops at the department
folder and never flows down.

Reconciliation already fans Read **upward** to ancestors so members can navigate
down (`FolderManager.tsx:1444-1456`). Needed is the mirror: a department-tier
row fanned **downward** onto every unit beneath it, at the row's own permission
level rather than Read.

Constraints:

- Only **department-tier** rows fan out. Unit-tier rows must not — that is what
  keeps units isolated.
- Fan-out respects the library rule in §4.2: a department `_APR` row reaches
  Staging only; a department `MEMBER` row reaches Documents only; a department
  `_HC` row reaches the HC pair only.
- Idempotent, like every other assignment.

---

## 7. Reconciliation changes

### 7.1 Two new library targets

Reconciliation currently loops Staging and Documents. It gains HC Approval and
HC Library, with one behavioural difference: **the Year × Document Type grid
loop is skipped** for the HC pair (`FolderManager.tsx:1457-1501`).

Everything else runs unchanged — the tier walk, abbreviation naming, break at
leaf, owner Full Control, group assignment by role, ancestor Read fan-up. A
per-library flag, not new logic, which is the main reason this design is cheap
to build.

Per unit folder in the HC pair:

1. `ensureFolder` using the abbreviation
2. `breakRoleInheritance(copyRoleAssignments=false, clearSubscopes=true)`
3. Site owners → Full Control
4. The unit's `HC` rows — Contribute on HC Approval, Read on HC Library
5. The unit's `APR` group → **DMS Approve** on HC Approval

Step 5 implements requirement 5: the unit's common approver approves HC.
**Accepted consequence: unit approvers can read HC documents** — approving
requires reading, and there is no way around that. Recorded in the client note
for explicit sign-off.

`UPL`, `MEMBER` and `DEL` are **not** granted anywhere in the HC pair.

### 7.2 Known pre-existing defect, unrelated but adjacent

The grid fast path (`FolderManager.tsx:1474-1481`) probes the **last** year's
last Document Type folder and skips the whole grid if present. Adding a year
that does not sort last — 2020 into a list ending 2030 — leaves the probe
satisfied and the new year's folders are silently never created.

Affects Staging and Documents only (the HC pair has no grid). Worth fixing in
the same pass; not caused by this design.

---

## 8. Form changes

### 8.1 Filtering the Confidential Level options

The form resolves the user's SP group memberships (`spGroups.ts`) and matches
them against DMS Group Map. Confidentiality options are filtered by the roles
held **at the resolved unit**:

| Roles held at unit  | Options offered          |
| ------------------- | ------------------------ |
| `UPL` (any variant) | Confidential, Restricted |
| `HC`                | Highly Confidential      |
| both                | all three                |
| neither             | upload blocked, as today |

Requirement 3 holds by construction — a non-member never sees the option.

PIC #2, holding `_HC` only, is offered Highly Confidential alone and never
attempts to write into Staging.

### 8.2 Routing

HC uploads target a different **library**, not a subfolder:

```ts
// Term GUID, not label: the term store is client-editable in-site, and a
// renamed term must not silently reroute confidential documents into the
// ordinary library with no error anywhere.
const isHighlyConfidential =
  normGuid(confidentiality) === normGuid(settings.hcTermGuid);
```

- **Ordinary:** unchanged — resolve the unit folder by UniqueId, ensure
  `Year/DocumentType`, upload there.
- **HC:** resolve the unit folder in **HC Approval** by UniqueId from DMS Folder
  Map, upload directly into it. No Year or Document Type folder is created;
  both are written as metadata.

`hcTermGuid` comes from a `DMS Config` setting row `term_highlyConfidential`
with a code fallback in `DEFAULT_SETTINGS`.

DMS Folder Map must therefore key on **(term GUID, library)**, not term GUID
alone — one unit term now maps to four folders. See §10.

### 8.3 The form never creates a folder needing a unique ACL

If the unit folder is missing from HC Approval, reconciliation has not run:

> "Your unit's Highly Confidential folder hasn't been set up yet. Ask an
> administrator to run reconciliation before uploading Highly Confidential
> documents."

The uploader holds Contribute and cannot `breakroleinheritance` (that needs
Manage Permissions, which only Full Control has). A folder created by the form
would inherit, exposing the file. This guard is the safety property the whole
design rests on.

---

## 9. Auto-route and the Created By problem

### 9.1 Second route

Auto-route currently splits the path on `Staging/` and copies approved files to
Documents. It gains a second route: `HC Approval/` → `HC Library/`, same
path-based logic, same trailing-slash strip, same "Create new folder" behaviour.

Because the HC tree has no Year or Document Type folders, the copied path is
just `Segment/Department/Unit`.

### 9.2 `UploadedBy`

**Problem:** after approval, the copied file's **Created By** shows the
approver, not the uploader.

**Cause:** Power Automate's *Copy file* creates a new file at the destination.
`Author` and `Editor` on a new file are set from the identity performing the
action; nothing carries the original uploader across. Approving also restamps
`Editor` to the approver — the reason memory `dms-uploader-column-is-author`
requires the Staging "Uploader" column to be `Author` and never `Editor`.

`Author` is read-only through the SharePoint connector, so it cannot be written
back after the copy.

**Fix:** a dedicated person column that nothing system-managed ever touches.

1. Add **`UploadedBy`** (Person or Group, single) to all four libraries.
2. The form writes it at upload time — the uploader is the current user, so it
   is one more field in the existing `validateUpdateListItem` call.
3. Auto-route maps it across with the other metadata.

It then survives approval, the copy, and later edits.

**Open:** whether the flow's destination step is *Copy file* or *Create file* +
*Update file properties* decides whether this is a metadata mapping or an extra
action. Needs confirming against the live flow.

---

## 10. Data model changes

`groupMapModel.ts:4`:

```ts
export type GroupMapRole =
  | "MEMBER" | "UPL" | "APR" | "DEL" | "HC" | "GLOBAL";
```

`roleFromGroupName` (`groupMapModel.ts:68`) gains two suffixes, checked before
the `MEMBER` default:

```ts
if (n.endsWith("_HC")) return "HC";
if (n.endsWith("_DEL")) return "DEL";
```

No schema change to `DMS Group Map` — `Role` is text and both values are new
members of an existing vocabulary.

**DMS Folder Map does change.** It currently keys term GUID → folder UniqueId,
assuming one folder per term per library pair. With four libraries a unit term
maps to four folders, so rows must carry the **library** as part of the key.
Existing Staging/Documents rows must keep working during and after the change.

### Columns on the HC libraries

Full parity with Staging so nothing is lost in the split: Business Segment /
Department / Unit and their Tid pairs, `Year_x002f_Period`,
`Document_x0020_Type`, `DocumentDate`, `Confidentiality_x0020_Level`, `Vendor`,
`_ExtendedDescription`, `PermTermGuid`, `UploadedBy`, plus `Remark` and
`LegallyPrivileged` from the 1.0.54.0 form work.

`Confidentiality_x0020_Level` is kept even though every item here is HC — it
keeps metadata consistent across all four libraries and avoids special-casing
the form's write path.

### Views

Default view on both HC libraries is **flat** — "Show all items without
folders" — grouped by Document Type, sorted by Year descending. Users land on
their own files with no navigation. SharePoint security-trims the view, so a
member sees only their units.

---

## 11. Bulk group provisioning

Five groups × 133 units ≈ **665**, plus department-tier groups. Hand-creation is
not viable — weeks of work, and a mistyped Group Map row surfaces only as
reconciliation's "group not found".

The Onboarding web part gains a generator driven by the same term walk
reconciliation uses. Every primitive exists:

- `spGroups.ts:61` — create a site group
- `spGroups.ts:99` / `:113` — add / remove members
- `groupMapModel.ts:93` `suggestGroupName` — derives
  `DMS_Group Head Office_Group Finance_TREASURY_UPL` from term labels

Per unit term: derive the five names → create those missing → **write the
matching `DMS Group Map` row automatically**. That last step eliminates the
hand-typing that causes most reconciliation failures.

Dry-run preview first, then execute. Idempotent.

### Defect that must be fixed first

`spGroups.ts:31` lists site groups with `$top=500`. At 665+ groups the response
truncates **silently**, and every group past the cut is reported as missing by
both the Group Map builder and reconciliation. Page it before bulk creation
runs, or the first symptom is several hundred spurious "group not found" errors
that look like a data problem and are not.

---

## 12. Out of scope

- **Documents-library delete UI.** `DEL` grants the SharePoint permission;
  deletion happens through the normal library interface.
- **Migrating HC documents already uploaded** into Staging/Documents. Moving
  them changes ACLs and breaks Auto-route's path assumption — separate decision.
- **Confidential and Restricted folders.** Not created; metadata only.
- **Phase 2 sub-folders below Unit.** Still deferred.

---

## 13. Design history

**2026-07-16 — elevated flow.** HC folders under Document Type, created and
secured on demand by an HTTP-triggered Power Automate flow under a service
account (secure → upload → tag). Correct given its requirement that *any unit
uploader may create an HC file*: an ordinary uploader writing into a folder that
excludes them is impossible without elevation.

**2026-07-31 morning — pre-secured folder under Document Type.** The client
restricted HC upload to `_HC` members, so the uploader now belongs to the group
the folder grants and pre-securing became possible. Reconciliation already
pre-creates the Year × Document Type grid, so it could secure an HC folder in
each cell. Abandoned because:

- ~7,300 unique permission scopes per library, 1.5× Microsoft's recommendation
- ~20 hour first reconciliation; ~3 hours for every new Year or Document Type
- **Unsolvable corridor leak:** PIC #2 needs Read on the Document Type folder to
  browse to its HC child, and that folder holds the Confidential files they must
  not see. Limited Access does not help — it permits direct-link access only,
  not browsing (memory `dms-two-layer-access-site-plus-folder`).

**2026-07-31 afternoon — separate libraries.** Client's proposal. Resolves all
three: 134 scopes per library, under 30 minutes to build, zero folders per new
Year or Document Type, and no corridor at all because HC and ordinary documents
never share a path.

Preserved throughout: no-exposure-window (met by pre-securing), `PermTermGuid`,
config-driven HC identification (tightened from label to GUID), and reuse of the
unit's `APR` group for HC approval.

---

## 14. Restoring what Phase 1 removed

Highly Confidential went **out of scope for Phase 1** on 2026-08-01: the term is
deleted from the term store so the dropdown cannot offer the level, and this
branch is parked until Phase 2. Four things must be undone before any of the code
above works, and three of them fail silently if missed.

| # | Step | If missed |
| - | ---- | --------- |
| 1 | **Recreate the Highly Confidential term** in the Confidentiality Level set | The level cannot be selected; nothing else matters |
| 2 | **Update the `term_highlyConfidential` row** in DMS Config to the term's NEW guid | Silent — nothing matches, so every HC document routes to Staging with no error |
| 3 | **Update `DEFAULT_SETTINGS.hcTermGuid`** in `Form.tsx` | Silent, and only when DMS Config is unreadable: the worst kind of stale fallback |
| 4 | **Restore the `Highly Confidential` definition** in the Confidential Level tooltip (`Form.tsx`), removed on `feat/folder-abbreviations` in commit `7745d2b` | Cosmetic — the level is selectable but unexplained |

**A recreated term gets a new GUID.** Term GUIDs are per term and per site, so
`420d75d5-f3b7-4525-9eb7-ec06590e7f22` — the value in the Phase 1 config row and
in this branch's code fallback — will be dead. Read the new one off the site:

```
GET <site>/_api/v2.1/termStore/sets/0d6d1da8-27e5-477f-8684-e8cf169f8fb9/terms
```

Steps 2 and 3 are the dangerous pair. This design matches on GUID rather than
label precisely so a term *rename* cannot reroute confidential documents
unnoticed — but a stale GUID produces that same failure by another route. If HC
uploads land in Staging once Phase 2 starts, check these two before anything else.

Rebasing this branch onto Phase 1 also reapplies the tooltip deletion, so step 4
is a real edit rather than a merge artefact to discard.

---

## 15. Open items

| Item | Owner | Blocking? |
| ---- | ----- | --------- |
| Create the four libraries with no-space URL names | client / admin | Yes |
| Create `DMS Approve` and `DMS Delete` permission levels | client / admin | Yes — reconciliation warns and skips until they exist |
| Read the Highly Confidential term GUID for `term_highlyConfidential` | Clarence | Yes — the form cannot route without it |
| Confirm the Auto-route destination step (Copy file vs Create + Update) | Clarence | Yes, for §9.2 only |
| Confirm approvers reading HC documents is acceptable | client | No — assumed yes; flagged in the client note |
| Confirm department-tier fan-out applies to `DEL` and `HC` | client | No — assumed yes, consistent with other roles |

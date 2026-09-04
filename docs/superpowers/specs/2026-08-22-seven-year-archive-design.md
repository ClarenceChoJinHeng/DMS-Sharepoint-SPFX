# The seven-year archive

**Date:** 2026-08-22
**Status:** design agreed, not yet built
**Client instruction:** *"After 7 years the file in the Documents Library should be archive. So we need
another library."* Then, on audience: *"knowing client they want the same group to have access."* Then:
*"for all the Archive files make sure all groups have read only."*

---

## 1. Scope, and the timing that shapes it

Two new libraries — **`Archive`** and **`HC Archive`** — mirroring `Documents` and `HC Documents`. A
document moves when it passes seven years from its **`Created`** date. Everyone who could read it before
can still read it, and **nobody can do anything else to it**.

⚠ **NOTHING ON THIS SYSTEM IS CLOSE TO SEVEN YEARS OLD.** The oldest document dates from July 2026, so
under this rule the first file becomes eligible in **2033**. That is not an argument against building it
— the client asked for it — but it means:

- the feature ships **untested against real data** unless it is tested with a backdated file, and it must
  be (§8)
- both libraries and their whole folder tree sit empty for seven years while reconciliation maintains
  them on every run
- **`Created` is the only input**, so a wrong `Created` is a wrong archive date. Both routing flows
  currently stamp `Created` ~8 hours early (a UTC value written in a locale format SharePoint reads as
  local), which pulls a document uploaded before 08:00 back to the previous day. Harmless at a seven-year
  boundary, but it is the field this feature reads and it is known to be wrong.

## 2. Folder shape — one tree, no year partition

The client's instinct was a year at the top of the archive tree:

```
Archive/2026/GHO/GF/TAX/2024/Term Sheet/
```

**The collision it is meant to prevent is real.** `report.pdf` uploaded 2026 leaves `Documents` in 2033;
a new `report.pdf` in the same folder uploaded 2034 archives in 2041; both want one archive path.

**It is not being built that way**, for two reasons:

1. ⚠ **Reconciliation cannot build it.** Reconciliation walks the **term tree** — segments, departments,
   units — and there is no year in the term store. A year on top is a complete copy of the whole tree per
   year, each with its own broken inheritance and its own grants on every department and unit folder. The
   scope count is fine (~130 per year against a 50,000 ceiling); the run cost is one extra full tree
   **per year, for ever**.
2. **Two year folders in one path mean different things** — the archive year and the document's own `Year`
   metadata tier — and will be misread by whoever browses it.

**So: one tree, identical in shape to `Documents`, and a collision is resolved by renaming.**

### 2.1 The rename rule

The mover suffixes with the document's creation year: `report (2026).pdf`.

**The suffix carries the value that distinguished the files, never a number.** That is the rule the
subtree migrator already runs on (`(Archive 2)` when Archive was removed, `(2025)` when Year was) and it
is right here for the same reason: `(1)` and `(2)` throw away the only thing that told the two documents
apart.

⚠ **ONE file keeps the original name.** The one already in place; only the arrival is renamed.

### 2.2 If the client rejects the rename

Agreed fallback, **not built unless asked**: on a collision the mover **moves nothing and notifies the
system administrator**, leaving the file in `Documents` for a person to resolve.

Worth stating the trade honestly when it is offered: a rename is automatic and lossless, while a
notification is a queue somebody has to work — and a file that failed to archive stays in the working
library looking like it was never eligible. If it is taken, the notification must **name the file, both
paths and the clash**, or it is unactionable.

### 2.3 Browsing by year, if they want it

A **view grouped by `Created` year** gives the client everything the year folder would have, with no
folders, no extra grants and no extra reconciliation cost.

## 3. Permissions — every role is Read

`LIBRARY_ROLES` gains two rows, `Archive` and `ArchiveHC`, carrying the same roles as their working-side
counterparts and mapping **every one of them to `Read`**.

- `Archive`: `MEMBER · DEPTVIEW · GLOBAL · SEGVIEW · UPL · APR · DEL · SHARE · UPLHC`
- `ArchiveHC`: `UPLHC · MEMBERHC · APR · DEL · DEPTVIEW · GLOBAL · SEGVIEW · SHARE`

⚠ **Never read `ROLE_TO_PERMISSION` directly.** That table is flat — one level per role — so it would
grant `CRS Upload` (`UPL`/`UPLHC`), `CRS Approve` (`APR`), `CRS Delete` (`DEL`) and `CRS Share` (`SHARE`)
on the archive. Everything goes through `permissionForRole(lib, role)`.

**The mechanism is a new `READ_ONLY_LIBS` set, checked BEFORE the role list:**

```ts
const READ_ONLY_LIBS: LibTarget[] = ["Archive", "ArchiveHC"];

function permissionForRole(lib: LibTarget, role: string): string | undefined {
  if (READ_ONLY_LIBS.indexOf(lib) > -1) return ROLE_TO_PERMISSION[role] ? "Read" : undefined;
  if (APPROVED_SIDE_LIBS.indexOf(lib) > -1 && DOCUMENTS_READ_ONLY_ROLES.indexOf(role) > -1) return "Read";
  return ROLE_TO_PERMISSION[role];
}
```

⚠ **A LIBRARY rule, never an extension of `DOCUMENTS_READ_ONLY_ROLES`.** Adding `DEL` and `SHARE` to that
list is the obvious-looking edit and it is wrong: that list applies to `APPROVED_SIDE_LIBS`, so it would
**silently strip a Head of Unit's delete and share rights on `Documents`** — reversing a client decision
of 2026-08-20 as a side effect of building an archive. The archive is read-only because of what the
library *is*, not because of which roles reach it.

⚠ **It still returns `undefined` for an unknown role.** Answering `"Read"` unconditionally would grant on
a hand-written row naming a role that does not exist, and `ENTRY` — deliberately absent from
`LIBRARY_ROLES` — must keep being skipped rather than approximated.

⚠ **`UPL`/`UPLHC` are in the list on purpose, at Read.** A PIC must still find their own archived
documents; leaving them out would hide an uploader's own history from them.

⚠ **`SEGVIEW` and `GLOBAL` are safe here and are not on the approval side** — the archive holds only
approved documents, so there is no unapproved draft for a segment-wide viewer to reach.

### 3.1 What this takes away, and the client must be told

**Archiving a document removes the ability to delete or share it.** `DEL` and `SHARE` are listed above
but resolve to Read, so a Head of Unit or Head of Department who can delete or share an approved document
in `Documents` **cannot** once it has moved.

That is almost certainly what "archive" is meant to mean, and it should be said out loud rather than
discovered.

### 3.2 ⚠ Deletion and share requests must REFUSE an archived file

This is the consequence that fails **silently** if it is missed.

**The approval executes in the approver's own browser session, with their own permissions.** An approver
holding Read on the archive would have the recycle call fail as them — the request is recorded `Failed`,
after the requester has been told it is being handled.

So `validateDraft` **refuses** a request against an archived file, exactly as it already refuses a share
against a pending one, and for the identical reason: `SHARE` is not in `LIBRARY_ROLES.Staging` there, and
neither `DEL` nor `SHARE` is effective here.

**Refused in the rules, not merely hidden in the UI** — hiding a button turns a permission rule into a
presentation decision, and the next screen to render a request forgets it.

The refusal message says an archived document cannot be deleted or shared, and that this is what
archiving means. It must not read as a fault.

## 4. Reconciliation builds it

Both libraries join `reconLibs()` / `allLibraryTitles()`, so the existing passes apply unchanged:

- the folder tree from the term tree, keyed on abbreviations
- inheritance broken per department and per unit folder
- group grants from the Group Map, through `permissionForRole`
- the ancestor-browse corridor, without which the library renders empty for every non-admin
- deleted-group detection, skip-existing-grants, the stray-folder quarantine

⚠ **This is what makes the archive's ACLs free and correct.** The alternative considered was building the
tree and the grants in Power Automate, and it was **rejected**: it would be a second implementation of
`LIBRARY_ROLES` living outside source control with no tests, holding its own copy of role-to-permission
rules that changed four times in one client instruction on 2026-08-20. Three specific traps came with it —
role definition ids are per-site integers, the flow would have to re-derive the Group Map, and a folder a
flow creates **inherits its parent**, which is exactly how a unit folder becomes readable by its whole
department.

**The flow's only job is moving files.**

### 4.1 Cost

`allLibraryTitles()` goes 4 → 6, so a full run does roughly 50% more work. MHO's first run was 71 minutes
across four libraries. Accepted on the client's instruction that the archive is always reconciled.

**The first run is the expensive one**; the skip-existing-grants pass (1.0.177.0) means a settled site
re-reads rather than re-writes.

### 4.2 Two assertions that must not be forgotten

⚠ **Content approval must be OFF on both archive libraries**, and reconciliation should say so if it is
on. With moderation on, every folder the mover creates and every file it moves arrives **Pending** and is
invisible to every reader — it presents as a permissions bug and is not one. The same trap is already
documented for `Documents`.

⚠ **The full column set, not a subset.** SharePoint's copy and move carry over only columns that **exist
at the destination**; the rest vanish with no error. That is the `Remark`/`LegallyPrivileged` gap of
2026-08-10, which stripped a legal marker off approved documents for months. So the archive pair needs
every tier column and its `Tid` sibling, `Document Type`, `Year`, `DocumentDate`, `Confidentiality Level`,
`LegallyPrivileged`, `Remark`, `Vendor/CustomerName`, `Full Name`, `SubmissionId`, `BatchId`,
`BulkImport` and the `CRS Folder` content type — asserted every run over `allLibraryTitles()`, which is
already how the last three are handled.

## 5. The `Archived` tag

An **`Archived`** Yes/No column, written by the mover, present on all six libraries.

It is not decoration. My Submissions and CRS Search both need it to **badge** a row, the same way the `HC`
tag was added on 2026-08-21 — otherwise an uploader sees one of their own files sitting somewhere new with
nothing to explain it. It is also what lets the request refusal in §3.2 be decided from the row already
loaded, with no extra read.

⚠ **Written CONDITIONALLY, on a confirmed column**, through `libraryHasColumns` — one unknown field name
fails the whole `validateUpdateListItem` call and takes every other column with it.

## 6. The mover

**Build steps: `2026-08-22-seven-year-archive-mover-runbook.md`.** Power Automate config is not in
source control, so that runbook is the only record the flow will ever have.

A scheduled Power Automate flow, built **as the service account** (a flow runs under its connection, and
the connection is created implicitly by the first action, so the builder's account is baked in for life).

- **Weekly**, not daily. A seven-year boundary does not need daily precision, and each run is a full
  library scan.
- Finds items in `Documents` where `Created` is more than seven years ago. `Created` is indexable; the
  filter must be index-served or it throws the 5,000-item threshold rather than returning fewer rows.
- ⚠ **`Documents` to `Archive`, `HC Documents` to `HC Archive`, never crossing.** If the archive pair is
  unresolved the run **refuses** rather than falling back — `hcRouting.ts` fails closed for exactly this
  reason, and a fallback here would move a Highly Confidential document into the open archive.
- ⚠ **Ensure-creating the destination folder is NOT safe below the unit.** A unit folder the flow creates
  inherits the department folder's ACL instead of carrying the unit's. **The unit folder is RESOLVED,
  never created** — absent means reconciliation has not run, and the run must say so rather than
  improvise. Below the unit, ensure-creation is correct: that tier inherits by design.
- ⚠ **Capped per run.** Auto-route's daily quota is shared, and an exhausted quota means the next real
  approval is not routed, **silently**. A large first sweep must be spread across runs.

### 6.1 Move, not copy-stamp-delete

Auto-route copies, stamps `Author`/`Editor`/`Created`, then deletes the source — because it has to
recreate authorship on the copy.

**`MoveTo` should be used instead**, and it is worth the verification it needs:

- `Created`, `Created By` and approval status survive a move (verified 2026-08-10/11 for folder and file
  moves during the subtree migration work)
- it avoids the `Created` stamping bug in §1 entirely, because nothing is re-stamped
- **`UniqueId` survives a move within a library**, which is what would let existing deletion and share
  request rows keep resolving

⚠ **Must be verified before it is relied on: does `UniqueId` survive a move ACROSS libraries?** It is
verified within one. If it does not, request rows pointing at an archived file break, and the mover needs
to record the mapping or the requests need to be closed on archive.

## 7. Reads that must be extended, or documents silently vanish

Every screen reading a fixed set of libraries is a change site. **This is the half of the feature that
costs more than the libraries do.**

| Screen | Reads today | Must also read |
|---|---|---|
| My Submissions | `Approval Document`, `Documents` (+ HC) | `Archive`, `HC Archive` |
| CRS Search | 4 libraries, KQL on the approved side | both archive libraries in `kqlPathScope` |
| Document details panel | per-item, library-agnostic | nothing — already fine |

⚠ **My Submissions is the one that matters.** Its reads are `AuthorId eq me`; miss the archive and an
uploader's own seven-year-old documents disappear from their history with nothing to say why. Index
**`Created By`** on both archive libraries at provisioning — past 5,000 items that filter starts failing.

⚠ **Doubling the reads on that page** is the accepted cost. It opens with four library reads instead of
two.

## 8. Testing it, without waiting until 2033

**Backdate a `Created` value.** It cannot be typed, but `validateUpdateListItem` with
`bNewDocumentUpdate: true` writes it — that is precisely what Auto-route does to preserve the uploader's
timestamp, and `Created` on that endpoint wants `M/d/yyyy h:mm tt`, not ISO.

Test order:

1. Reconcile, and confirm both archive libraries have the tree, the grants and the browse corridor
2. Backdate one file in `Documents` to 2018; run the mover; confirm it lands with metadata,
   `Created By` and `Created` intact, and `Archived` ticked
3. Backdate a second file **with the same name in the same folder** and confirm the rename
   (`report (2018).pdf`), with the first file untouched
4. Sign in as a PIC: the archived file is visible and read-only; it appears in My Submissions badged
   `Archived`; the deletion and share buttons are refused with the archive message
5. Repeat 2 and 4 on the HC pair, and confirm an uncleared uploader sees nothing

## 9. What the client must be told

1. **Archiving removes the ability to delete or share the document.** Everyone who could read it still
   can; nobody can act on it.
2. **The first document becomes eligible in 2033.** The libraries will be empty until then.
3. **A name collision is resolved by renaming**, with the year appended — or, if they prefer, by
   notifying an administrator and leaving the file where it is (§2.2).
4. **Reconciliation runs ~50% longer**, every time, from the day this ships.
5. **Nothing enforces that an archived document stays archived.** An administrator can move it back.

## 10. Out of scope

- Un-archiving as a feature. Moving a file back is a manual action.
- Retention or disposal — this moves documents, it never deletes them.
- A year-partitioned archive tree (§2), and a per-segment archive schedule. Seven years is site-wide.
- Archiving anything in either approval library. A document that never got approved is not a record.

# Deleting a segment

**Date:** 2026-08-14
**Status:** Agreed, not yet built
**Completes:** `2026-08-12-add-segment-design.md` §6, which excluded this as having "no safe meaning"
**Related:** `2026-08-11-subtree-migration-design.md` (moving the files out first),
`2026-08-14-group-management-separation-design.md` (the delete-confirmation pattern this follows)

---

## 1. Why

A segment can be created from the UI but removed only by hand-editing a row in `CRS Config`. The
client will not do that after handover — the same reasoning that produced the Term Abbreviation
editor and the Group Management page. The current advice is literally "delete its `mode` row": a
list view, a `ConfigType` filter, and a row nobody wants to be wrong about.

It is not hypothetical. The live site carries an `Industrial` segment whose department term is named
`department` and whose unit is named `company` — test data that provisioned real folders
(`Industrial/test/test`) on 2026-08-14, with no way out.

The 2026-08-12 spec was right that there is no safe meaning for deleting *everything*. This one
answers that by making the default delete **neither the documents nor the columns**.

The client's framing of the optional half (verbatim): *"let say they already move all the files then
deleting would be easier for them."* That is the workflow this is built around — **move the files
out with the subtree migrator, then remove the empty shell.**

---

## 2. Decisions

| # | Decision | Chosen |
|---|---|---|
| D1 | What Delete means by default | **Retire** — the `mode` row and its Group Map rows. Nothing else. |
| D2 | Folders | **Optional**, off by default, behind a typed confirmation |
| D3 | Columns | **Never deleted**, under any option |
| D4 | Abbreviation rows | **Kept** — authored data with no other copy |
| D5 | Folder Map rows | Deleted **only** when the folders are |
| D6 | Order | Retire FIRST, folders second |
| D7 | Placement | The **New segment** tab, retitled **Segments** |

### D1 — Delete retires; it does not destroy documents

Removing the `mode` row is the whole of what "this segment should stop existing" means to the
system: the upload form stops offering it, and reconciliation stops walking it. It is also
**reversible** — re-creating the row with the same `Title`, `TermSetGuid`, `StagingFolder` and
`Levels` restores the segment over folders that never moved.

Its Group Map rows go with it, because a row whose segment no longer exists is a grant nobody can
see or manage: it appears on Folder Access under no segment, yet reconciliation keeps re-applying
it. That is the worst residue to leave — access that is live and invisible.

**Folder permissions survive until Folder Reconciliation runs.** Same sentence, same reason, as the
group delete: removing the mapping is not removing the access.

### D2 — folders are opt-in, and the recycle bin is what makes it acceptable

Deleting the segment's top folder in each library takes its whole subtree with it. Off by default
and behind a typed confirmation, because one mis-click would otherwise take a live segment's entire
archive.

Two things make it safe enough to offer at all:

- A SharePoint folder delete goes to the **recycle bin** (93 days), so it is recoverable. This is
  stated on screen: it is the difference between careful and terrified.
- The count is shown before confirming (§4.2), so an admin who believes the segment is empty finds
  out here if it is not.

### D3 — columns are never deleted

The tier columns (`Region`/`RegionTid`, …) hold metadata on documents. Deleting a SharePoint column
deletes its data irreversibly and does **not** go to the recycle bin — and those documents may have
been MOVED elsewhere rather than deleted, which is exactly the workflow this feature serves. An
unused column costs a line in the field list; a deleted one costs the filing history of every
document that was ever in that segment.

If the client wants them gone, that is a deliberate act in list settings, one column at a time, with
SharePoint's own warning in front of them.

### D4 / D5 — abbreviations stay, Folder Map rows follow the folders

An abbreviation row is **authored data with no other copy**: someone chose `UPSUPPORT` for that
term, and if the segment is ever re-created that choice must survive or every folder silently gets a
new name. Keyed by term GUID, dormant at no cost. This matches reconciliation, which has never
auto-deleted them.

Folder Map rows are **derivable** — reconciliation rebuilds them — so they go exactly when the
folders do. Deleting them while the folders stand just makes reconciliation re-create them; keeping
them after the folders are gone leaves rows pointing at dead `UniqueId`s.

### D6 — retire first, folders second

If the folder delete ran first and the row write then failed, the segment would still be live with
no folders: the upload form offers it, and every upload fails or silently re-creates a shell. The
chosen order fails safe — the segment is gone from every screen and some folders remain, visible
only to admins and removable by hand.

### D7 — the Segments tab, not Folder levels

The client asked for it "in the Folder levels itself". Putting it on **New segment** instead, and
retitling that tab **Segments**: Folder levels edits the structure *inside* a segment, so a Delete
button under a list of levels reads as "delete this level" — the most expensive misreading available
on that screen, since removing a level triggers a migration.

---

## 3. Scope

### In

- A segment list on the Segments tab, each row with a **Delete** button
- Counting folders and documents in both libraries before confirming
- Retire: delete the `mode` row and the segment's Group Map rows
- Optional: delete the segment's top folder in both libraries, plus its Folder Map rows
- A `SegmentDeleted` audit event
- Retitling the tab to **Segments**

### Out

- Deleting term sets or terms — a term GUID is referenced by three lists; rename, never delete
- Deleting tier columns (D3)
- Deleting abbreviation rows (D4)
- Editing an existing segment's label, `StagingFolder` or term set — a separate job
- Undo. The recycle bin is the undo for folders; re-creating the row is the undo for the rest.

---

## 4. The flow

### 4.1 The list

One entry per `mode` row: label, key, `StagingFolder`, and **Delete**. The existing load selects
`Title,ModeLabel,StagingFolder,SortOrder` and must also take **`Id`** (to delete) and
**`TermSetGuid`** (to find the segment's Group Map rows).

### 4.2 What the dialog counts, and what it does when it cannot

On opening Delete, count in **both** libraries beneath `<StagingFolder>`: folders, and documents.
Also count the segment's **Group Map rows**, matched on `Segment` = the mode's `TermSetGuid`.

| State | Retire | Delete folders |
|---|---|---|
| Counted, no documents | plain confirm | offered, plain confirm |
| Counted, documents present | typed confirmation | offered, typed confirmation, count stated |
| **Could not count** | allowed, typed confirmation | **NOT OFFERED** |

**An uncountable segment must not offer the folder delete.** Everywhere else in this codebase a
failed read fails OPEN, because the cost is a dropdown showing too much. Here the cost is deleting
an archive nobody could confirm was empty — so this one fails CLOSED, and says why, rather than
greying out a checkbox with no explanation.

### 4.3 The confirmation

The typed confirmation is the **segment's own label**, not the word `DELETE`: typing the name is the
check that the admin is looking at the segment they think they are.

The dialog states, in this order:

1. What stops: the segment disappears from the upload form, and reconciliation stops walking it.
2. What is deleted: N Group Map rows — and, if ticked, the folders and their N documents.
3. What survives: **the tier columns, the abbreviations, and (unless ticked) every folder and file.**
4. **Folder permissions stay in place until Folder Reconciliation runs.**
5. If folders are ticked: **they go to the recycle bin and can be restored for 93 days.**

### 4.4 Execution, in order

1. Delete the Group Map rows for this segment; count successes.
2. Delete the `mode` row. **If this fails, stop** — and say the segment is still live. A half-run
   that removed the grants but left the segment is a segment whose uploaders have quietly lost
   access.
3. If folders were ticked: delete `<StagingFolder>` in each library, then this segment's Folder Map
   rows.
4. Reload the list.

Every step reports what it did. A partial run is reported as partial, never rounded up to success.

### 4.5 Audit

`SegmentDeleted`, carrying: label and key, whether folders were deleted, the counts (Group Map rows,
folders, documents), what survived, and the reconciliation caveat. `Outcome: "Failed"` whenever any
step fell short of what it attempted.

---

## 5. Risks

| Risk | Handling |
|---|---|
| Deleting a live segment believing it is a test one | Typed confirmation is the segment's own label; counts shown first |
| Believing the documents are gone when they are not | The dialog lists what survives, explicitly, on its own line |
| Believing access is revoked when it is not | The reconciliation caveat, in the dialog and in the audit row |
| Deleting folders that were not actually empty | Counted first; uncountable ⇒ the option is withheld |
| Losing the abbreviation choices | Never deleted (D4) |
| Losing document metadata | Columns never deleted (D3) |

---

## 6. Test plan

Pure logic (`src/shared/`):

- the confirmation gate: label match is trimmed and case-insensitive; blank never passes
- "counted zero" and "could not count" are distinct states, and the second withholds the folder option

Site verification:

1. Create a throwaway segment with no folders — Delete, plain confirm, gone from the upload form
2. Re-create it with the same values — it returns over its old folders (D1's reversibility)
3. Delete a segment that has Group Map rows — confirm the rows go and the **groups do not**
4. Reconcile after 3 — confirm the folder grants are now removed
5. Delete `Industrial` with folders ticked — confirm the folders leave both libraries, the Folder
   Map rows go, and the **abbreviation rows and columns remain**
6. Restore one folder from the recycle bin — confirm it returns with its documents
7. One `SegmentDeleted` audit row per run, with accurate counts

---

## 7. Implementation order

1. `shared/segmentDeletion.ts` — the pure rules (confirmation gate, count states) + tests
2. The counting reads
3. The dialog
4. Execution + audit
5. Retitle the tab; package, deploy, run §6

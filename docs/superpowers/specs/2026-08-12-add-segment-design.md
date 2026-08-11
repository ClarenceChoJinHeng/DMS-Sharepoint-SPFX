# Add a Business Segment — slice B of the Structure Manager

**Date:** 2026-08-12
**Status:** designed, not built
**Slice B** of [2026-08-10-structure-manager-ui-design.md](2026-08-10-structure-manager-ui-design.md),
which built slice A (editing the levels of an existing segment) and deferred this.

---

## 1. Why this exists

Onboarding a segment today means hand-editing a `DMS Config` row: a `mode` Title, a Category, a term
set GUID, a StagingFolder, a SortOrder, and a `Levels` JSON chain — then creating the matching columns
in both libraries by hand, under names whose internal form is frozen at creation. The client cannot
edit JSON, and after handover there is nobody to do it for them.

Eight of the twelve segments are still unbuilt
([2026-07-24-twelve-segment-expansion-design.md](2026-07-24-twelve-segment-expansion-design.md)), and
they do **not** reuse the head-office shape:

| Family | Permissioned chain | Columns needed |
|---|---|---|
| Head Office ×4 | `Department → Unit` | none — already exist |
| Upstream Ops ×3 | `Region → Estate/Mill` | `Region`, `Estate` + their `…Tid` |
| SDGI ×2 | per client | per client |
| I&T ×2 | `I&T Operating Unit` — **one tier** | `ITOperatingUnit` + `…Tid` |
| Group-led Project ×1 | per client, `Category = Project` | per client |

So the form cannot offer a fixed `Department / Unit` prefix. **It has to let the admin name the
permissioned tiers and create their columns**, or it cannot onboard the segments it exists for.

## 2. What the form collects

| Field | Written to | Notes |
|---|---|---|
| Segment name | `ModeLabel` | What uploaders see in the Segment dropdown |
| Family | `Category` | `BusinessSegment` or `Project` — drives which top-level tab it appears under |
| Term set | `TermSetGuid` | Pasted ID, validated live (name + count shown back), as in slice A |
| Top folder name | `StagingFolder` | The folder reconciliation creates, e.g. `UPOPS` |
| Permissioned tiers | `Levels` prefix | Named by the admin; ACL'd; columns created |
| Below-Unit tiers | `Levels` suffix | Same editor as slice A, seeded with Year → Document Type |
| — | `Title` | **Derived**, not typed: `mode_<slug of label>`, uniqueness checked |
| — | `SortOrder` | **Derived**: max existing + 1 |

`Title` is derived for the same reason column internal names are: it is a key, it is permanent, and
offering an untrained admin a free-text key invites a duplicate or a typo whose only symptom is a
segment that never appears. `SortOrder` is derived because its only job is ordering a dropdown.

## 3. The validation that earns its place

**Term-set depth must equal the number of permissioned tiers.** This is the one check without which
the whole segment mis-provisions, and it is not guessable from the UI.

Reconciliation walks the segment's TERM TREE, not `Levels`, and caps its depth at the permissioned
tier count (`permissionedDepth` in `FolderManager.tsx`). So if the admin names two tiers
(`Region → Estate`) against a term set three levels deep, reconciliation stops one level early: the
folders it ACLs are Regions, the Group Map rows point at Estates, and every unit-level grant lands on
the wrong folder. Nothing errors — it provisions the wrong tree, and the failure surfaces weeks later
as "this person can see too much".

So the form walks the term set, measures its depth, and **refuses on mismatch**, naming both numbers:

> This term set is **3** levels deep, but you have named **2** permissioned levels. Add a level, or
> point at a different term set.

Cost: one request per term while measuring. A 14-department set is ~15 requests at save time, which is
acceptable for an action performed a handful of times per site.

Also validated:

- **Term set exists** (live, as slice A) — and a TERM's GUID pasted instead of a set's 404s, which is
  the outcome we want
- **`StagingFolder`** is a usable folder name after sanitizing, and is **not already used** by another
  mode row — two segments sharing a top folder would merge two segments' trees under one ACL
- **Segment name** is unique among existing rows, and its derived `Title` is not taken
- **The chain passes `validateChain`** — permissioned tiers contiguous, no duplicate columns

## 4. What it writes, and in what order

Same order and the same reasoning as slice A: a column with no row is a harmless orphan, a row naming
a missing column breaks every upload in that segment, and `validateUpdateListItem` returns HTTP 200
with `HasException` (gotcha #4) so that failure is not even loud.

1. **Validate** everything above. Nothing is written until all of it passes.
2. **Create the tier columns** in `Approval Document` **and** `Documents` — `labelCol` + `tidCol` per
   tier, via `CreateFieldAsXml` with `Options: 8`, created under the space-free internal name then
   retitled. Idempotent: a tier named `Department` reuses the existing column rather than making a
   second one, which is exactly what the four head offices need.
3. **Create the `mode` row** with every field from §2.
4. **Report** the row, the columns created per library, and then §5's checklist.

A failure at step 2 on the second library stops the run and says so: a column present in one library
and absent from the other routes files correctly and loses their metadata in one library only — the
hardest of these states to notice.

## 5. The form ENDS on a checklist, not a success message

A mode row is one of six things a segment needs (slice A §5.1). Getting this wrong is the difference
between a client who finishes the job and one who believes they have:

| Step | Automated here? |
|---|---|
| `DMS Config` mode row | **YES** |
| Tier columns in both libraries | **YES** |
| Term set + its terms | No — term store, and it must exist BEFORE this form |
| **An abbreviation row for every term** | No — **a term with no abbreviation gets NO folder** |
| Security groups per unit + Group Map rows | No — Folder Access page |
| Folder Reconciliation run | No — button on the Folder Manager page |

The abbreviation line is the one that silently does nothing: reconciliation skips a term with no
abbreviation rather than guessing, so a fully configured segment with none provisions an empty tree
and reports success. It belongs on the screen, not only in a document.

The new segment is **immediately visible in the upload form** — it is a mode row, and the form reads
those live. So it will be offered to uploaders before its folders exist. That is acceptable only
because an upload into it fails loudly (no Folder Map row → *"has no folder yet… needs an abbreviation
then reconciliation"*), and because the checklist says so. **Do not** add a "draft" state to hide it:
that would be a second staging mechanism, and the one we have exists for a different problem.

## 6. Out of scope, deliberately

- **Creating the term set or its terms.** Typing 14 departments and their units belongs in the term
  store, and term sets are shared across segments — the argument in
  [2026-08-11-subtree-migration-design.md](2026-08-11-subtree-migration-design.md) against deleting
  them from this page applies to creating them too: this screen should not be where the taxonomy is
  authored.
- **Abbreviation rows.** Authored data — the client's codes, not derivable. Reconciliation already
  reports every term missing one.
- **Groups and Group Map rows.** That is the Folder Access page, and the backlog item that makes it
  grant on save.
- **Deleting a segment.** No safe meaning: the folders hold documents, the columns hold history, and
  the term set is shared. To retire one, stop offering it — which today means removing its mode row by
  hand, deliberately.
- **Reordering segments.** `SortOrder` is derived; if the client wants a different order, that is a
  list edit.

## 7. Errors

| Condition | Behaviour |
|---|---|
| Term set ID malformed / not found | Refuse, as slice A |
| Term-set depth ≠ permissioned tier count | **Refuse**, naming both numbers (§3) |
| `StagingFolder` already used by another mode row | Refuse, naming the other segment |
| Segment name or derived `Title` already exists | Refuse, naming the clash |
| Chain fails `validateChain` | Refuse, quoting the rule |
| Column exists with a different type | Refuse; never attempt conversion. Name the column and library |
| Columns created but the row write fails | Report the orphan columns explicitly — harmless, but must not be silent |

## 8. Testing

- A new segment with **two** permissioned tiers (`Region → Estate`) against a 2-deep term set: columns
  created in both libraries, row correct, segment appears in the upload form, upload fails with the
  abbreviation message until reconciliation has run.
- A new segment with **one** permissioned tier (I&T): unit folders are direct children of the segment
  folder, and the migrator's `depth` maths still holds.
- A **depth mismatch** both ways: 2 tiers against a 3-deep set, and 3 against a 2-deep set — refused,
  with both numbers named.
- A tier named `Department`: reuses the existing column, creates nothing, and the head offices remain
  untouched.
- A duplicate `StagingFolder`: refused before anything is written.
- `Category = Project`: appears under the Project tab in the upload form, not Business Segment.
- After reconciliation: folders exist, a term with no abbreviation is reported and skipped rather than
  guessed, and an upload lands at the full chain depth.

## 9. Related

- [2026-08-10-structure-manager-ui-design.md](2026-08-10-structure-manager-ui-design.md) — slice A,
  whose level editor and column creation this reuses
- [2026-07-24-twelve-segment-expansion-design.md](2026-07-24-twelve-segment-expansion-design.md) —
  which segments exist and their intended chains
- [2026-07-21-segment-onboarding-plan.md](2026-07-21-segment-onboarding-plan.md) — the manual runbook
  whose first two steps this replaces

# Group creation: personas, derived names, and bulk provisioning

**Date:** 2026-08-18
**Status:** Agreed in conversation. NOT built.
**Revises:** `2026-08-14-group-management-separation-design.md` (partially reverses it — see §7),
`2026-07-22-role-from-group-name-suffix.md`, `2026-08-17-hc-clearance-and-role-revision-design.md`.

---

## 1. What is wrong today

Provisioning the client's site needs a group for every persona in every unit — **132 units** (GHO 63,
Minamas 49, NBPOL 20). Group Management creates them one at a time, and Folder Access then maps each one
separately: roughly 660 creations and 660 mappings by hand.

Three faults, in the order they bite:

**1. The name is free text, and the name is load-bearing.** Clarence, 2026-08-18: *"client doesn't care
about naming convention and they will simple type it in for no reason"*. Folder Access recovers a group's
role by **parsing the suffix back** out of its name (`roleFromGroupName`). A typo does not itself grant
anything — the authoritative value is the `Role` column on the Group Map row — but it pre-selects the
wrong role, and the admin accepts a pre-selection that looks right. The mistake is one click away, not
zero. And `GHO_GF_TAX_HC_UPLOADER`, with the suffix in the middle, parses as a **plain uploader**.

**2. The screen asks for the same facts twice.** Group Management's cascade already collects segment,
tier and persona — purely to build a name — then discards them, and Folder Access asks for all three
again. The parse in fault 1 exists *only* because the first screen never wrote them down.

**3. The suggested name matched nothing until 1.0.143.0.** It used full term labels, so
`Group Head Office_Group Finance_…` lined up neither with the folder (named from abbreviation codes) nor
with `FolderAdmin`'s `groupsExist` check (which looks for the segment's `StagingFolder` prefix). Fixed;
recorded because it is the same root cause — a name assembled from the wrong source.

---

## 2. The five personas per unit

Confirmed by Clarence, 2026-08-18. For `GHO / Group Finance / Tax`:

| Persona | Naming role | Group name |
|---|---|---|
| PIC | `UPL` | `GHO_GF_TAX_UPLOADER` |
| Head of Unit | `APR` | `GHO_GF_TAX_APPROVER` |
| SDG Employee | `MEMBER` | **`GHO_GF_TAX_EMPLOYEE`** |
| PIC — HC cleared | `UPLHC` | `GHO_GF_TAX_UPL_HIGHLY_CONFIDENTIAL` |
| SDG Employee — HC cleared | `MEMBERHC` | `GHO_GF_TAX_VIEWER_HIGHLY_CONFIDENTIAL` |

**A person belongs to exactly ONE of these.** `UPLHC` and `MEMBERHC` are supersets — they appear on all
four libraries — so an HC-cleared PIC needs no plain `_UPLOADER` group as well. Putting them in both is
not dangerous, just two grants saying one thing and one more row to unpick when they leave.

**Head of Unit, Head of Department and C-Level need no HC group at all**; they reach HC through roles they
already hold. Worth repeating to the client because it reads oddly: *a Head of Unit needs no clearance, an
SDG Employee does.*

### 2.1 `_EMPLOYEE` is a new suffix, and it closes a hole

`MEMBER` has carried **no suffix** — `suggestGroupName` returned a bare `GHO_GF_TAX`. Clarence asked for
`_EMPLOYEE`, and it is more than cosmetic: `roleFromGroupName` **falls through to `MEMBER`** for any name
it does not recognise, so a bare name and an unrecognisable name are currently the same answer. That is
precisely how a mistyped approver group reads as view-only.

Once every generated name carries a suffix, an unrecognised name *can* be treated as unknown rather than
silently becoming a viewer. **That fallback change is NOT in this spec** — it would reclassify groups on
existing sites — but adding `_EMPLOYEE` is what makes it possible later.

Additive and safe: verified against the live table, there is no `_EMPLOYEE` and no `_MEMBER` suffix, so
nothing collides, and existing bare-named base groups keep resolving through the fallback.

**`MEMBERHC` keeps `_VIEWER_HIGHLY_CONFIDENTIAL`.** Its original reason — that `MEMBER` had no suffix to
extend — is now obsolete, but a suffix is a public name once shipped, and renaming it would strand groups
already created.

### 2.2 A persona needs an EXPLICIT naming role

Not "its first role". `PERSONAS` order gives the right answer for five personas and the wrong one for
`employee_hc` = `["MEMBER", "MEMBERHC"]`, whose first role is `MEMBER` — so it would be named
`GHO_GF_TAX_EMPLOYEE`, parse back as plain `MEMBER`, and present an **HC-cleared viewer group as having no
clearance**. Silent, and exactly what the suffix table exists to prevent.

So: one declared naming role per persona, pinned by a test that asserts every persona has one and that
each round-trips through `roleFromGroupName` back to itself.

---

## 3. The name becomes an output

```
Segment      [ Group Head Office ▼ ]
Department   [ Group Finance ▼ ]
Unit         [ Tax ▼ ]
Persona      [ PIC — HC cleared ▼ ]

Group name   GHO_GF_TAX_UPL_HIGHLY_CONFIDENTIAL          (derived · read-only)
```

- **Persona, not role.** Folder Access already dropped its role chips for this reason; Group Management
  kept them, which is the same trap in a second place. Nobody should have to choose between *"Delete
  pending files"* and *"View only — whole department"*.
- **Read-only, derived from the dropdowns.** Same derivation as the bulk run, so a bulk name and a
  single-creation name cannot differ.
- **Free text stays, behind an explicit `Type the name myself (advanced)` toggle.** It cannot be removed:
  `CRS_SITE_MEMBERS` follows no convention, and C-Level groups sit at segment scope with no unit. The
  client must opt in, which is the opposite of today's default.
- The cascade may stop at any tier — a HoD group maps at department, C-Level at segment.

---

## 4. Creating a group writes its Group Map rows

**The substantive change.** The dropdowns already collect everything the row needs, so creation writes it:

| Column | Value |
|---|---|
| `GroupId` | the new SP group's **integer** id |
| `GroupName` | its title (a cache — see §8) |
| `Segment` | the segment's **term-set GUID** |
| `UnitTermGuid` | the chosen leaf term |
| `Role` | one row **per role in the persona**, `Role` as **Text** |
| `Target` | per `personaTouchesStaging()` |

One persona yields several rows (`hou` = `APR`, `DELS`, `DEL`, `SHARE`, `UPLHC`, `DELSHC`). That is already
how Folder Access writes them; the write moves, it does not change.

**Consequences worth stating:**

- The **parse leaves the critical path**. The role is written from the persona that was chosen, never
  recovered from a string. Fault 1 stops being reachable for anything created here.
- **"Created but not yet assigned" stops being the normal state.** It stays legal — a group made in
  SharePoint directly has no row — but it is no longer where every group begins.
- **Nothing is granted until reconciliation runs.** Unchanged, and the confirmation must keep saying so.

### 4.1 Partial failure

The group is created first, then the rows. If a row write fails the group **remains** — deleting it would
destroy a group that may already have members, and re-running is safe because the group is found by name.
The screen reports *"group created, N of M mappings written"* naming the failures, never a flat success.
Same rule as the audit-row half-success on the abbreviations page.

---

## 5. Bulk provisioning

**Source: the `CRS Term Abbreviation` rows.** A row with a non-blank `Abbreviation` is exactly a term that
will get a folder, which is exactly a term that needs groups. No second list to keep in step.

```
Segment  [ Group Head Office ▼ ]        63 units with a folder code · 0 without

Create for each unit:
  [x] PIC                        [x] PIC — HC cleared
  [x] Head of Unit               [x] SDG Employee — HC cleared
  [x] SDG Employee

Preview (315 groups · 47 already exist · 268 to create)
  GHO_GCA_EG_UPLOADER                         new
  GHO_GCA_EG_APPROVER                         new
  GHO_GF_TAX_EMPLOYEE                         already exists — will be mapped only
  …
                                   [ Export CSV ]  [ Create 268 groups ]
```

- **Preview before anything is written.** 315 groups is not an action to take on trust, and group deletion
  is manual. This is the review step Clarence asked for.
- **Per-persona tick boxes, all five ticked by default** — the client's instruction is that every unit gets
  HC. Unticking is how they run three-per-unit for a segment if the extra groups prove noisy. 660 groups
  across three segments at five each; 396 at three.
- **A unit with no folder code is EXCLUDED and counted**, never silently skipped. No code means no folder,
  so its groups would grant nothing — and a silent skip is exactly how the abbreviation step's own failure
  hides.
- **Idempotent.** An existing group is not re-created; its rows are still asserted, because a group with no
  row is the state that grants nothing.
  - ⚠ **"ASSERTED" WAS BUILT AS "WRITTEN AGAIN", AND THAT IS THE 642-ROW BUG** (found on the rehearsal
    site, fixed in 1.0.151.0 by `splitPlannedRows`). Asserting a row means writing it *if it is not
    already there* — so the run reads every Group Map row first and partitions each group's planned rows
    against them, per ROW rather than per group, because a run stopped part-way leaves a group made with
    only some of its rows written and pressing Run again is the only resume this has.
  - The asymmetry is why nobody caught it: group creation *was* idempotent, so a second press logged
    `= already existed (mapping only)` on every line while doubling the mappings behind them. No screen
    shows a row twice.
  - The read is **paged** (`$top` caps a page, it does not lift the 5,000-item threshold; one segment on
    CRS is ~790 rows) and **fails closed** — an unreadable Group Map holds the run rather than being
    treated as empty, the opposite of this codebase's usual rule and for the usual reason: here the cost
    of guessing is hundreds of rows nobody would ever find.
- **CSV export for the cross-check**, reusing `groupExportCsv.ts`.
- **~660 creations will throttle.** Needs the same retry as reconciliation's group pass, plus a run log — a
  bulk run that dies half-way must say what it made.

**Members are NOT bulk-assigned.** That is the irreducible manual work, and it is per person.

---

## 6. Folder Access is demoted, not deleted

Two jobs still need it, and both are why it survives:

1. **Mapping a group we did not create.** The client renames every group at import
   (`dms-to-crs-rename-pending`), and any group made directly in SharePoint has no row. Without this there
   is no route to map one at all.
2. **Review and export** — the mappings table is where the cross-check happens.

**Member editing is removed from it.** Members belong in exactly one place and Group Management already has
them; two lists of the same people drift apart. (Per-person *removal* stays where the 2026-08-14
per-person-removal design put it.)

---

## 7. Why this does not reintroduce the 2026-08-14 problem

That spec **split** creation from mapping because the client said the combined screen was confusing. The
actual defect was narrower: creating a group was reachable only as half of *"Create group & add mapping"*,
which rolled the group back if any row failed, and its mirror **deleted the SharePoint group when its last
mapping row went** — so "exists but not yet assigned" could not be expressed.

**That auto-delete rule is gone.** This spec recombines *create + map* while keeping *map an existing
group* on its own screen, and §4.1 keeps the group on a row failure. The state that could not be expressed
is still expressible; it is simply no longer the default.

---

## 8. Unchanged, and stated so nobody re-derives it

- **`Role` stays Text, never Choice.** A value absent from a Choice column's `Choices` fails the whole
  write, silently.
- **Grants key on the integer `GroupId`.** `GroupName` is a cache, so a rename desynchronises every screen
  and log line while the grants stay correct (open finding #12). Bulk creation does not change that.
- **`SEGVIEW`/`GLOBAL` must never reach an approval library.** The persona list enforces it; no persona
  offered here contradicts `LIBRARY_ROLES`.
- **Nothing here grants anything.** Reconciliation applies the ACLs, and only for terms that have codes.

---

## 9. Build order

1. Per-persona naming role + `_EMPLOYEE` suffix, with tests (pure, no UI)
2. Group Management: persona dropdown, derived read-only name, advanced free-text toggle
3. Group Management writes the Group Map rows; §4.1 partial-failure reporting
4. Folder Access: remove member editing, keep map-existing + review + export
5. Bulk screen: preview → create → run log → CSV, with throttle retry

Steps 1–3 are usable on their own and remove the mis-parse. Step 5 is the one that saves the days.

---

## 10. Open

- **Does the bulk run also write the rows, or only create the groups?** Written here as *both*, because a
  group with no row grants nothing and the preview plus CSV already provide the cross-check. If the client
  wants a hard gap between "groups exist" and "access granted", the second half becomes its own button —
  the preview does not change either way.
- **Which units genuinely need the HC pair.** Defaulting to all of them is the client's instruction, and an
  empty group grants nothing, so the cost is noise in the people picker rather than exposure.

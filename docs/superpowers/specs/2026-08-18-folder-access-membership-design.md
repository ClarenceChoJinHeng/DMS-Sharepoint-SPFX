# Folder Access owns membership; Group Management owns groups

**Date:** 2026-08-18 · **Client's instruction, in this conversation.**
Supersedes the membership half of `2026-08-14-group-management-separation-design.md`.
Builds on `2026-08-18-group-creation-and-bulk-provisioning-design.md`.

---

## 1. What changed underneath

On 2026-08-14 the group LIFECYCLE left Folder Access, because "this group exists but is not assigned
yet" was a state the old combined screen could not express. Folder Access kept one job: map an existing
group to a segment, tier and role.

**Bulk provisioning removed that job.** Creating a group writes its Group Map rows, and a bulk run
writes 790 of them in one press. So the mapping form is no longer the everyday route — it is the
fallback for a hand-named group and for repairing a row someone deleted. Yet it sat at the top of the
page, above the table, which states that mapping is manual work done here. It is not.

Meanwhile the one genuinely recurring job in the whole system — **putting a person into their unit's
group** — was on Group Management, a page an admin otherwise visits twice in a segment's life.

So the split is re-drawn along what each page is *for*, not along what the code used to do:

| | Group Management | Folder Access |
|---|---|---|
| Answers | what groups exist | who is in them, and what they reach |
| Frequency | twice per segment | weekly, forever |
| Does | create (bulk or one), delete | add and remove people; read the mappings |
| Persona | **chosen here**, at creation | shown, never chosen |

**The persona cannot move to Folder Access, and that is structural, not a preference.** A persona
decides the group's NAME (`GHO_GF_TAX_APPROVER` comes from picking Head of Unit), so it has to be known
*before* the group exists. Anything Folder Access does is necessarily after the fact.

## 2. Folder Access

### 2.1 One row per GROUP, not per mapping

The mappings table collapses from 790 flat rows to **308 expandable group rows**. This is forced, not
cosmetic: membership is a property of the GROUP, and a unit's approver group has six mapping rows — so
a per-mapping member editor would show the same people thirteen times per unit, with thirteen Add
boxes that all did the same thing.

Expanding a group shows its mappings (segment · Tier 1 · Tier 2 · role, with Delete per row) and its
**people** (add, remove). The grouping is a pure function, `groupMappingsByGroup`, so the shape the
screen renders is testable without a tenant.

- **Rows keep their per-mapping Delete.** Deleting a mapping and removing a person are different acts
  with different consequences, and collapsing them into one control is how an admin revokes more than
  they meant to.
- **The group row shows its mapping count**, so a group with 13 reads as provisioned and one with 1
  reads as a `_HOD` — the distinction the Tier 1 / Tier 2 columns were added to make.
- **Scroll capped**, as the group list on Group Management is, and for the same reason: 308 rows push
  everything else off the screen. Released while a group is expanded, because the people picker inside
  it is absolutely positioned and a scroll container clips it.

### 2.2 Adding a mapping by hand LEAVES the page entirely

**Corrected 2026-08-18, same day: the client said twice that Folder Access must not create mappings,
and a closed disclosure on the page was still on the page.** The form now renders only on **Group
Management**, collapsed, as *"Map a group by hand — rarely needed"*.

Group Management is the right home rather than a compromise: both cases that need the form originate
there. A group created with a free-typed name and no persona is created there, and a group that must
cover a second tier is created there. Mapping either from the page where it was made is the shorter
path anyway.

Mechanically it is a `show` prop — `"members"` for Folder Access, `"form"` for Group Management —
**two mount points of one component, never a copy.** The form and the list share `existing`,
`postRow` and `isDuplicateRow`; separate components would give the site two definitions of a mapping
row, and the one that drifted would be the rarely-used one.

Deliberately NOT added to the guided flow's Group Management step. That step exists to create groups;
hand-mapping is rare enough to live on the standalone page, and mounting this component there would
cost a second 790-row read on a step that does not need it.

**Deleting it outright was considered and rejected**, because two narrow cases have no other route:

1. A group created with the **advanced free-text name and no persona** gets no rows —
   `rowsForNewGroup` returns `[]` deliberately. Mapping it later would otherwise mean deleting and
   re-creating the group, which loses its members.
2. **One group covering two tiers.** Not a stated requirement, and nothing else can express it.

A disclosure satisfies the actual complaint — the client never meets it — without stranding anyone.
The library toggle and the persona picker go inside it, since they exist only to build a row.

### 2.3 Membership is a GROUP change, and the page says so

Unchanged from the 2026-08-14 per-person removal work, and it has to be repeated here because the
page's new framing invites the wrong reading: a group's grants are held at folder scope, so adding
someone to `GHO_GF_TAX_UPLOADER` gives them that unit's folder in every library that group is mapped
to — not "access to this page's list". The intro says it; the toasts say it.

`addMemberWithSiteEntry` is what does the write, so a person also lands in the site-entry group. Its
`note` is never dropped: without site entry they ARE in the group and cannot open the site, which
neither they nor the admin discovers until they try.

## 3. Group Management loses the member editor

Not hidden — **moved**. Two lists of the same people drift, and re-merging the two pages is precisely
what the client called confusing on 2026-08-14. `GroupMembersEditor` is extracted to its own file and
mounted by Folder Access only; Group Management keeps the group list, the mapping-count badges,
create and delete.

**Accepted cost:** after creating a group you change page to put the first person in it. That is the
price of one editor rather than two, and the client accepted it.

The `not mapped` badge stays on Group Management. It answers "did creation write the rows", which is
that page's own question.

## 4. Folder Access is never a required step

Client's instruction: *"if they don't know who yet to add then we can allow them to add later."*

- The flow step already does not gate Next (`NEXT_GATED_STEPS` holds `createSegment` and
  `abbreviations` only), and must not be added to it. Membership is **intent**, and intent is not
  checkable: nothing can tell whether the RIGHT people are in a group.
- The step's copy now says so, and names the standalone page as where to do it later.
- **Reconciliation does not care.** It grants to a group, not to its members, so an empty group is a
  valid end state: the grant is in place and works the moment someone is added. Nothing needs re-running.

## 5. What is deliberately NOT built

- **No bulk member import.** The client's own document assigns people per unit by hand; a CSV import
  would need a name-to-login resolution step whose failure mode is silently granting the wrong person.
- **No member management on the other access pages.** Site Access, Page Access and Approval Library
  Access keep their existing per-person REMOVAL (`accessMemberUi.tsx`) and gain no Add. Adding a person
  there would look scoped to that library or page and never is.

# HC clearance for the Head of Unit — `hou_hc` (2026-08-24)

**Status:** BUILT (1.0.234.0), not yet provisioned on any site.
**⚠ THIS DOCUMENT WAS REWRITTEN THE SAME DAY IT WAS FIRST WRITTEN.** The original version described a
softer rule — clearance gates filing only, an uncleared Head of Unit still sees and approves HC
work. That was accurate for about an hour. The client then clarified further, and the design below is
what actually shipped. **CLAUDE.md's `hou_hc` section is the current summary; read that first.**

**Client instructions, verbatim, in order:**
1. *"we will need to make another highly confidential group for HOU, so not every HOU can upload into
   Highly Confidential, it will be the same pattern as PIC."* (upload-only split — the original design)
2. Asked to confirm scope, the client settled it narrower: *"a normal HOU cannot see HC Approval
   Document and HC Documents … only HC HOU can see and approve and go into HC libraries."*

Instruction 2 governs. It reverses the 2026-08-17 rule
(`2026-08-17-hc-clearance-and-role-revision-design.md`) in full — that rule's whole premise, *"any
approver which is HOU can see Highly Confidential files as well… we do not need a dedicated HOU,"* is
now false.

---

## 1. The rule, in one sentence

**No plain Head of Unit reaches either HC library, for any reason.** Seeing, approving, deleting and
sharing HC documents all require the `hou_hc` group — there is no partial access.

| Persona | Roles | Named |
|---|---|---|
| `hou` | `APR, DELS, DEL, SHARE, UPL` | `…_APPROVER` |
| `hou_hc` | `APRHC, DELS, DEL, SHARE, UPLHC, DELSHC, DELHC, SHAREHC` | `…_APR_HIGHLY_CONFIDENTIAL` |

`hou_hc` is a strict superset of `hou`'s ORDINARY powers (it still approves, deletes and shares
normal documents through `DELS`/`DEL`/`SHARE`) plus every HC role. A person gets ONE of the two
groups, never both — giving someone both is redundant, not additive.

## 2. Why APR/DEL/SHARE had to LEAVE the HC library rows

The plain Head of Unit holds `APR`, `DEL` and `SHARE` for their normal, non-HC work. A role held by
two personas cannot grant to one and withhold from the other — this codebase's recurring rule,
learned first with `DELS` (2026-08-17) and `SEGVIEW`/`GLOBAL` before it. So as long as `APR` sat on
`StagingHC`/`DocumentsHC`, EVERY Head of Unit could see HC through it, whether cleared or not.

The fix is three HC-specific twins, added to `LIBRARY_ROLES`' HC rows in place of the plain roles:

- **`APRHC`** replaces `APR` on both `StagingHC` and `DocumentsHC`. A superset exactly like `UPLHC` —
  it approves the NORMAL approval library too, so `hou_hc` carries no separate `APR` row.
- **`DELHC`** replaces `DEL` on `DocumentsHC` — delete an approved HC document.
- **`SHAREHC`** replaces `SHARE` on `DocumentsHC` — share an approved HC document.
- **`DELSHC`** (pending-HC delete) is unchanged in shape, just narrowed to `hou_hc` alone — it was
  already its own role for exactly this reason (2026-08-17: *"a role held by two personas..."*).

## 3. The staffing rule this creates — say this to the client plainly

**A unit that files HC documents MUST have someone in its `_APR_HIGHLY_CONFIDENTIAL` group.** Without
one, every HC document that unit's PIC uploads sits pending forever — visible only to its author,
because nobody else in the unit (including its plain Head of Unit) can even open the HC approval
library to see it. There is no error, no warning, no timeout. It simply never gets approved.

This is the direct, accepted consequence of instruction 2 — not an oversight. The 2026-08-17 rule
existed specifically to avoid this deadlock (*"we do not need a dedicated HOU"*); the client chose to
re-introduce it in exchange for stricter isolation.

## 4. Head of Department and C-Level keep their HC oversight

Confirmed with the client directly (2026-08-24): management oversight of HC is NOT what "only HC HOU
can see and approve" was about — that sentence is about the Head-of-**Unit** tier specifically.

- **HoD** still reads every unit's approved documents in their department, `DocumentsHC` included
  (`DEPTVIEW`, unchanged), and still deletes/shares them department-wide — now through the new
  **`DELHC`/`SHAREHC`** roles, added to the `hod` persona alongside its existing `DEL`/`SHARE`.
- **C-Level** (`GLOBAL`/`SEGVIEW`) still reads `DocumentsHC` globally or per-segment, unchanged —
  those roles were never on either HC approval-library row and are unaffected by this change.

Neither HoD nor C-Level ever held anything on `StagingHC` (the HC **approval** library) — that was
never in scope, and remains a `hou_hc`-only door.

## 5. APRHC: revived as a real role, not the naming-only shim first tried

An earlier same-day draft made `APRHC` a *naming-only* role — present only so `hou_hc` could derive
a distinct group name (`_APR_HIGHLY_CONFIDENTIAL`), with no entry in any persona's `roles` or in
`LIBRARY_ROLES`. That was correct for instruction 1 alone. Instruction 2 requires `APRHC` to be a
real, granting role, so it is now wired everywhere `APR` is:

- `LIBRARY_ROLES.Staging` / `.Documents` (a superset, like `APR`)
- `LIBRARY_ROLES.StagingHC` / `.DocumentsHC` (replacing `APR`)
- `ROLE_TO_PERMISSION` (→ the site's Approve level, downgraded to Read on the approved side)
- `STAGING_FACING_ROLES`
- `DOCUMENTS_READ_ONLY_ROLES` (downgrades on the approved side, same as `APR`)
- Page policies: `Upload-Form.aspx`, `ApprovalDocument.aspx`, `My-Submissions.aspx`,
  `Bulk-Upload.aspx`, `Requests.aspx` — all list `APRHC` beside `APR`
- Request routing (`Requests.tsx`, `MySubmissions.tsx`) — both now match `APR` **or** `APRHC`

**The `ROLE_ALIASES.APRHC → APR` rewrite from the 2026-08-17 retirement is REMOVED.** A stored
`APRHC` value is now a real, distinct grant and must not be silently downgraded to `APR`.

## 6. ⚠ The fingerprint consequence — read before running bulk provisioning

Every `_APPROVER` group provisioned under the 2026-08-17 shape (`APR, DELS, DEL, SHARE, UPL, UPLHC`)
now matches **neither** persona exactly — `hou` no longer includes `UPLHC`, and `hou_hc` no longer
includes `APR`. `personaForRoles` reports such a group unmatched, which is accurate: it is neither.

**Running bulk provisioning without deleting the old groups first still makes a mess** — the name
check finds the old `_APPROVER` group and writes new rows onto it rather than creating fresh groups,
and the old group's HC grant is untouched either way.

## 7. Migration — per segment, on BOTH sites

1. **Export the Group Management CSV first** — it carries each group's members, and deleting a group
   loses its membership.
2. **Delete every `_APPROVER` group** for the segment (Group Management). Deleting the group is what
   actually withdraws the SharePoint grants — a removed Group Map row alone does not (2026-08-20).
3. **Run Bulk provisioning** for the segment. It creates a narrow `_APPROVER` group (no HC access at
   all) AND an `_APR_HIGHLY_CONFIDENTIAL` group per unit (empty — a valid end state), with their rows.
4. **Re-run reconciliation** for the segment. Incremental: the grant-skip (1.0.177.0) makes this far
   cheaper than a first run.
5. **Re-add the people from the CSV**: everyone who should have HC access → `_APR_HIGHLY_CONFIDENTIAL`
   only. Everyone else → `_APPROVER`.

**`hod`'s migration is ADDITIVE, not destructive.** Its group keeps its name and its existing
`DEPTVIEW`/`DEL`/`SHARE` rows; re-running bulk provisioning maps the two new `DELHC`/`SHAREHC` rows
onto the SAME group (the name check finds it and merges), then reconciliation grants them. No
deletion, no membership loss.

Until a unit's `_APPROVER` group is re-created, its Head of Unit **keeps full HC access** — the
widening this change removes stays live per unit until that unit is migrated.

## 8. What needed no change

- The upload form's HC gate — it probes the HC folder for write access directly, so it answers
  correctly for whichever group a person is in, with no code change.
- `collectMembership` — it counts UPLOAD roles (`UPL`/`UPLHC`); `hou_hc` still uploads through
  `UPLHC` exactly as before, unaffected by the approve/delete/share changes.

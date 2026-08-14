# Per-person access removal — Approval Library Access and Page Access

**Date:** 2026-08-14
**Status:** Built, not yet site-tested
**Screens:** `Approval Document Access` (`StagingAccess.tsx`), `Page Access` (`PageAccess.tsx`)
**Shared rules:** `src/shared/accessMembers.ts` (pure, 28 tests)
**Shared UI:** `src/webparts/userAccess/components/accessMemberUi.tsx`
**Supplements:** `2026-08-03-access-scope-mapping-design.md` (which designs both screens as
group-only) and `2026-08-14-group-management-separation-design.md` (the group lifecycle).

---

## 1. What the client said

> "Client is saying the Approval Library Access and Page Access doesn't make sense in terms of user
> experience.
>
> If client wants to remove someone from staging access it should be removing user and not group,
> but ofcourse if possilbe let them remove the group and user, so they can choose to remove user
> individually from the group itself or remove an entire group.
>
> Same goes for Page access."

They are right, and the complaint is precise: both screens were **group-only**. `StagingAccess`
listed uploader/approver/deleter groups with Allow and Remove and showed no people at all;
`PageAccess` showed members in a read-only column with a "View" modal, which could answer *who is in
here* but never *take this one person out*. So "stop this person opening the approval library" had
exactly one lever — remove their whole group — which hits everyone else in it too.

## 2. The constraint that shapes the answer

**Access at library and page scope is granted to a GROUP.** A person has it because they are in that
group. So per-person removal can only mean *removing them from the group* — and that group is the
same `*_UPL` / `*_APR` group reconciliation maps at **folder** scope.

Therefore **this is not a library-scoped or page-scoped action**, however the button is labelled.
Removing someone from `GHO_GF_CORU_UPL` also takes away their upload permission on their unit
folder. Every screen states this before the removal runs, and the audit row states it after.

The rejected alternative was **individual direct grants** — a role assignment on a *user* principal
on the library or page. That would be genuinely narrow, and was declined: it manufactures exactly
the state both screens already flag in red ("granted directly in SharePoint… not managed by this
tab, reconciliation will not remove them"), with no Group Map row explaining why it exists and
reconciliation blind to it. Deciding against it is a design choice, not a limitation — it is
achievable, and should not be re-proposed without that trade being re-argued.

## 3. Three facts the UI must carry, because an admin cannot check them

Each has a plausible wrong version that reads as correct on screen. All three live in
`accessMembers.ts` so both screens answer them identically.

1. **A person in TWO allowed groups keeps access after one removal.** Remove them from `..._UPL`
   while they are also in `..._APR` and they still get in. A per-row Remove with no cross-group view
   makes a *correct* removal look like a *failed* one — so the other allowed groups are named on the
   member row itself (`also in GHO_GF_CORU_APR`), before the click, not only in the dialog. By the
   dialog the admin has already decided which row to act on.
2. **Removing the last member does NOT remove the grant.** The group stays mapped and allowed, just
   empty — which reads as configured. Called out in the dialog, and the collapsed row shows
   `no members` in amber.
3. **An unreadable member list is not an empty one.** This was advisory while the column was
   decoration; it is load-bearing now a removal is decided from it. `MemberLoad` is a three-state
   union (`loading` / `loaded` / `error`), and `memberCountLabel` says `could not read members`
   rather than `no members`. An admin told a group is empty **stops looking for the person they came
   to remove**.

`removalVerdict` returns **three** answers, never a boolean: `ends`, `survives`, `unknown`. The third
exists because asserting "this ends their access" from a partially-read picture is the failure worth
avoiding, and the unread half is invisible — a group whose members failed to load renders as a
collapsed row like any other. `survives` beats `unknown` when both apply: a known survivor is the
more actionable fact.

Note the polarity: this is the usual **empty ≠ unknown** rule of this codebase, and here it does
*not* fail open. An unreadable group makes the verdict `unknown`, which still permits the removal —
what it refuses to do is *claim* the removal was sufficient. `lastMember` likewise stays false on an
unreadable list: understating is the safe direction for a sentence about what is left behind.

## 4. What was built

**One shared module and one shared UI file, mounted by both screens.** The rules are the ones an
admin cannot verify themselves, so two copies would eventually disagree about a permission change
while each screen stayed individually convincing.

- `MemberSummary` — the collapsed cell: a chevron plus `3 members` / `no members` (amber) /
  `could not read members` (red).
- `MemberRows` — the expansion, as a single `<tr>` with a `colSpan` cell so it inherits the parent
  table's columns. A nested table would misalign the moment either screen gains a column, and the
  two screens do not have the same columns (7 vs 6).
- `MemberRemovalDialog` — leads with the verdict, then the caller's `scopeWarning` (the wider
  effect differs per screen and is not derivable in the shared file), then the last-member note.
- `useGroupMembers` — sequential fetch, one group at a time, for the reason the original PageAccess
  version gave: ten simultaneous calls is how a tenant starts returning 429s. `refresh(groupId)`
  re-reads **one** group after a removal; a full reload would re-fetch every group to show one name
  disappearing and would lose which rows the admin had expanded. A group that fails is **not**
  retried on every set change — that would hammer a group failing for a durable reason (deleted, no
  rights); the row carries an explicit "Try again".

**Expansion is offered only on groups that actually grant access here** (`row !== undefined`). A
group with no mapping grants nobody anything at this scope, so its member list has nothing to remove
*from* — the action on that row is Allow. The **count** still shows on every row, because "this group
has no members" is worth knowing one click *before* allowing it.

**A set of expanded ids, not one.** Comparing two units' member lists is why an admin opens this; an
accordion that closes the previous row makes that impossible.

**`pendingUsers` is a separate namespace from `pending`.** The latter holds *group* ids. One shared
array would spin a group's button because a person inside it was being removed.

**`allowedGroupIds` is derived from the Group Map rows, and `titleOf` from the live group list** — so
a row naming a group that has since been deleted contributes no title, and is therefore never listed
as a reason access survives.

### Page Access only

- **An unrestricted page makes the whole control inert.** Everyone with site access can open it, so
  removing someone from a group changes nothing about who can. Said on the expanded row
  (`inertNote`), in the dialog (`inertWarning`), in the toast and in the audit row — four places,
  because this is the case where the control looks like it worked and did nothing.
- **The read-only members modal was removed, not kept alongside.** It showed the same names; keeping
  both would be two lists of the same people that can disagree after a write. That is the drift the
  rejected "Both" layout option would have introduced.
- **This screen had no audit trail at all** — group grants and revokes on a page went unrecorded.
  Added with this change rather than separately: removing one named person's access is exactly the
  change somebody asks about weeks later, and "who removed them" cannot be reconstructed from the
  page's ACL, which only shows who is left.

## 5. Audit

`EVENT.membersChanged` (already exists; `EventType` is a **Text** column, so a new value cannot fail
a write). One row per removal, and the row records **what actually happened**, not what was asked
for:

- `Outcome: "Success"` only when the removal genuinely ended access — i.e. verdict `ends`, and on
  Page Access only when the page is actually restricted.
- `Outcome: "Failed"` when it survived, was unconfirmed, was inert, or the call errored. A removal
  that leaves the access in place is not a clean outcome, and a row reading only "member removed"
  would stop someone looking. Same principle as the half-applied revoke already recorded here.
- Details always name the wider effect (the group also grants folder permissions) and, where
  relevant, the surviving or unreadable groups **by name**.

## 6. A bug found while wiring this in — `getbytitle('Staging')`

`StagingAccess` receives `library="Staging"`, which is the **logical `LibTarget` key**: it is the
stored `Target` value on Group Map library-scope rows and must keep being written and filtered on
unchanged. But `listBase` built its URL straight from it, so every live-ACL call asked for a library
titled `Staging` — which does not exist since the library was recreated as `Approval Document`.

Consequences, all quiet: **"Access now" read `unknown` on every row**, `grantLive` and `revokeLive`
could never apply a permission, and the warning banner said "Could not read the current permissions
of Staging". `ApprovalLibraryAccessPage` already *claimed* `libApiTitle` did this translation at the
API boundary — it just never happened, because `libApiTitle` was a **private const inside
`FolderManager.tsx`**.

Fixed by moving `libApiTitle` to `shared/naming.ts` (one copy, imported by both) and using it for the
URL only. This is gotcha #12 again: **translate at the API boundary, never in stored data.** The
`Target` filter and the row written by `postRow` still use the logical key.

The same screen was also showing the retired name `Staging` in eleven user-facing strings, its
dialogs and its audit rows; those now use the live title (`libLabel`). The page heading already
resolved it, which is what made the mismatch visible.

## 7. Deliberately not done

- **No individual direct grants** (§2).
- **No flat "people with access" panel.** Considered and declined: it answers "who can get in"
  directly, but alongside the inline expansion it is a second list of the same names that can drift
  after a write. The `also in …` label on each member row carries the cross-group fact that panel
  would have existed for.
- **No bulk person removal.** The group-level bulk Remove already covers "everyone in this group",
  and a multi-select over people spanning several groups cannot state a single honest verdict.
- **No adding people.** That is the Group Management page's job, deliberately, per the separation
  agreed earlier the same day. These screens remove; they do not become a second member editor.
- **No change to who can reach these screens.** Both remain admin screens under the existing page
  access policy.
- **Nothing about `Documents`.** The unit folder remains the smallest confidentiality boundary; this
  changes how access is *administered*, not what the model permits.

## 8. Testing

Site test, on the Approval Document Access screen first:

1. A mapped group with several members — expand it. Count matches, names sorted, emails shown.
2. Someone in **two** allowed groups — confirm the `also in …` label on the row, then confirm the
   dialog says access **survives** and names the other group. Remove; the toast must say they still
   have access.
3. Someone in **one** allowed group — dialog says access ends; toast says they can no longer open
   Approval Document.
4. A group with **one** member — dialog carries the last-member note; after removal the row reads
   `no members` in amber and the group is still mapped and still granted.
5. **"Access now" must no longer read `unknown`** on every row (§6), and no visible text should say
   "Staging".
6. Page Access: repeat 1–4 on a **restricted** page, then on an **unrestricted** one — the second
   must warn that removal changes nothing about who can open it.
7. Audit log: one `MembersChanged` row per removal, `Outcome` matching the verdict, details naming
   the folder-permission consequence.

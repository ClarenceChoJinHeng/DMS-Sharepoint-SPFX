# Quick Search: a Remove button per group

2026-09-13. Client, looking at a Quick Search result listing `PCAR_EB_BMW_APPROVER`: *"add a remove
button for the group the person is in, makes it easier to remove someone."*

## The ask

Quick Search (Group Management's "What can this person reach?" panel) already answers *"which
groups is this person in, and what do they grant"* — it has no way to act on the answer. Removing
someone today means leaving Quick Search, finding the same group in the group list below, expanding
it, and finding them again in `GroupMembersEditor`. This adds a **Remove** button directly on each
group row Quick Search already renders, so the answer and the fix sit in one place.

## What this is not

**Not a second implementation of group removal.** `GroupMembersEditor`'s `onRemove` already exists,
already handles the two genuine hazards (site-entry asymmetry, site-collection-admin coupling on the
Owners group), and is mounted in two places already (the group list, and Folder Access). A third
copy is how the three come to disagree about what "removed" means. This reuses the same primitives
(`removeGroupMember`, `setSiteAdmin`, `countSiteAdmins`) rather than re-deriving the logic, and keeps
the SAME safety rules — it is a new call site, not a new rule.

**Not a way to remove from the site entirely.** Exactly like `GroupMembersEditor`: removing from one
group is removing from ONE group. `CRS_SITE_MEMBERS` (site entry) and every other group the person
holds are untouched. Stated in the toast and the audit row, as `GroupMembersEditor` already does.

## The one wrinkle Quick Search has that `GroupMembersEditor` does not

`GroupMembersEditor` is scoped to ONE group, so its member list comes from ONE
`sitegroups(id)/users` read and each member row already carries its own site-user id
(`SpGroupMember.id`).

Quick Search is scoped to ONE PERSON across every group on the site, and its lookup
(`runLookup` in `GroupManager.tsx`) walks `GetUserById(u.Id)/groups` **per site user**, because a
guest can exist twice on this site for one address (2026-08-18) and the two accounts can hold
different group memberships. `UserGroupRef` (`{id, title}`) is deliberately group-shaped, not
person-shaped, and carries no site-user id today.

So removal needs a map the lookup does not currently keep: **which site-user id(s) actually hold
this group.** Built alongside the existing walk, in the component only — `UserGroupRef` and
`summarizeUserAccess` in `shared/userAccess.ts` stay untouched, because that module is pure and
SPFx-free by design and has no reason to know about site-user ids.

```ts
// keyed on GROUP id -> every site-user id (of this person's accounts) that holds it.
const membershipUserIds: Record<number, number[]> = {};
```

Built during the existing loop in `runLookup`, populated as each group is fetched — a group already
seen from an earlier `u.Id` still records the new `u.Id` too, so a genuinely duplicated membership
(the SAME group held by both of a person's two accounts) is fully addressable, not silently
half-fixed.

## Removal itself

One new function, `removeFromLookupGroup(groupId, groupTitle)`, called from the button:

1. Look up `membershipUserIds[groupId]`. Empty or missing -> refuse with a message naming the cause
   (the mapping was never built, most likely because the group list is stale — re-run the search)
   rather than doing nothing silently.
2. **If `groupId === owners?.id`** (the site's Owners group — the SAME `owners` state Quick Search
   already reads for the administrators card): apply the SAME site-collection-administrator guard
   `GroupMembersEditor.onRemove` applies —
   - never demote the SIGNED-IN admin themselves (checked by login name, case-insensitive);
   - never demote the LAST site collection administrator (`countSiteAdmins`, refuse on an unreadable
     count exactly as `onRemove` does — guessing "there must be others" is how a site is lost);
   - otherwise `setSiteAdmin(..., false)`, with a failure reported (not swallowed) because it would
     leave someone still holding full control while the toast says they were removed.
3. Call `removeGroupMember(sp, siteUrl, groupId, userId)` for **every** id in
   `membershipUserIds[groupId]` — covers the duplicate-account case without a second code path.
4. Toast: `"<name> removed from <group>"` plus the SCA caveat when it fired, same wording pattern as
   `GroupMembersEditor`.
5. Audit row: `EVENT.membersChanged`, source `"GroupManagement"` (matching every other Quick-Search
   and group-list write on this page), naming the group, the person, and — like
   `GroupMembersEditor`'s own log line — stating plainly that this removes ONE group and nothing else.
6. **Re-run `runLookup(lookupPerson)`** rather than hand-editing local state. The lookup is already
   the single source of truth for "what groups is this person in right now"; patching `lookupGroups`
   in place would be a second, competing definition of that fact, and one that cannot see a
   membership added or removed from another tab in the meantime.

## UI

- Each group row in Quick Search's result list (`lkRow`) gets a **Remove** button, right-aligned,
  matching `GroupMembersEditor`'s exact confirm pattern: click **Remove** -> the row shows
  **Confirm remove** / **Cancel** in place of it. No modal — this is a single, reversible-by-hand
  membership change, the same weight `GroupMembersEditor` already gives it, and a dialog for it there
  would be a double standard for the identical action.
- Disabled while its own removal is in flight (`lookupRemoveBusy === groupId`), so a double-click
  cannot fire the request twice; the REST call is not naturally idempotent against a fast double
  click (two `removeGroupMember` calls against an already-removed member both simply no-op on
  SharePoint's side in practice, but the SCA guard's `countSiteAdmins` read is not free, so this is a
  courtesy against wasted requests rather than a correctness requirement).
- Rendered for every group, Owners included — `describeGroupAccess` already tells the admin what a
  row is (including the special-cased Owners/site-entry copy), and hiding the button on those two
  would silently reintroduce the two workflows this feature exists to remove.

## Non-goals

- No bulk removal from Quick Search — that already exists on the group list itself
  (`Select all shown` / `Delete N selected`), and this feature is about a single person, not a bulk
  operation.
- No change to `GroupMembersEditor`, `removeGroupMember`, `setSiteAdmin`, or `countSiteAdmins` — all
  reused as-is.

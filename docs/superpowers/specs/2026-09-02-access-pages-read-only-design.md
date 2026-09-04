# Approval Library Access, Site Access, Page Access become read-only

**Date:** 2026-09-02
**Status:** AGREED, implementing now.

## Why

Client, 2026-09-02: *"the group assignment is all done by Folder recon, and how a group is have
access to anything is also done by folder recon, so I am see there is no point for all this pages to
allow client to manually strip or assign group to a page or library."*

Bulk provisioning creates groups and writes their Group Map rows; reconciliation grants the folder,
library and page ACLs from those rows. The write actions on these three pages either duplicate what
reconciliation already asserts, or manipulate a live ACL directly — a route that can drift from the
Group Map and get silently fought or ignored on the next reconciliation run.

## What was checked before agreeing

A survey of every write action on `StagingAccess.tsx` (Approval Library Access), `SiteAccess.tsx`
(Site Access) and `PageAccess.tsx` (Page Access) found some actions have no reconciliation
equivalent — deleting a Group Map row, revoking a live library grant, "Reopen to everyone" on a
locked page, Site Access's add/remove person, and the per-person `accessMemberUi` removal shared by
the first and third pages.

Each was individually confirmed as safe to drop:
- **Revoking a library grant** — the client already uses "delete the group, then recreate it" via
  Group Management as the standing repair pattern for narrowing access (used repeatedly this project
  for persona changes). SharePoint drops a deleted principal's role assignments, so this covers it.
- **"Reopen to everyone" on Page Access** — client explicitly does NOT want a locked admin page to
  have an un-lock button at all: *"when client locks a page, they won't want the page to be reopen
  for normal clients to see, that beats the whole purpose."* Removed outright, no replacement.
- **Site Access's add/remove person** — `CRS_SITE_MEMBERS` is a normal SharePoint group and already
  appears in Group Management's own group list; its member editor (`GroupMembersEditor`) already
  adds/removes people from it. No gap.
- **Per-person removal (`accessMemberUi`)** — same underlying `removeGroupMember` call exists on
  Group Management. The one thing `accessMemberUi` adds beyond that — checking whether a person is
  ALSO in a second group reaching the same library/page (the `removalVerdict`/"also in X" check) — is
  **not being replaced with new UI**. The client's own reasoning covers it: removing someone from ANY
  group ends whatever that group granted them, full stop — reconciliation asserts ACLs from Group Map
  rows and never leaves a stale grant behind for a group someone is no longer in. Nothing further
  needed.

## The change

1. **`StagingAccess.tsx`, `SiteAccess.tsx`, `PageAccess.tsx`** — remove every write action:
   - Approval Library Access: `onAdd`/`onAllowSelected`, `onRemoveMany`/single Remove, the
     `accessMemberUi` mount and its remove-member dialog.
   - Site Access: `onSetUp`, `onFixMissing`, `onAddPerson`, `onRemovePerson`.
   - Page Access: `runAllow`/`onAllow`, `onRemoveMany`, `onRestrictAdminOnly`,
     `onResetInheritance`, the `accessMemberUi` mount.
   - Each page keeps its READ-ONLY display (live ACL vs Group Map comparison, drift/"unexpected"
     banners, member counts, policy explanations) and gains a short banner: *"This page is read-only.
     To add or remove access, use Group Management."* — resolved link via the same `resolveLink`/
     Site Pages lookup the CRS Settings landing page and the retired Folder Access signpost already
     use, never hardcoded (this client renames pages at import).
2. **`GroupManager.tsx`** — add three link buttons beside "Select all N shown" / "Export what is
   shown": **Approval Library Access →**, **Site Access →**, **Page Access →**, resolved the same way,
   so an admin looking at the group list can jump straight to the detailed library/page/segment
   breakdown for any group.

## What is NOT changing

- Group Management's own group creation, deletion, and member-add/remove stay exactly as they are —
  they are the one remaining route for every action these three pages used to offer.
- The read-only comparison logic, drift detection, and persona/role display on all three pages is
  untouched — only the write controls come off.
- No schema change, no reconciliation change, no migration. Every Group Map row and every live ACL
  is unaffected; this only removes UI.

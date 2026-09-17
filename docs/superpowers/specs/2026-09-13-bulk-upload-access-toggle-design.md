# A toggle to restrict Bulk Upload to System Admins (QA item #23)

2026-09-13. Client: *"can you put a toggle at CRS Settings? it is to ensure nobody can access to
bulk upload, except for system admin."* Default state confirmed with the client: **open by
default** — nothing changes for anyone on deploy, an admin must deliberately switch it ON to
restrict the page.

## The mechanism this hooks into

`Bulk-Upload.aspx`'s access is entirely derived by reconciliation from
`shared/pageAccessPolicy.ts`'s static `RULES` table — the `bulk` rule currently reads
`{ roles: ["UPL", "UPLHC", "APR", "APRHC"], adminOnly: false }`, set that way on 2026-08-22 when
Bulk Upload was deliberately opened to every uploader. That table is read by THREE passes inside
`FolderManager.tsx`'s reconciliation run: the derived-page-access grant loop, the "refuse if
adminOnly" guard beside it, and the separate admin-page LOCKDOWN pass (which breaks a page's
inheritance and restricts it to Owners whenever `policyForPage(file).adminOnly` is true).

System Admin (CRS Owners) needs no special-casing anywhere in this. Owners hold Full Control on
the web regardless of any page-scope ACL, so a `Bulk-Upload.aspx` locked to `adminOnly: true` is
still open to every System Admin automatically - the same reasoning the admin-page lockdown
already relies on for every other locked page.

## The change

A NEW `CRS Config` setting row, `bulkUploadRestricted` (`yes`/`no`, same shape as the existing
`uploadsPaused` row), read once per reconciliation run and threaded into the page-policy functions
as an optional override - never baked into the static `RULES` table, which stays pure and
config-free.

- `shared/bulkUploadAccess.ts` - mirrors `shared/uploadPause.ts` exactly: a parser
  (`bulkUploadIsRestricted`) reading the same yes/true/on/1 spellings, a writer
  (`bulkAccessSettingValue`), and the config row's name as a constant. FAILS OPEN (missing, blank
  or unreadable -> not restricted) - the confirmed default, and also the safer failure direction
  here: a wrong `true` would silently take Bulk Upload away from every uploader on the site; a
  wrong `false` only means the page stays as open as it is today.
- `pageAccessPolicy.ts` - `policyForPage`, `derivedRolesForPage` and `isRoleEligibleForPage` each
  gain an optional `overrides?: { bulkUploadRestricted?: boolean }` parameter. When
  `bulkUploadRestricted === true` and the page name matches `/bulk/i`, the function returns the
  SAME shape an adminOnly rule already returns (`roles: []`, `adminOnly: true`) instead of
  consulting the static table. Omitting the parameter - every existing caller, and every existing
  test - behaves exactly as it does today.
- `FolderManager.tsx` - reads the `bulkUploadRestricted` row once near the top of
  `runReconciliation` (same pattern as the existing config-settings read at that point), and
  passes the resulting `{ bulkUploadRestricted }` object into every call site of
  `policyForPage`/`derivedRolesForPage` inside the page-access grant loop and the lockdown pass.
- A new component, `BulkUploadAccessToggle.tsx`, mounted from `CrsConfigurationWebPart.ts`
  alongside `FileTypeSettings`, on the SAME page reached via CRS Settings' "File Type Management"
  card. Mirrors `UploadPauseToggle.tsx`'s exact read/write/status-switch shape (SharePoint status
  switch, a "Refresh" button, an error state that never guesses when the read fails) - but unlike
  that component, this one stands alone rather than inside a guided-flow step, so it carries its
  own heading and one-line explanation.
- `shared/adminPages.ts` - the "File Type Management" card's blurb is widened to mention Bulk
  Upload access, since the underlying page (`CrsConfigurationWebPart` -> "CRS Configuration") now
  does two things rather than one. The card's `title`/`self.match` (which resolves the real page
  in Site Pages) are untouched - renaming either risks the dead-link failure this directory page
  is built to avoid.

## What does NOT change

- The static `RULES` table's `bulk` entry - the override is applied on top of it, never inside it.
- `Bulk-Upload.aspx`'s existing behaviour for anyone when the toggle is untouched (default OFF /
  not restricted).
- The folder-ACL / `AddListItems` probe Bulk Upload already does at write time - this toggle only
  ever controls whether the PAGE is reachable at all, exactly like every other page lock in this
  project; it grants nobody a new place to write.
- `FileTypeSettings.tsx` itself - the new toggle is a sibling component, not a change to that
  file's own logic.

## Non-goals

- No per-segment or per-persona granularity - the client's ask is a single site-wide switch,
  matching `uploadsPaused`'s own scope.
- No change to how System Admin/Owners access is determined anywhere else in the codebase.

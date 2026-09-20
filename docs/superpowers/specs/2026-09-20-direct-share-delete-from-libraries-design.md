# Direct share/delete from the four approved-side libraries — design

**Client's ask, 2026-09-20:** *"Can we reuse this page from My submission when open a file in the
document library and the archive library? When client click on the file it will open this page?
Same behaviour like ApprovalDocument. Reason being Approver cannot share since the native sharing
is disabled so the best way is to replicate the same behaviour in approved libraries to open the
page and let them share a file. Follow the same logic as well, Approver can share without
permission and same goes for system admin."* Scope confirmed in conversation: all four
approved-side libraries — `Documents`, `HC Documents`, `Archive`, `HC Archive`.

## The problem

Native SharePoint sharing was disabled site-wide on 2026-09-07 (*"Only site owners can share
files..."*, per CLAUDE.md's "MEMBERS CANNOT SHARE" entry). That closed the hole it was built to
close — a plain member can no longer grant access to anything — but it also took away an Approver's
only route to share an already-approved document straight from the library, since Approvers are
not site owners.

My Submissions and the Requests page both already solve this for their own contexts: a signed-in
user can open one of their own submissions, or an approver can open a document behind a request,
and both land on the same `FileDetailPanel` view with a direct Share (and sometimes Delete) button
when `probeFileRights` says they hold the right. There is no equivalent entry point when someone is
simply browsing `Documents`/`HC Documents`/`Archive`/`HC Archive` directly and clicks a file.

## The approach: a new mode on the Requests page, not a new web part

`Requests.tsx` already has a `?request=<Id>` mode that replaces its whole page with a single
`FileDetailPanel` plus an action card — added 2026-09-10 so a request-notification email can link
straight to one file. This adds a **sibling mode**, `?file=<UniqueId>`, that skips the request
entirely: given just a file's identity, resolve it, show the same panel, and offer direct actions
if the viewer's rights allow it.

**Why not a new web part:** this project has been burned twice by exactly that shape of change —
the `+ New Folder` customizer and `CRS Requests` itself both sat silently undeployable for weeks
after being bundled but never registered as a page component. Reusing the already-registered
`Requests` web part and its already-published page avoids that class of risk entirely. It also
avoids creating and publishing a fourth page for the client to keep track of.

## Routing: `?file=<UniqueId>` only — no library parameter

`GetFileById(guid'<UniqueId>')` is web-scoped, and a file's `UniqueId` is unique site-wide, not
per-library. It resolves across all four target libraries (and in fact any library) with no
ambiguity, so there is nothing for a `lib=` parameter to disambiguate. This mirrors the convention
My Submissions already uses (`?file=<UniqueId>&sfi=<SubmissionFileId>`), simplified further here
because there is no submission record to also carry.

⚠ **One open risk, to verify on first deploy, with a documented fallback.** Column formatting JSON
needs a token exposing the clicked row's `UniqueId` — this design assumes `[$UniqueId]` is
available the same way `[$ID]` already is on the approval libraries' existing Name-column
formatting. If SharePoint does not expose it in that context on this tenant, the fallback is to add
`&lib=<key>` back (using the existing `LibTarget` keys — `Documents`, `DocumentsHC`, `Archive`,
`ArchiveHC` — already defined in `shared/naming.ts`) and resolve via `[$ID]` plus the known library
title instead. Confirm which is needed before finalising the column-formatting JSON for the client.

## Resolution: a sibling loader, not a reuse of the request-based one

`Requests.tsx`'s existing `loadFileView` resolves a file behind a *request* — it tries
`row.itemUniqueId` first, then falls back to `findByStamp` because a request's target may have been
routed (moved from the approval library to `Documents`, changing its `UniqueId`) since the request
was raised. That fallback does not apply here: a file already sitting in one of the four target
libraries is not going to move again on its own, so a plain `GetFileById` resolution is enough. The
new loader mirrors the existing one's shape (web-scoped `GetFileById` → `ServerRelativeUrl` →
`GetFileByServerRelativeUrl(@f)` alias form for `Name`/`Length`/`TimeLastModified` →
`FieldValuesAsText` → the raw `DocumentDate` re-read and reformat) without the stamp-fallback
branch, and without a request row to attach to.

Everything else the panel needs is derived from the resolved path, not the URL:

- **Location / folder trail** — `folderTrail`/`trailText` against `libraryTargets()`'s URL
  segments, exactly as the existing mode already does.
- **HC tag** — `isHcRow(librarySegmentOf(fileRef, segments), hcSegs)`. All three functions are
  already imported into `Requests.tsx`.
- **Archive classification** — `isArchivedRow(...)`, same imports. This is what gates the rule
  below.

No new imports are needed for any of this; `Requests.tsx` already has `librarySegmentOf`, `isHcRow`,
`isArchivedRow`, `cachedHcLibraries`, and `cachedArchiveLibraries` in its import list from earlier
work. The one new import is `probeFileRights`, from `shared/dmsFolderMap.ts`.

## Rights gating: three tiers, and Archive overrides all of them

`probeFileRights(spHttpClient, siteUrl, uniqueId)` — the same `EffectiveBasePermissions` check
built 2026-09-15 for My Submissions — answers whether *this viewer* can remove or share *this
document*, accounting for the folder ACL, inheritance, and site-admin status in one read. It is
already library-agnostic; nothing new is needed to make it work here.

**Archive is unconditionally read-only, regardless of what the probe answers.**
`probeFileRights`'s own documentation is explicit that it answers permission only, not
appropriateness — *"an archived document is read-only to everybody by design, and the caller must
enforce that"*. So for any file resolving into `Archive` or `ArchiveHC`, Delete and Share are both
suppressed outright, even for a system administrator whose Full Control would otherwise make the
probe say "granted."

For everything else, the outcome has three tiers, confirmed in conversation:

1. **A plain uploader/PIC with no direct rights** — read-only file view. No Delete, no Share. (This
   was the explicit choice over offering a "Request…" fallback here — see Non-goals.)
2. **An ordinary Approver** — **Share only.** Since the 2026-09-17 proxy-delete persona change
   removed `DEL`/`DELHC` from every persona including Approver (so that nobody but the proxy
   account can delete a *folder* through the native UI), an Approver's `probeFileRights` check for
   `remove` will come back denied on an already-approved document — same as it already does on My
   Submissions today. This is a consequence of existing, already-shipped behaviour, not a new
   restriction introduced by this feature.
3. **System Admin** — both Delete and Share, via Full Control.

## Actions

**Delete** follows the same proxy mechanism as every other delete in the app since 2026-09-17:
write a `CRS Requests` row with `Status` set to `"Approved"` from the moment it is created (nobody
decides it), and the existing `CRS — Execute approved deletion` flow performs the actual recycle,
guarded by the ETag check from `2026-09-20-etag-guarded-proxy-deletion-design.md`. This will be a
small sibling function to `MySubmissions.tsx`'s `writeApprovedDeletionRequest`, not a call into it
directly — that function is typed around My Submissions' own `MergedRow`, and this project's
existing practice is already one local copy per page (`Requests.tsx` and `MySubmissions.tsx` each
already have their own `performShare`/delete-writer rather than a shared one), so a third small
copy of the same shape is consistent with that, not a new pattern.

**Share** calls `SP.Web.ShareObject` directly and immediately — no request row at all — matching My
Submissions' existing "direct share" path: View-only permission only (no Edit option, per the
client's 2026-08-27 decision to remove that choice generally), no expiry field (nothing enforces an
expiry on the direct path today). The recipient entry reuses the already-shared
`searchTenantPeople` lookup but gets its own small, purpose-built chip-entry UI rather than reusing
My Submissions' full request dialog, which carries a reason field and a request/direct branch that
do not apply here — this page never raises a request.

Both actions get a plain Yes/No confirm dialog, no typed confirmation — matching the existing
"Delete this file? / Yes / Cancel" pattern already used for direct delete on My Submissions.
Nothing here is more destructive than what that dialog already guards.

## UI details

- **Read-only for everyone without direct rights** (confirmed): the panel and its details, no
  action buttons, no explanation banner needed beyond what `FileDetailPanel` already shows.
- **"Back" is `history.back()`**, not "Back to requests." The existing `?request=<Id>` mode's back
  link makes sense because a click there always originates from the Requests page's own list; a
  click reaching this new mode originates from a *different* page (the library view), so pointing
  it at the Requests queue would be a wrong destination. This mirrors the same fallback this
  project's own `BackBand`/`backToSettings` pattern already uses when it cannot resolve a specific
  page to return to.
- **A small, testable extraction**: `directActionsFor(rights: FileRights, isArchived: boolean):
  { canDelete: boolean; canShare: boolean }` as a new pure function, so the three-tier + archive
  override rule has one definition and a unit test, rather than being reasoned about only inline in
  JSX.

## Error handling

Mirrors the existing `loadFileView`'s three states (`loading` / `gone` / `ready`) exactly — a
missing or inaccessible file says so rather than rendering a blank panel, and a failed
`FieldValuesAsText` read degrades to "no details were recorded," never blocks the preview.

## Column formatting — delivered once the code is built and deployed

The client will need the Name-column formatting JSON for all four libraries, pointing at
`<published Requests page URL>?file=[$UniqueId]` (or the `[$ID]`+`lib=` fallback if `[$UniqueId]`
proves unavailable — see the routing section above). This is a manual, per-library, per-site step,
same as the existing HC library Name-column formatting already is — it is not shippable via the
SPFx package. Delivered as a follow-up once the page is confirmed working, since the exact JSON
depends on confirming the token first and on knowing this site's actual published page URL (this
client renames every page at import).

## ⚠ The page's role policy must widen, or a plain uploader gets AccessDenied

Found during spec self-review, not assumed: `Requests.aspx`'s current page policy
(`shared/pageAccessPolicy.ts`, `/request/i`) is `roles: ["APR", "APRHC", "DEPTVIEW"]`. `UPL` was
deliberately *removed* from it on 2026-08-21, precisely because a plain uploader's queue there was
always empty — their own requests moved to My Submissions instead, so the page had no purpose for
them.

This feature changes that: a PIC browsing `Documents` and clicking a file now has a genuine reason
to reach this page — the read-only `?file=` view. Left as-is, they would get AccessDenied on the
whole page instead. So this design **adds `UPL` and `UPLHC` back to that page's role list** —
the fifth instance of the pattern this file's own comments already track (*"a page keyed on a role
its audience does not literally hold"* — upload form → `APR`, Requests → `DEPTVIEW`, My Submissions
→ `UPLHC`, and now this).

This is safe by the same reasoning already written into that rule: *"widening this list cannot
grant anyone a new place to act — the page grant opens the screen, while `canDecide`/`isVisibleTo`
decide each row, and the approval itself runs in the viewer's own session."* A plain uploader
reaching the page's default (queue) view still sees nothing to act on — `canDecide`/`isVisibleTo`
are unchanged by this. Migration is a page-policy code change plus a reconciliation re-run, same as
every other page-role change in this project's history.

## Non-goals

- **No "Request…" fallback on this page for viewers without direct rights.** That flow already
  exists on My Submissions (for a viewer's own files) and the Requests queue (for an approver
  deciding on someone else's). Duplicating it here was considered and explicitly declined in favour
  of a smaller build.
- **No extraction of `performShare`/`writeApprovedDeletionRequest` into a shared module.** Each
  page already has its own copy; adding a third small copy of the same shape is consistent with
  existing practice and avoids an unrelated refactor of two already-working functions.

## Testing

- New pure function `directActionsFor` (rights tier + archive override) gets a unit test — this is
  the one piece of genuinely new decision logic in the feature.
- The resolution loader, the delete/share writers, and the UI itself are not unit-testable in this
  project's existing style (no UI tests anywhere in the codebase) — verified by `tsc`/lint/the full
  suite passing with no regressions, and by a live click-through once deployed, same as every other
  feature in this project's history.

# Archive — audit log tracking + access narrowed to C-Level only

**Date:** 2026-09-02
**Client's instruction, verbatim:** *"we need to make sure Archive is also track in audit log and
client update, only C level and system administrator have access to archive, so no more HOD till
Normal viewer have access to Archive Library, and that would mean the My Submission should not
update to archive either."*

Three changes. Two are code, shipped in `1.0.363.0`. One is a Power Automate change on the two
archive mover flows, not yet built — this file is the runbook for it.

---

## 1. DONE (code, `1.0.363.0`) — Archive/ArchiveHC narrowed to C-Level only

`LIBRARY_ROLES.Archive` / `LIBRARY_ROLES.ArchiveHC` in `FolderManager.tsx` were
`["MEMBER", "DEPTVIEW", "DEL", "GLOBAL", "SEGVIEW", "UPL", "APR", "SHARE", "UPLHC", "APRHC", ...]`
(every persona, mirroring the working-side libraries, per the 2026-08-22 design). Narrowed to
`["GLOBAL", "SEGVIEW"]` — the two C-Level roles (`clevel_global`, `clevel_segment`) — only.

**"System administrator" needed no row.** `CRS Owners` holds Full Control on the web regardless of
any folder ACL, the same reason the admin-page lockdown never has to name it explicitly.

**⚠ THIS DOES NOT REVOKE ANYTHING ALREADY GRANTED.** Reconciliation only ever ADDS folder-scope
grants — `groupsToRemove`'s full-ACL assertion is page-scope only, because a folder root also
carries SharePoint's automatic Limited Access entries for everyone granted deeper in the tree, and
asserting there would strip those too. So on ClarenceDMSTesting, which has already run
reconciliation against the WIDER (pre-2026-09-02) rule — including today's 57 + 17 test archives —
every HOD/PIC/HOU/MEMBER group that was already granted Read on `Archive`/`Archive Highly
Confidential Document` folders **keeps that Read** until it is removed by hand. Narrowing
`LIBRARY_ROLES` only changes what a **future** reconciliation run will grant.

- **On SDG's tenant this is a non-issue**: nothing has been archived there (nothing is 7 years old
  yet, and the mover's real rolling cutoff was restored today) — the narrower rule takes full effect
  the very first time anything is archived, with no legacy grants to clean up.
- **✅ DONE, RUN LIVE ON BOTH SITES, AND REMOVED — a one-time repair button, "Prune archive access
  to C-Level only", on the Reconciliation tab of Folder Administration (shipped `1.0.364.0`, removed
  `1.0.365.0`).** Walked every `Archive`/`Archive Highly Confidential Document` folder down to
  segment/department/unit depth, removed any group not currently mapped `GLOBAL` or `SEGVIEW` for
  that folder's own segment, reported every removal by folder + group name. Site Owners and Limited
  Access (navigation) entries were never touched, confirmed in both run logs — `CRS Owners` appears
  nowhere in either removal list.
  - **ClarenceDMSTesting**: 97 grants removed, all on `Archive Highly Confidential Document/MHO`'s
    segment root — the rest of that subtree inherits from it rather than breaking its own
    inheritance per department/unit.
  - **SDG (production tenant)**: run across GHO, MHO and NBPOLHO in both libraries. Hundreds of
    stale `<SEG>_..._UPLOADER`/`APPROVER`/`VIEWER`/`HOD` grants removed.
  - **⚠ ONE RECURRING NON-GROUP ENTRY ON SDG, "Guthrie Central Repository System" (the site's own
    title)**, removed once per segment root on both libraries — almost certainly a SharePoint
    auto-created default group from initial site provisioning that was never renamed, holding one
    blanket ancestor-Read grant per segment. Plausible, not independently confirmed. If archive
    access looks wrong on SDG later, check that principal first.
  - **Removed from the codebase the same day** the client asked (*"we don't want to confuse the
    client with a new feature like this"*) — the button, `runArchivePrune`, and its state are gone;
    `removeRoleAssignment` is back to "unreachable on purpose" (kept, not deleted, per this
    project's usual habit for capabilities that may be needed again). This file is the record of
    the exact shape if the same class of repair is ever needed a second time.
  - **⚠ STILL OUTSTANDING ON BOTH SITES: re-run Folder Reconciliation.** The prune only removes —
    confirming the surviving GLOBAL/SEGVIEW (and Owners) grants are still correctly in place needs
    the ordinary reconciliation run, same as after any other folder-ACL change.

## 2. DONE (code, `1.0.363.0`) — My Submissions no longer reads Archive

`MySubmissions.tsx`'s `load()` no longer fetches rows from `Archive`/`Archive Highly Confidential
Document`. My Submissions is a PIC/HOU page (`Upload-Form.aspx`'s page policy is mirrored here);
neither role holds `GLOBAL` or `SEGVIEW`, so the read would now just 403/404 for every viewer of
this page — removing it stops the wasted request and matches the client's own framing that a normal
uploader should see nothing of the archive from here.

The archive library's title is still read from cache (`cachedArchiveLibraries()`, no network cost)
because `isHcRecord` still needs it to correctly classify a `CRS Submissions` record stamped against
the archive HC library, independent of whether this page ever fetches rows from it.

## 3. TODO (Power Automate, both archive flows) — write an `Archived` audit row per file

**`EventType: "Archived"` is registered in code** (`auditLog.ts` — `EVENT.archived`, `EVENT_LABEL`,
`ALL_EVENT_TYPES`), so it is filterable in the Audit Log viewer's Action dropdown the moment a row
with that value exists. Nothing else needs to change in code — the column (`EventType`) is plain
Text, so it accepts the new value from a flow with no schema change.

### Where to add it

In **both** `CRS — Archive after seven years` and `CRS — HC Archive after seven years`, inside the
`For each` loop, add a **`Create item`** action (the native SharePoint connector action, not a raw
HTTP call — it handles JSON-light encoding itself, unlike the raw HTTP calls this flow already makes
for the move/stamp/reset).

**Placement: AFTER `Reset_inheritance` succeeds**, i.e. as the last action in the loop for that file.
That is the point at which the move, the `Archived` stamp and the permission reset have all
completed — logging before that would record files that then failed to move cleanly.

`Run after`: **is successful** only (leave the other two run-after boxes unticked). Unlike
Auto-route's audit writes — which must never block a document routing and so run after both success
AND failure — an archive audit row that never gets written because the move itself failed is the
correct outcome: nothing happened to log.

### Fields on the `Create item` action

Site Address / List Name: the `CRS Audit Log` list (same site).

| Column | Value |
|---|---|
| `Title` | `Archived: ` + the file name (dynamic content from the loop item), matching the summary shape every other event type uses |
| `EventTime` | `utcNow()` — ISO, matching every other flow-written row (`EventTime` is written and filtered as ISO throughout this list; the `M/D/YYYY` locale rule is a DIFFERENT endpoint's gotcha and does not apply here) |
| `EventType` | the literal text `Archived` |
| `Outcome` | the literal text `Success` |
| `ActorName` | a fixed literal — e.g. `Archive mover (automated)`. There is no human decision per file the way there is for an approval; unlike `Routed`'s `ApprovedBy`, this event has no approver to attribute to. |
| `ActorEmail` | the SERVICE ACCOUNT's own email, typed as a literal. The flow already runs as that account (§0 of the base runbook: sign in as the service account before the first action) — this just states it on the row rather than leaving it blank. |
| `Source` | the literal text `Flow:ArchiveMover` (normal) / `Flow:ArchiveMoverHC` (HC clone) — matching the `Flow:<Name>` convention every other flow-written audit row already uses |
| `LibraryName` | the literal text `Archive` (normal) / `Archive Highly Confidential Document` (HC) — or, better, the live library title if it is already available as dynamic content from an earlier action in the loop, so a future rename does not leave this stale |
| `ItemUniqueId` | the file's UniqueId — see below, this needs one more action if the loop does not already capture it |
| `ItemName` | the file name |
| `ItemPath` | the destination path (`DestPath`'s output, already computed earlier in the loop for the move itself) |
| `Segment` / `UnitPath` | leave blank unless already trivially available from an earlier Compose in the loop — not worth adding a new lookup just for these two optional columns |
| `Details` | optional; a short note such as the source library name, if useful |

### Getting `ItemUniqueId`

**Confirmed today (this session): a cross-library `MoveTo` preserves `UniqueId` exactly** — verified
live by reading the same file's UniqueId before and after the move. So the value logged here can be
read either from the SOURCE item (before the move, if the loop already has it from `Get items` on
`Documents`/`HC Documents`) or from the DESTINATION item (after the move) — they are the same GUID
either way, and using it is what lets a future "History of this file" style feature thread an
archived document's whole lifecycle (Uploaded → Approved → Routed → Archived) on one id, the same
way the existing file-lifecycle audit flows already do for the earlier stages.

If the loop does not already expose it as dynamic content, add one more `Send an HTTP request to
SharePoint` — `GET`, `_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields/UniqueId?@f='...'`
against the DESTINATION path — placed after the move and before this `Create item`, mirroring the
pattern the base runbook already uses for the stamp/reset calls.

### The HC clone — same six-reference discipline as every other HC clone in this project

Build the HC flow's `Create item` as a Save-As-derived clone, and change:
- `Source` → `Flow:ArchiveMoverHC`
- `LibraryName` → the HC archive library's title
- whichever action supplies `ItemUniqueId` → pointed at the HC archive library, not the normal one

This project's own history (six faults in the original HC Auto-route clone, one missed reference in
this session's own HC archive clone build) makes this worth stating plainly: check every field this
action reads dynamic content from, not just the ones that look library-specific at a glance.

### Testing

Run either flow manually against a small batch (or wait for the next real archive), then confirm in
the Audit Log viewer: a row reading `Archived`, the correct file name and path, `Source:
Flow:ArchiveMover` (or the HC variant), and — the one that actually matters — that `ItemUniqueId`
matches the same file's earlier `Uploaded`/`Approved`/`Routed` rows, so "History of this file" (if
that feature is ever pointed at this list) can follow the document all the way to the archive.

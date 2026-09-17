# QA deck follow-up (rows 17, 22, 23, 26, 32-34, 46-48, 50, 53, 54, 57)

2026-09-13. Client pasted a QA tracker deck; #49 explicitly excluded (separate feature, tracked
elsewhere). This spec is the record of what each row actually meant once checked against current
source, what was built, and what needs the client's own answer before anything is guessed at.

**Rule followed throughout, per this project's own standing lesson: check current source before
building a fix.** Several rows describe a state the CURRENT code does not have — either because an
earlier session already fixed it, or because the mechanism the row assumes does not exist. Guessing
a fix for a symptom that is not reproducible in source risks introducing a real regression to "fix"
something that already works.

## Already fixed in current source (no code change made)

- **#22 - Search shows "one library not searched, staging http 404".** `DocumentSearch.tsx`'s job
  list no longer includes `Staging`/`StagingHC` at all (removed 1.0.526.0, 2026-09-10 — client:
  *"it is not suppose to search anything from staging library only document library"*). The row is
  dated the same day as that fix shipped; the deployed package the QA account was testing was almost
  certainly one build behind. Nothing to change; deploy the current build and re-test.
- **#32 - "Disable to pull remarks... under Remarks when user key in in homepage's search field."**
  `Remark` is not in `TEXT_COLUMNS` (the KQL free-text match list) nor in `buildListFilter` (the REST
  free-text match list) anywhere in current source — confirmed by direct read of
  `shared/documentSearch.ts`. Only `ProjectName`, `Vendor_x002f_CustomerName` and `Keyword` are
  matched. Whatever build the client tested this against, it is not the current one.
- **#50 - "Unlock the feature for client to rename the Term Abbreviation for Segments."** This is the
  segment top-folder re-code feature (`shared/segmentRecode.ts`), built and wired into
  `SegmentCreator.tsx`'s Segments tab on 2026-09-11 (`allowRecode !== false` gates it, false only on
  the Retire screen). Already reachable; nothing to add.

## Built this session

### #33 - Approver's reject/approve reason
Confirmed live: `Requests.tsx`'s decision dialog currently reads "Note (optional)" on approve and
only demands a reason on reject — a comment at the site says this was itself a deliberate reversal
"(client, 2026-09-11)". The QA deck, dated the same day, reports the opposite as a bug. Given the
deck is the instruction in front of me now, reason is made mandatory on BOTH decisions again, on
both Delete and Share requests (one dialog serves both types). Flagged in the final report rather
than silently overridden, since there is direct evidence in the code of an explicit client reversal on
the same date — worth the client confirming which way they actually want it to land.

### #34 - My Submissions copy (3 items)
1. Summary sentence -> "Files that you have uploaded and its location will be shown here."
2. The footer note under the Permission (Requests) tab is removed.
3. Request references gain the literal prefix wording "Form Submission No.: " before the
   `SUB-YYYYMMDD-XXXX` reference, wherever a submission reference is shown on this page.

### #46 - File Type Management copy (FileTypeSettings.tsx)
The two remove-confirmation paragraphs currently read:
"It stops being offered on this page, and can no longer be uploaded." / "...it already could not
be uploaded, so what may be uploaded does not change." and "Files already uploaded are untouched -
this only affects new uploads. You can add {ext} again at any time."

Read as ONE combined instruction (item 1 names the sentence to remove, item 2 gives its
replacement): the first paragraph becomes a single line, "Remarks: Files already uploaded won't be
affected." - dropping the "can no longer be uploaded" framing entirely, which is what item 1 asked
to remove. The still-important "only allowed type left" warning block is untouched - it is a
different, separately-triggered warning. "at any time" -> "anytime" wherever it appears in this file
(the removal dialog and the post-removal toast), matching item 3.

### #47 - Page Access, no target text given
The deck names the string ("Go to Group Management.") but does not say what to change it to. Not
guessed at - flagged for the client to supply replacement copy. Present in three files
(`ApprovalLibraryAccessPage.tsx`, `FolderAccessPage.tsx`, `SiteAccessPage.tsx`), all resolved from
Site Pages rather than hardcoded, so whichever wording is chosen goes in one place per file.

### #48 - Audit Log layout (4 items)
1. Date From/To wrap onto two lines in practice, despite already being `display:flex` - the
   shared `s.input` style's `minWidth: 180` on both date inputs, inside a grid column of roughly
   250px (two columns per the client's own 2026-09-07 design), leaves no room for two side by side.
   New `dateInput` style drops the shared min-width to a size that actually fits two per row without
   spanning both grid columns (which would push Action back onto its own row, undoing an earlier
   explicit fix).
2. Calendar popup width matching the input is a NATIVE BROWSER WIDGET
   (`<input type="date">`'s picker) and cannot be sized from CSS in any current browser - flagged as
   a platform limitation, not attempted. The only way to get a stylable calendar is to replace the
   native date input with a custom JS date-picker component, which is a materially larger change than
   this row implies and is not undertaken without the client being told what it trades away (native
   keyboard entry, native accessibility, native mobile date pickers).
3. Export + its caption moved together: "export what is shown" moves from a separate row spanning
   the whole filter card into the `filterActions` button column, directly under the Export button,
   rather than sitting in a row that merely happens to align near it.
4. "Who" column standardises to a display name. `r.ActorName || r.ActorEmail || "-"` already
   prefers the name; rows still showing a raw address are ones where `ActorName` was never written
   (mostly flow-authored rows). A UI-only fallback derives a readable name from the email's local part
   (dots/underscores -> spaces, each word capitalised) rather than showing the bare address, without
   touching what is actually stored - the CSV export and the underlying data are unaffected.

### #53 - "Business Segment" -> "Segment"
Twelve files reference the literal string; most are comments. The DEFAULT display label (used
whenever a document's segment is not a Group-Led Project) changes from "Business Segment" to
"Segment" everywhere it renders to a user:
- `shared/documentDetails.ts` - `tierRows`'s fallback for the segment tier (the one function every
  other screen's default flows through: My Submissions' detail and batch views, the approval page's
  shared field list). The Group-Led Project override is untouched.
- `ApprovalDocument.tsx` - its own local `segmentLabel` resolver, which feeds the shared function
  above with an explicit value for THIS document; its two literal returns of "Business Segment"
  become "Segment".
- `Form.tsx` and `BulkUpload.tsx` - the "Upload to" family radio (Business Segment / Group-Led
  Project) and the saved-batch-card label that mirrors it.
- `SegmentCreator.tsx` - the "Appears under" radio, for the same reason: it is the identical
  two-option pairing used everywhere else, and leaving it as the one exception would be the
  inconsistency this row exists to remove.
- `DocumentSearch.tsx` - the Advanced Filters "Business segment" dropdown label.

Not touched: the SharePoint column names (`Business_x0020_Segment`, `BusinessSegmentTid`), the
`labelFromInternalName` generic decoder (still correctly decodes that internal name to "Business
Segment" as a raw fact about the column - its own unit test asserts this and stays valid), and the
`Category === "Project"` comparison logic anywhere (data values, not labels).

### #54b - Group Management Quick Search: person in zero groups
Confirmed the gap by reading `runLookup`: when a lookup finds a real site account with zero
group memberships, nothing renders under their name at all - no message of any kind, because the
existing "not a member" note is only set when the address matches NO site account, a different case.
A `lookupGroups.length === 0` (successful read, empty result) branch now renders "Not in any group
yet.", the client's own wording.

### #54a - Audit Log not showing Archived items
Not code. `EventType: "Archived"` is registered in `shared/auditLog.ts` (`EVENT.archived`,
`EVENT_LABEL`, `ALL_EVENT_TYPES`) and is filterable in the viewer the moment a row with that value
exists. Nothing currently WRITES that row - the two seven-year archive movers are Power Automate
flows, config that lives outside this repository, and the runbook for adding their `Create item`
action (`docs/superpowers/specs/2026-09-02-archive-audit-log-and-access-narrowing-runbook.md`) has
never been carried out on either flow. This needs a Power Automate change, not an SPFx one.

### #57 - Refresh beside "Open Term Store Management"
`StructureManager.tsx` already has this exact pairing (link + "Re-check" button) beside its below-Unit
level's Term set ID field, added 2026-09-09. `SegmentCreator.tsx` - the screen used when creating a
NEW segment, arguably the more commonly hit one - has neither: its Term set ID field is a bare input
with a status line and no way to open the term store or force a re-check without retyping the GUID.
Added the identical pairing there: a "recheck" counter added to the existing check effect's
dependency array, an `openInNewTab`-driven link to `{siteUrl}/_layouts/15/termstoremanager.aspx`
(the classic, site-level page - the modern one 404s for a site collection administrator, per this
project's own prior finding), and a "Re-check" button shown only once the verdict came from the
server (checking/found/notfound/unknown - never on blank/malformed, where re-asking cannot change
the answer).

## Flagged, not built - needs the client's own answer rather than a guess

- **#17 - Document Upload approval email restricted to a single service-account address.** This is
  Power Automate flow logic (the approver-notification flow's recipient filtering), not SPFx code.
  Nothing in this repository controls who a flow emails. Needs a runbook change to
  `CRS — Approval reminder` / the notify-approvers flow, on the live tenant, by whoever administers
  it - the same category of work as #54a.
- **#23 - Bulk Upload "Enable a function to display Bulk Upload".** Marked TENTATIVE by the client's
  own deck. No concrete, actionable change is stated (Bulk Upload already exists as a web part and a
  page). Not guessed at.
- **#26 - "Disable ALL SHARE Function for Viewers" / 404 when a Viewer requests sharing.** Checked
  the code path: `My-Submissions.aspx` (the only page with an in-app "Request sharing" button) is
  page-ACL-restricted to `UPL`/`UPLHC`/`APR`/`APRHC` roles - a pure Viewer (SDG Employee / `MEMBER`)
  should not be able to open this page at all, which makes the reported 404 hard to reproduce from
  source alone. The mechanism this project already built for exactly this ask - "Site permissions ->
  Sharing settings -> Only site owners can share files", which blocks the NATIVE SharePoint share
  button for everyone but owners - was verified working end to end on 2026-09-07. Recommend
  re-confirming that site setting is still in place before any code change; guessing at a code-level
  persona gate without a reproducible path risks adding a check that never fires, giving false
  confidence that the row is closed.

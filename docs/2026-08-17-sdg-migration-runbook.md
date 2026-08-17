# CRS — migration runbook for SD Guthrie's tenant

**Date:** 2026-08-17
**Package:** 1.0.128.0
**Target:** SDG tenant, one site collection. Site collection app catalog **already exists**.
**Flow identity:** an SDG-issued ("Guthrie") account, never the Trinergy one.

> Work top to bottom. The order is a **dependency order**, not a preference: several steps create
> nothing and report success when the step before them was skipped. Every step has a **Verify** —
> do it before moving on. A step that cannot be verified has not been done.

---

## 0. Before anything

| # | Check | How | If it fails |
|---|---|---|---|
| 0.1 | Your permission level | Site Settings → Site collection administrators | **Site Collection Administrator is required.** Site Owner alone cannot manage the in-site term store (§3) and may not create permission levels (§6). Establish this before starting, not at step 3. |
| 0.2 | App catalog | Site Contents shows **Apps for SharePoint** | Ask a SharePoint Administrator for one line: `Add-SPOSiteCollectionAppCatalog -Site <url>`. Nothing else can proceed. |
| 0.3 | Existing content | Site Contents | The built-in `Documents` library has the URL segment **`Shared Documents`**, not `Documents`. Never build a path from the title. |
| 0.4 | The flow account | Sign in to Power Automate as it once | If it is a personal SDG account rather than a shared service account, **say so in writing now.** A flow runs under the connection its first action creates, so that identity is permanent; when the account is deactivated both flows stop **silently** and present as "my colleague's folder doesn't exist". |

**Naming.** This client uses the **`CRS`** prefix. `naming.ts` probes `CRS <suffix>` before
`DMS <suffix>` per list, so a partial rename degrades gracefully — but provision everything as `CRS`
from the start and each probe costs one request instead of two.

---

## 1. Deploy the package

1. Upload `sharepoint/solution/sd-gatrie.sppkg` to the site collection app catalog.
2. Trust it when prompted.
3. Site Contents → **Add an app** → add the solution.

**`skipFeatureDeployment` is `false` and must stay false.** It is what activates the feature that
provisions `elements.xml`, which is the only thing registering the `+ New Folder` customizer. With it
`true` the package deploys, every web part works, and the customizer never runs — with nothing
anywhere connecting the two.

**Verify:** the app appears in Site Contents and its web parts are offered when editing a page.

**After any later re-upload you must also click Update in Site Contents.** Feature elements run on
install and update only. The bundle refreshes on its own, so web parts appear to update while the
customizer does not — which reads as "the fix didn't work" rather than "the app wasn't updated".

---

## 2. The `+ New Folder` customizer

Bundling a component only makes it *available*. An application customizer runs only where a
`UserCustomAction` points at it.

**Verify first:**

```
/_api/web/usercustomactions?$select=Title,Location,ClientSideComponentId
```

- **One entry** for `c1d2e3f4-a5b6-47c8-d9e0-f1a2b3c4d5e6` → done.
- **Zero** → the feature did not provision it. This happened on ClarenceDMSTesting even with
  `skipFeatureDeployment: false`, so treat it as likely rather than exceptional, and register by
  hand:

```js
await fetch(`${web}/_api/web/usercustomactions`, {
  method: 'POST',
  headers: {
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata',
    'X-RequestDigest': digest,
  },
  body: JSON.stringify({
    Title: 'CRSNewFolderButton',
    Location: 'ClientSideExtension.ApplicationCustomizer',
    ClientSideComponentId: 'c1d2e3f4-a5b6-47c8-d9e0-f1a2b3c4d5e6',
    ClientSideComponentProperties: '{}',
  }),
});
```

- **Two entries** → the component loads twice: two buttons, two MutationObservers. Delete one.

---

## 3. Term store (in-site)

Managed **in the site**, never the tenant admin centre — the client refuses tenant term-store access.
Needs Site Collection Administrator, or Contributor on the term group.

1. Site Settings → **Term store management**.
2. Create the site-collection term group.
3. Create one term set **per business segment**, plus `Document Type`, `Year/Period`,
   `Confidentiality`.
4. Author Department and Unit terms under each segment set.

**Rules to give the client in writing, because every one of these fails silently:**

- **Rename terms; never delete and re-add.** A re-created term gets a new GUID and orphans the
  Abbreviation, Folder Map and Group Map rows at once.
- **Ampersands must be fullwidth `＆`**, not `&`.
- **Subunits are authored UNDER the unit term**, never in a term set of their own — subunits are
  unit-specific, and a flat set offers every unit the same list. ⚠ The dropdown source for this is
  **not built yet** (§15).
- **Term-set depth must equal the permissioned tier count** (2 for a Head Office: Department, Unit).
  A mismatch makes reconciliation ACL the wrong tier, with no error.

**Verify:** read each set back and record its GUID —

```
/_api/v2.1/termStore/sets/{guid}/terms
```

**GUIDs are per-site.** Every GUID in `CLAUDE.md` belongs to ClarenceDMSTesting and is wrong here.
They go into `CRS Config` in §9; the code fallbacks are only a safety net.

*(Writes from our own UI would use legacy CSOM — the v2.1 endpoint refuses browser writes with
`403 notAllowed`. Relevant only once the Term Manager is built.)*

---

## 4. Lists

Create as `CRS <suffix>`.

| List | Key columns |
|---|---|
| `CRS Config` | `Title`, `ConfigType`, `SettingValue`, `ModeLabel`, `Category`, `TermSetGuid`, `StagingFolder`, `SortOrder`, `Levels` (Note), `PendingLevels` (Note), **`AllowedFileTypes` (multi-select Choice)** |
| `CRS Group Map` | `GroupId` (Number), `GroupName` (Text), `Segment` (Text), `UnitTermGuid` (Text), `Role` (**Text**), `Scope` (Text), `Target` (Text) |
| `CRS Folder Map` | as written by the Folder Map web part |
| `CRS Term Abbreviation` | keyed by term GUID; `Title` holds the term's label |
| `CRS Audit Log` | **self-provisioned by the Audit Log page — do not hand-create** |
| `CRS Requests` | `RequestType` (**Text**), `Status` (**Text**), `RequestedAt` / `ExpiresAt` (DateTime, ISO) |

**`Role`, `RequestType`, `Status` and `EventType` must be TEXT, never Choice.** Writing a value
absent from a Choice column's `Choices` fails the **whole** write, silently — so the day someone adds
a value in code, every row of that kind is lost with no error anywhere.

**`AllowedFileTypes` is the single source of truth for allowed extensions** and is required.
Choices: `.pdf .doc .docx .xls .xlsx`. Keep **FillInChoice disabled**. Nothing ticked is a deliberate
hard block; the column absent falls back to code defaults plus an admin-only warning.

**Audit Log permissions are a manual step.** Only Owners and the flow account may write.
Provisioning deliberately does not set this — breaking inheritance means naming the service account,
and a wrong guess locks the flows out of the list they write to.

**Verify:** every list resolves by title, and `AllowedFileTypes` returns its `Choices`.

---

## 5. Libraries

Four, in two pairs:

| Title | Create as | Then retitle to |
|---|---|---|
| Approval Document | `ApprovalDocument` | `Approval Document` |
| Documents | *(built-in)* | — |
| HC Approval Document | `HCApprovalDocument` | `HC Approval Document` |
| HC Documents | `HCDocuments` | `HC Documents` |

**Create without the space, then retitle**, so the URL stays clean. A list's title and its URL are
independent and only one of them fails loudly: a wrong title 404s; a wrong URL segment makes
`split("/X/")` return a one-element array, so the caller reads `undefined`, routes the file nowhere,
and **logs success**.

### 5.1 Columns — all four libraries, identical internal names

Internal names derive from the title a column is **created** with, once, permanently. Create under
the name that yields the required internal name, then rename.

```
Document_x0020_Type        <- create as "Document Type"
Year                       <- create as "Year"  (NOT "Year/Period")
DocumentDate               <- create as "DocumentDate", then rename to "Document Date"
Confidentiality_x0020_Level
LegallyPrivileged          <- Yes/No
Remark                     <- a dedicated column, not _ExtendedDescription
Full_x0020_Name            <- create as "Full Name"
Vendor_x002f_CustomerName  <- create as "Vendor/CustomerName"
Business_x0020_Segment / BusinessSegmentTid
Department             / DepartmentTid
Unit                   / UnitTid
```

Plus one `<Tier>` + `<Tier>Tid` pair for every tier of every segment onboarded
(`Region`/`RegionTid`, `EstateMill`/`EstateMillTid`, …). The Add-segment tool creates these in all
four libraries automatically.

**⚠ Column parity across all four is load-bearing, and its absence produces two failures that look
nothing like a column problem:**

1. **Bulk upload tags nothing.** One unknown field name fails the whole `validateUpdateListItem`
   call — every column lost, not just the missing one. Shows only as a "No tags" badge per file.
2. **Auto-route drops metadata on every approved document.** SharePoint's copy carries only columns
   that exist at the destination; the rest vanish with no error and a green run.

**Verify:** diff `/fields?$select=Title,InternalName` across all four. A matching display name over a
*different* internal name fails exactly like an absent column, and looks right.

### 5.2 Content approval

| Library | Require approval | Draft Item Security |
|---|---|---|
| Approval Document | **Yes** | Only users who can approve items (and the author) |
| HC Approval Document | **Yes** | Only users who can approve items (and the author) |
| Documents | **No** | — |
| HC Documents | **No** | — |

**Approval left ON in `Documents` is a live outage that presents as a permissions bug.** Every item
Auto-route creates there arrives Pending and is invisible to read-only viewers; the uploader's own
approval email link is denied. Tell-tale without a query: the view bar shows *Approve/reject Items*
and *Show All Files*.

### 5.3 `CRS Folder` content type

Site content type, group `CRS Content Types`, **parent group `Folder Content Types`, parent
`Folder`**. A Document or Item parent will not attach to a folder.

Add it to **all four** libraries with the `Full Name` column on it, and hide it from the New button.
On ClarenceDMSTesting the HC pair was missed and reconciliation warned on every run.

### 5.4 Indexes

- `Documents` → index **Created By**. Past 5,000 items My Submissions' `AuthorId` filter starts
  failing.
- `CRS Audit Log` → `EventTime`, `EventType`, `ActorEmail`, `ItemUniqueId`.

---

## 6. Permission levels

| Level | Must contain |
|---|---|
| `CRS Upload` | Add Items + View |
| `CRS Approve` | **Approve Items** — draft-item security and the whole approval flow key off this |
| `CRS Delete` | Read + Delete Items |
| `CRS Share` | per the deletion-and-share spec |

**Verify `CRS Approve` actually contains Approve Items** — never infer it from the name. Note that
Approve Items cannot be separated from Edit Items; that is a SharePoint constraint, not a mistake.

**Open question carried from testing:** whether `CRS Approve` also includes **Add Items**. The
client's rule is that a Head of Unit uploads as well as approves. If the level has no Add Items, the
Head of Unit persona's `UPL` role supplies it — which is another reason to use the persona rather
than a hand-made approver group (§11).

---

## 7. Site entry group

Create **`CRS_SITE_MEMBERS`** and grant it **Read at site level**. Everyone who uses the system joins
it, in addition to their unit group.

Without it a user granted only a folder can reach that folder by direct link and is **denied on
Home** — which reads as a broken site.

Reconciliation asserts the rest on every run: `CRS_SITE_MEMBERS` holds Read on `Documents` and
**nothing** on `Approval Document`, `HC Approval Document` or `HC Documents`. Do not hand-adjust
those — let it do it, and read the log.

---

## 8. Pages

One page per web part. The landing page resolves links from Site Pages **by pattern**, so names
matter:

| Page | Web part |
|---|---|
| `CRS-Settings.aspx` | CRS Settings (the directory) |
| `Upload-Form.aspx` | Form |
| `Approval-Document.aspx` | Approval Document |
| `My-Submissions.aspx` | My Submissions |
| `Folder-Administration.aspx` | Folder Administration (five tabs) |
| `Group-Management.aspx` | Group Management |
| `Site-Access.aspx` | Site Access |
| `Approval-Library-Access.aspx` | Approval Library Access |
| `Page-Access.aspx` | Page Access |
| `Folder-Access.aspx` | Folder Access |
| `CRS-Audit-Log.aspx` | CRS Audit Log |
| `Bulk-Upload.aspx` | Bulk Upload |
| `CRS-Requests.aspx` | CRS Requests |
| Home | CRS Search |

If a name must differ, use the CRS Settings property-pane override (`link_<key>`) rather than hoping
the pattern matches. A name the pattern cannot match fails as a **dead link** — no error, no clue —
on the one page whose job is telling people where to go.

### 8.1 Admin pages — locked automatically, but VERIFY

**Reconciliation now locks these itself** (1.0.129.0, spec
`2026-08-17-admin-page-lockdown-design.md`), so this is no longer a manual step. It breaks
inheritance on every page whose name is `adminOnly`, grants site Owners Full Control, and strips
every other grant — asserted on **every** run, not once:

```
CRS-Settings · Folder-Administration · Group-Management · Site-Access
Approval-Library-Access · Page-Access · Folder-Access · CRS-Audit-Log · Bulk-Upload
```

It runs in §12. Watch the log for `LOCKED to site owners` on the first run and
`already locked — correct` on the second.

**You must still verify, with a non-admin account.** Not as yourself — you are an owner and will see
everything regardless. Performing and verifying are different things, and this is the single most
important verification in this document.

**If a page is missing from the log**, its file name did not match the policy pattern and it is
**still open**. Either rename it to match the §8 table or add the pattern to `pageAccessPolicy.ts` —
do not assume the lock covered it. Five admin pages were in exactly that state until 2026-08-17.

Grant the working pages normally: `Upload-Form` to uploader **and approver** groups,
`Approval-Document` to approver groups, `My-Submissions` to uploader groups.

---

## 9. Config rows

In `CRS Config`:

- **Setting rows** — `termSet_documentType`, `termSet_yearPeriod`, `termSet_confidentiality`,
  `allowedExtensions` (the Choice column does the work), `legallyPrivilegedFor`,
  `hcConfidentialityLevel`, `allowExternalSharing`, `tenantDomains`, and the `recon_*` toggles.
- **Mode rows**, one per segment: `ConfigType=mode`, `Title=mode_<slug>`, `ModeLabel`, `Category`,
  `TermSetGuid` (from §3), `StagingFolder` (the segment's folder code), `SortOrder`, `Levels` (JSON).

`Levels` for a Head Office is
`[{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]`.

**Do not carry over a `stagingLibrary` row reading `Staging`** — it is superseded and would 404.

**`allowExternalSharing` fails CLOSED.** Absent, unreadable, or anything but an explicit yes means
internal only. `tenantDomains` decides what counts as external — **supplying no domains means every
recipient reads as external**, never as internal.

**Verify:** open the Upload Form as an admin. Every segment is offered; an unready one is labelled
*"— not fully set up yet"*.

---

## 10. Abbreviations

Folder Administration → **Term Abbreviations**. Fill in a code for **every** term at every
permissioned tier.

- **A term with no abbreviation is SKIPPED** by reconciliation — no folder, no error, and nobody can
  upload there. This is the silent failure of the whole system.
- Codes must be **unique among siblings**, or two units merge into one folder with one ACL.
  Reconciliation aborts before creating anything if they collide.
- Changing one later **renames a live folder** on the next run.

**Verify:** the page shows no amber *"no code"* rows at any permissioned tier.

---

## 11. Groups and mappings

1. **Group Management** — create one group per unit per persona. Use the suggested names
   (`GHO_GF_CORU_UPLOADER`); the suffix is read back to pre-select a role.
2. **Folder Access** — map each group to its Segment, Tier and **persona**. Picking the persona
   applies its roles; do not hand-pick roles.
3. Add people. Everyone also joins `CRS_SITE_MEMBERS` automatically.

Personas: **C-Level (global / segment)**, **Head of Department**, **Head of Unit**, **PIC**,
**SDG Employee**, plus the HC variants.

**Head of Unit carries `APR`, `DELS`, `UPL`, `DEL`, `SHARE`.** It is the only persona that approves,
and it uploads too. Do not create a separate approver-only group — that is a pre-persona shape, and
it produces someone who can approve but not upload.

**Consequences to state to the client rather than bury:**

- A Head of Unit **approves their own uploads**.
- A Head of Unit is the unit's **sharing authority** — able to share directly, not merely to approve
  other people's share requests.
- Every PIC reads **every approved document in their unit**, at any confidentiality level.
- The **unit is the smallest confidentiality boundary.** Two people who must not see each other's
  documents belong in different units. See `docs/client/document-visibility-within-a-unit.md`.

---

## 12. Reconciliation

Folder Administration → **Folder Reconciliation**. Run it.

- **Budget an hour and do not close the tab.** It scales with folders × four libraries, has no
  resume, and gives no "what remains" summary if interrupted (finding #9).
- Read the log. `⚠ no group-map groups for this unit (locked admin-only)` is expected for units
  nobody has been assigned to yet.
- **Run it a second time.** Every line should read `already locked, skipped` or `— correct`. Anything
  it *does* on the second run that it also did on the first is a defect worth reporting.

**Verify:** the folder trees exist in all four libraries, and a test uploader's unit folder is
reachable.

---

## 13. Power Automate — two flows, both as the SDG account

**Sign in as the SDG account before creating the first action.** The connection is created implicitly
by that action and the identity is then permanent.

Build from `docs/superpowers/specs/2026-08-08-auto-route-flow-and-draft-isolation.md` **verbatim** —
Power Automate config is not in source control and that spec is the only record of it. Several
settings look cosmetic and are not.

| Flow | Trigger condition | Job |
|---|---|---|
| Auto-route | `@equals(triggerOutputs()?['body/{IsFolder}'], false)` | Copy the approved file to `Documents`, stamp Author/Editor/Created, then **delete the source by item id** |
| Folder approval | `@equals(triggerOutputs()?['body/{IsFolder}'], true)` | MERGE `{"OData__ModerationStatus": 0}` |

**The polarity is the whole thing, and both mistakes have already been made once.** Set Auto-route to
`true` and no file is ever routed — silently, because a flow that never fires leaves no run history.
Omit the condition and every ensure-created folder triggers Auto-route, which has already copied a
file into `Documents` while it was still *Waiting for Approval*: approval bypassed.

**The delete is load-bearing for security, not housekeeping.** An approved file left in the approval
library is visible to every PIC in the unit, which defeats draft isolation.

**Power Automate header keys must NOT include the colon.** The key box wants `Accept`, not `Accept:`.
With the colon the header does not exist, and the symptoms look unrelated to one another: responses
come back `odata=verbose` so every `body('X')?['Field']` is null; `validateUpdateListItem` reports
`HasException: false` and changes nothing; a MERGE goes as a plain POST. Read the action's raw
**Inputs** — the colon is visible there and nowhere else.

**Never conclude a field is unwritable from a clean response. Re-read the item.**

If HC is in scope, build the **two HC equivalents** with the same polarity. Nothing HC works without
them.

---

## 14. Verification — before handing over

Not optional, and **not doable as yourself**. Use two real accounts: one uploader, one Head of Unit.

1. Uploader opens the Upload Form — **only their own path is offered**.
2. Upload one PDF with Year, Document Type and a Remark.
3. Head of Unit opens Approval Document, sees the file, approves it.
4. `Documents/<SEG>/<DEPT>/<UNIT>/<Year>/<Doc Type>` contains it, **Created By is the uploader**, and
   the Remark survived.
5. It is **gone** from Approval Document.
6. A second uploader in a different unit **cannot see** any of it.
7. A non-admin **cannot open** any page listed in §8.1.
8. An uncleared user is **never shown** the Highly Confidential level.

Steps 4, 5 and 7 are the ones that have failed before. **This whole sequence has never yet been run
end to end on any site** — expect to find something.

---

## 15. Known gaps that ship with this

State these to the client rather than letting them be discovered.

| # | Gap | Effect |
|---|---|---|
| #1/#6 | Admin pages are not auto-restricted | §8.1 is manual; skipped, every uploader can open every admin tool. |
| #8 | **SubUnit is not built** | Terms authored under a unit are safe to create, but the dropdown reads from a term-set id and will be **empty**. Tell the client to author the terms; do **not** tell them it works. |
| #9 | Reconciliation ~1 hour, single tab, no resume | Operational, not correctness. This site is larger than the test site. |
| #12 | Group Map caches a group's NAME | Renaming a group makes every screen and every log line show a name that no longer exists. Grants stay correct. **Renaming groups at import triggers this across the board.** |
| — | HC documents already filed | Migrating existing Highly Confidential documents is **out of scope**. Until a sweep runs, HC protection applies only to documents filed after deployment. |
| — | Share revoke | Not built. Every approved share creates a permanent unique permission scope. |
| — | CRS Search managed properties | `<InternalName>OWSTEXT` / `OWSDATE` is an **unverified assumption**. Test the search filters live; fall back to mapping `RefinableString00`–`99` in Site Settings → Search Schema, which a site collection admin can do without tenant access. |
| — | Audit log | Records **admin** activity only until the flow-side writers are built. Do not read an empty file history as "nothing happened to that file". |

---

## 16. Rollback

Nothing here destroys data if it goes wrong part-way:

- Removing the app leaves every list, library, folder, permission and document in place.
- Reconciliation only adds and corrects. It never deletes a folder whose term is missing.
- The one destructive action in the system is **Retire a segment** with the folder-delete option
  ticked, which recycles rather than purges and is restorable for 93 days.

The genuinely hard-to-reverse steps are the **Power Automate flows** — a wrong `{IsFolder}` polarity
can route an unapproved file — and **granting a group access to the wrong folder**. Verify §13 and
§11 before letting real documents through.

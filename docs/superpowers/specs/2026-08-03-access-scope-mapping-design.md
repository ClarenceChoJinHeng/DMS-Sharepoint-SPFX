# Access Scope Mapping — Site, Library, Folder, Page — Design

**Date:** 2026-08-03
**Status:** BUILT. `Scope`/`Target` are live — reconciliation reads them (`FolderManager.tsx` library and page passes) and `PageAccess.tsx` writes `Scope: "Page"` rows. Header corrected 2026-08-19; it had said "not yet implemented" long after shipping.
**Related:** `2026-07-27-site-entry-access-layer-design.md` (why a site-entry layer exists at
all — this spec automates its manual step), `2026-07-22-group-map-builder.md`, and the Phase 1b
notes in `2026-07-28-client-site-migration-runbook.md`
**Not related to:** `2026-07-30-audit-log-and-access-matrix-design.md`, which *reports* who has
access. This one *grants* it.

---

## 1. What changes

A `DMS Group Map` row means exactly one thing today: **this group gets this role on the folder
for this term**. Everything else — site entry, library-level grants, who may open a page — is
done by hand in native SharePoint, on screens the client has to be walked to.

That is the actual complaint. Not that native SharePoint cannot do it, but that it is four
different screens, none of them beside the mapping that motivated the change, and nothing
records that the step was done. The site-entry group is the proof: reconciliation warns
`CRS_SITE_MEMBERS not found — run "Set up site entry" first`, and "Set up site entry" is a
sentence in a runbook, not a button. A step that exists only in prose gets skipped.

So: one list, one page, four scopes.

## 2. Data model

Two new columns on `DMS Group Map`:

| Column | Type | Holds |
|---|---|---|
| `Scope` | Choice: `Site` \| `Library` \| `Folder` \| `Page` | what the row grants access to |
| `Target` | Text | the library title or the page's file name; empty for `Site` |

`Segment` and `UnitTermGuid` stay, and remain meaningful **only for `Folder`**.

**A blank `Scope` reads as `Folder`.** Every row that exists today is a folder row, and the
column arrives after them. Defaulting blank to `Folder` means an existing site keeps working
before anyone edits a single row — the alternative is a site whose folder permissions silently
stop being applied the moment the column is added.

## 3. Scope semantics

**Site.** Grants the group Read on the web. No term, no target. This is the layer that lets
someone reach the site root at all; without it a folder grant gives only Limited Access, which
reaches the folder by direct link and denies the Home page (memory
`dms-two-layer-access-site-plus-folder`).

**Library.** Breaks the library's inheritance if it is still inheriting, then grants the role on
the library. `Target` is the library title.

**Folder.** Unchanged. The existing per-term behaviour.

**Page.** Breaks inheritance on **one item in `Site Pages`** and grants Read. `Target` is the
page's file name (`Upload.aspx`). Deliberately not the library — locking the whole `Site Pages`
library takes the Home page with it.

## 4. Site entry stops being a manual step

Reconciliation gains a first pass that:

1. finds or **creates** the site-entry group (`CRS_SITE_MEMBERS`),
2. grants it **Read on the web** if it does not already hold a role there,
3. keeps the existing membership sync — every member of any `CRS_*` group is added to it.

Steps 1 and 2 are new; step 3 exists today and does nothing useful when 1 has been skipped.

This pass runs **before** any folder is locked. Ordering is not cosmetic: locking folders first
and granting site entry afterwards leaves a window in which a correctly-provisioned uploader can
reach nothing, and if the run dies in that window, that is the state it stays in.

> **Not the M365 group's Members group.** Its permission level *can* be changed, but only from
> the classic `/_layouts/15/user.aspx` page — and that is not the reason to avoid it. Its
> membership is owned by the M365 group: anyone added from Teams, Outlook or the admin centre
> lands in it automatically and cannot be gated. A group we do not control, holding a site-wide
> grant, inverts a model built on deny-by-default. Use a dedicated group, set as the associated
> **Visitors** group via `/_layouts/15/permsetup.aspx`.

## 5. Rules

- **Only ever add a role assignment; never remove one.** Already true of the folder pass, and it
  stays true. Removing assignments is how a tool locks its own administrator out of the site,
  and at site scope there is no folder-level escape hatch to recover through.
- **Deleting a row does not revoke anything.** Also already true, and a sharper edge at site and
  library scope than at folder scope because the blast radius is larger. §7 covers what happens
  instead.
- **Never restrict the home page.** A `Page` row targeting the site's welcome page makes site
  entry worthless: users get Read on the web and then 403 on the only thing they can navigate
  to. Reconciliation refuses that target by name rather than reporting it after the fact.
- **Page permissions are not a document boundary.** Hiding the page that hosts the upload form
  does not hide the library behind it. This is navigation hygiene; folder ACLs remain the
  security model. Worth stating in the UI, because "only specific people can access the page"
  reads like a stronger guarantee than it is.

### 5a. What a Page row does and does not guarantee

Asked directly during the design review, and the answer belongs in the spec rather than in a
chat log.

**It does hold.** Item-level permissions on a `Site Pages` item are enforced server-side. A
non-member opening the URL gets Access Denied; the page is trimmed from search results; and it
disappears from the Site Pages library listing, so its very existence is hidden. This is a real
boundary, not a cosmetic one.

**Four limits to state to the client before promising anything:**

1. **It hides the page, not the data on it.** Documents shown by a web part remain reachable by
   direct library URL, search, the mobile app and sync. What actually stops that is the folder
   ACLs, which are independent and already in force — the web parts run as the signed-in user,
   so someone who *did* reach a page still sees only what their folder permissions allow. Page
   permissions are the outer of two layers, and the weaker one.
2. **Hand-written navigation links are not permission-trimmed** against their target in modern
   SharePoint. A restricted page still shows its nav link, and clicking it gives Access Denied.
   Not a leak — but it reads as a bug, so it will be raised.

   Hiding the link needs **audience targeting** on the navigation link, and audience targeting
   accepts only **Azure AD security groups or M365 groups**. It does **not** accept SharePoint
   site groups, which is what every `DMS_*`/`CRS_*` group in this system is. So the blocker is
   the *kind* of group, not the client's willingness to grant group-creation rights — worth
   getting that distinction right before saying it out loud, and worth verifying in the tenant
   by opening the nav editor and checking whether a `DMS_` group appears in the audience picker
   at all.

   **The remedy that needs no groups: remove the link.** The page stays reachable by direct URL
   for those permitted, and everyone else neither sees it nor can open it. For the admin pages
   (Folder Manager, Bulk Upload) that is better than targeting anyway — administrators bookmark
   them, and there is no link for anyone else to click. If the client later wants a genuinely
   audience-targeted link, the ask is one Azure AD security group, which is a far smaller
   request than open group-creation rights.
3. **Site collection administrators and the site Owners group always retain access.** "Only
   specific people" means "only specific people, plus administrators".
4. **The home page is excluded** — see the rule above.
- **Break inheritance before granting**, at every scope — the order the folder pass already
  uses. Granting first and breaking second discards the grant.
- **Re-add the site-entry group with Read after breaking a LIBRARY.** Found while planning the
  implementation on 2026-08-04, and it is the sharpest edge in this spec.

  The approval guard resolves the destination folder in Documents **as the signed-in
  approver** (role-personas spec §10). That works only because `DMS_SITE_MEMBERS` holds Read
  on the Documents library *by inheritance from the web* — approvers are deliberately not in
  any Documents folder group. A `Library` row for `Documents` breaks that inheritance with
  `copyRoleAssignments=false`, which discards the inherited Read, and then **every approver on
  the site 404s on every destination folder and approval is refused** — while the
  reconciliation log reports a clean, successful run.

  So the library pass re-adds **two** principals after every break, not one: site Owners with
  Full Control (already required above) and the site-entry group with Read. The second is not
  a convenience; it is what keeps approval working.

  Stated as a rule rather than left to the implementer because the symptom is maximally
  misleading: approval breaks for everyone, nothing in the log mentions it, and the change
  that caused it was made against a different scope entirely.
- **`copyRoleAssignments=false`, always.** This single parameter decides whether the feature
  does anything. With `true`, breaking inheritance copies every inherited grant forward, so the
  page stays visible to exactly the same people — and the run reports success. The failure is
  invisible from the log; the only symptom is that nothing changed. `clearSubscopes=true` with
  it, matching `breakInheritance` in `FolderManager.tsx`.
- **Re-add the site Owners group after every break.** The corollary of the rule above: once
  inheritance is broken with `copyRoleAssignments=false`, the only remaining access is site
  collection administrators plus whatever is explicitly granted. An owner who is not also a site
  collection admin loses the page. The folder pass already re-adds `ownerGroupId` with Full
  Control; the library and page passes must do the same.
- The site's root web has unique permissions by definition, so the `Site` pass grants directly
  with no break step. A library and a page each need the break.

## 6. UI

`Scope` becomes a filter above **Existing mappings**: `All · Site · Library · Folder · Page`,
with counts, mirroring the log tabs on the reconciliation screen.

The add-mapping form gains a scope picker **first**, because the scope decides which of the
other fields are meaningful: `Site` hides segment, tier and target; `Library` and `Page` swap
the term cascade for a target picker; `Folder` is the form as it stands.

The tab is labelled **Page Access**, not "Web Part access". It breaks inheritance on Site Pages
items; a web part is not what it acts on, and a log line saying otherwise sends the reader
looking in the wrong place.

## 7. Drift, not revocation

Because nothing is ever removed, the list and the site can disagree: a grant made by hand in
native SharePoint has no row, and a deleted row leaves its grant in place.

Reconciliation reports both and changes neither:

- a principal holding a role at a mapped scope with no row → `UNMAPPED GRANT`
- a row whose grant is absent → applied, as now

Same posture as the unclaimed-folder report in
`2026-08-02-term-guid-orphan-repair-design.md` §8, for the same reason: the tool is not allowed
to assume a human's manual grant was a mistake.

## 8. Out of scope

- **Audience targeting.** It hides a web part on a page; it denies nothing. Putting it in a list
  called "access" invites someone to rely on it.
- **Revocation.** Stays manual, stays documented.
- **Breaking inheritance on the `Site Pages` library itself.** Per-page only.
- **Anything below a page** — web part instances have no permissions of their own.

## 9. Verification

Needs the tenant.

1. Fresh site, no site-entry group: run reconciliation, confirm the group is created, holds Read
   on the web, and that a plain uploader can open Home.
2. Confirm the site-entry pass ran **before** the first folder lock — check the log order, not
   just the outcome.
3. Add a `Library` row for `Documents`, reconcile, confirm the library's inheritance is broken
   and the group holds the role.
4. Add a `Page` row for a non-home page, reconcile, confirm a non-member gets 403 on that page
   and the rest of the site still works. Check as a **non-member who previously could open it** —
   a break that copied its role assignments forward looks identical in the log and only differs
   here. Also confirm the page is gone from that user's search results and from their Site Pages
   listing, and that a site Owner who is not a site collection admin can still open it.
5. Add a `Page` row for the home page and confirm the run **refuses** it by name.
6. Grant a group Read on a library by hand, reconcile, confirm `UNMAPPED GRANT` is reported and
   the grant is left alone.
7. Delete a `Site` row, reconcile, confirm access is **not** revoked and the log says so.
8. On a site with existing rows and no `Scope` values, confirm every row still applies as
   `Folder`.

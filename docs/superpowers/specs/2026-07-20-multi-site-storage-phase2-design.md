# Multi-Site Storage (Phase 2) — Design

**Status:** SUPERSEDED / RETIRED 2026-07-26 — see `2026-07-26-cross-site-upload-retirement.md`.
A new site collection draws from the same tenant storage quota, so multi-site saves no space;
the DMS stays single-site. Kept for historical context only.

## Problem

A single SharePoint site collection caps at **25 TB**. The client fears one or more business
segments' documents will exceed that. Solution: split **some** segments onto their own sites
(each with its own 25 TB), while users continue to work from the central **DMS site**.

Goal: upload via the existing form (same group rules, same term-store values), have approved
files stored on the segment's own site, and let users **find and view** those files from the
DMS site.

## Chosen architecture — B: Central Staging, route approved files out

The 25 TB is consumed by **approved documents accumulating over years** — the **Documents**
library — not by **pending** uploads, which are transient (they leave Staging on approval).
So only approved documents need to live on segment sites.

```
Upload form (UNCHANGED)
  -> DMS site Staging (central, pending)          <- existing per-folder group security
  -> approval (central, UNCHANGED)
  -> auto-route flow copies APPROVED file
        to the SEGMENT site's Documents library    <- new: cross-site, server-side
  -> users view via cross-site browser web part on the DMS site
```

### Why B over A (upload direct to segment site)

| | A: direct-to-segment | **B: central Staging (chosen)** |
|---|---|---|
| Upload form | cross-site rewrite | **unchanged** |
| Staging library | one per segment site | **one, central** |
| Approval | duplicated per site | **central, unchanged** |
| Segment site needs | Staging+Documents+approval+flow+groups | **Documents + groups only** |
| Cross-site write | from web part (per-user perms/digest) | **by the flow (elevated service account)** |

B keeps the whole upstream pipeline intact and does the one cross-site action (moving an
approved file) server-side in the flow, where it already runs elevated.

## What stays UNCHANGED

- Upload form (`Form.tsx`), its per-mode cascade, field order, rename-proof folder routing.
- **DMS Group Map** + tiered hard-check (group -> term authorization). Term store is tenant-wide.
- **DMS Staging** library, its columns, and per-unit folder ACLs.
- **Content approval** — approvers keep working in the one central Staging library.
- **DMS Folder Map** + Reconciliation (they serve the DMS Staging upload routing, untouched).

## Security model (confirmed)

Visibility in a library is governed by **security groups assigned on the folder**:

- **Central Staging (pending):** already correct — inheritance broken per unit folder; unit
  **uploader** group = Contribute (sees own items via draft-item-security), **approver** group
  = Design (sees all pending in that folder). No change.
- **Segment site Documents (approved):** provisioning must assign the unit's **reader** group
  (and any approver/HC groups) on the corresponding folders, with inheritance broken per unit
  folder — mirroring the DMS Documents model, just on the segment site. This is the new place
  ACLs are applied. Reuse the same DMS_* security groups (matched by Object ID) — assign them
  on the segment site's folders.

## New components (the Phase 2 build)

### 1. Segment -> Site routing config
A mapping of **which segment routes to which site + Documents library**. Options:
- Extend **DMS Config** modes with an optional `DocumentsSiteUrl` (blank = local DMS site).
- Or a small `DMS Site Map` list: `Segment` (term-set GUID) -> `SiteUrl`, `DocumentsLibrary`.
The auto-route flow reads this to decide the destination. A blank/absent entry means the segment
stays entirely local (today's behaviour) — so the split is opt-in per segment.

### 2. Auto-route flow change (the core of Phase 2)
The existing auto-route flow (fires after content approval, copies approved file to Documents by
**live folder path** — rename-proof) is extended to:
- Look up the file's **segment** -> destination **site URL + Documents library** from the config.
- If a remote site: build the destination folder path on the **segment site** (same
  Unit/Year/DocType structure), **ensure-create** folders there, and **Copy/Create file**
  cross-site (SharePoint connector supports a per-action site address).
- If no remote mapping: current local behaviour, unchanged.
- Preserve metadata columns on the destination (same internal names must exist on the segment
  site's Documents library).
- Keep the existing trailing-slash / "create folder" handling.

### 3. Segment site provisioning (scripted)
A repeatable step (PnP site template, or an extension of the Onboarding web part) that stamps a
new segment site with:
- A **Documents** library with the **identical columns** (exact internal names + types — see the
  Phase 1 `/fields` capture; a missing/renamed column will fail the flow's metadata write, like
  the Vendor-column incident).
- The **folder scaffold** (segment -> dept -> unit), inheritance broken per unit folder.
- **Group ACLs** — assign the unit's reader/approver/HC DMS_* groups (by Object ID) via the
  `c:0o.c|federateddirectoryclaimprovider|{GroupId}` claim.
Scripted chosen (over manual runbook) so every segment site is guaranteed consistent.

### 4. View surface — how users reach split-out segments from the DMS site

**Chosen (MVP): native SharePoint Link items.** No code. In the DMS **Documents** library,
segments that stay local remain **real folders**; each split-out segment is added as a native
`+ New -> Link` item pointing at that segment site's Documents library (or a folder within it).

```
DMS site — Documents library
  📁 Group Head Office        <- real folder, files stored locally
  🔗 Upstream Operations      <- Link -> https://.../sites/Upstream/Shared Documents
  🔗 SDGI                     <- Link -> https://.../sites/SDGI/Shared Documents
  🔗 I&T                      <- Link -> https://.../sites/IandT/Shared Documents
```

Clicking a link navigates the user to that segment site, where they browse the real folders/files.
- **Zero code, fully native**, always in sync with SharePoint's own UI.
- **Permissions carry over unchanged** — the link is just a URL; the user needs **Read** on the
  segment site (the same DMS_* group reuse) or they hit access-denied after the jump.
- **Opt-in per segment** — only split-out segments become links; local segments stay real folders.
- **Trade-off:** a visible context switch — the user lands on the segment site (different URL/nav),
  not an in-place browse. Acceptable for the Phase 2 goal ("find and view" from the DMS site).
- **Prerequisite:** confirm `+ New -> Link` is enabled for the library (can be disabled at
  tenant/library level).

**Later (optional): productionize the `CrossSiteBrowser` POC** into a supported web part for a
seamless in-place experience — only if the client wants to avoid the context switch. It would:
- Read a segment site's **Documents** library (folder tree + open-in-place) on a DMS-site page.
- Drive the target site(s) from the Segment->Site config rather than hand-typed properties.
- Optionally aggregate multiple segment sites in one view (tabs or a segment picker).
- Theme-aware styling, empty/error states, and paging for large folders.
- Opening a file is still authorised by the user's **Read** on the segment site (group reuse).

## Open decisions to finalise at build time

- **Config home:** extend DMS Config modes vs a dedicated `DMS Site Map` list. (Lean: dedicated
  list — clean separation, easy for admins.)
- **Cross-site move mechanism:** Power Automate SharePoint connector vs an Azure Function for
  very large files (tens of GB) or high volume.
- **View surface:** DECIDED — native Link items for MVP (see #4); optional web part later.
- **HC files:** if a Highly Confidential file is approved, where is it secured — does the segment
  site's Documents folder get the HC ACL applied by the flow (reuse the HC securing design)?
- **Rename-proofing the destination:** the flow routes by live path on the segment site; confirm
  whether a segment-site folder map is needed or path routing suffices (as it does today).

## Prerequisites / sequencing

1. **Phase 1 verified live** (form upload + tag + tiered detection + security proof-test).
2. Remove the temporary `CrossSiteBrowser` POC (or fold its proven read/upload code into the
   productionized view web part).
3. Build order: provisioning script -> Segment->Site config -> flow change -> view web part ->
   end-to-end test on one real segment site -> roll out remaining split-out segments.

## Risks / watch-items

- **Column parity:** any column mismatch on the segment site's Documents library breaks the
  flow's metadata write. Provisioning must stamp exact internal names.
- **Permissions drift:** the DMS_* groups must be assigned on the segment site too; a user with
  no rights there cannot open the file even though the DMS-site browser lists it.
- **Flow throughput:** large files / high volume may exceed Power Automate limits -> Azure Function.
- **Pending backlog (edge case):** if approvals lag badly, central Staging could grow; monitor,
  but pending files remain transient by design.

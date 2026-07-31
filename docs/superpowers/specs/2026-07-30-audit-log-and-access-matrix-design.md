# Audit Log & Access Matrix — Design

**Date:** 2026-07-30
**Status:** proposed — design agreed in discussion, capability NOT yet verified
**Backlog:** closes `.claude/backlog.md` → "Audit Logging" and `requirements.md` §12
**Related:** `2026-07-26-in-site-term-store-management-design.md` (same client access constraint),
`2026-07-23-share-guard-retirement.md` (why the package is Graph-free)

---

## Problem

The client asked for two things and called them both "audit":

1. **Activity log** — download, upload, approve, deletion, sharing.
2. **Access matrix** — access granted to a particular user at folder and file level.
   *"Folders that Anil as head of department has access to, and files that have been shared with Anil."*

These are **different problems with different answers**, and conflating them is the main risk in
this piece of work.

- (1) is a log of **events over time**.
- (2) is a report of **current state**.

No single source answers both. Purview answers (1) and cannot answer (2). Our own data answers
(2) and only partially answers (1).

## The access constraint (unchanged from the term-store decision)

The client refuses M365 Admin Center access on security grounds and does not want a security-team
ticket per change. Any design depending on a tenant-level role grant is likely to stall.

**Critical correction to a common assumption:** SharePoint Administrator does **not** grant audit
log access. Neither does Site Collection Administrator. The unified audit log is gated by the
`Audit Logs` / `View-Only Audit Logs` roles in Exchange Online (least-privilege role group:
**Audit Reader**). It is a different role family from anything SharePoint-side.

| Role | Unified audit log access |
|---|---|
| Regular user / member | ❌ |
| Site Owner | ❌ |
| Site Collection Administrator | ❌ |
| **SharePoint Administrator** | ❌ |
| Audit Reader | ✅ read-only, search + export |
| Compliance Admin / Security Reader / Global Reader | ✅ (plus much else) |
| Global Administrator | ✅ |

The unified audit log **cannot be scoped to a single site** — it is one tenant-wide store covering
Exchange, Teams, OneDrive and every SharePoint site. So "Audit Reader for just the DMS site" does
not exist, and the security team's likely objection is legitimate rather than obstructive.

---

## Part 1 — Activity log

### Per-event capability

| Event | In-app custom log | Notes |
|---|---|---|
| Upload | ✅ complete | Already in library data (`Author` + `Created`); log row adds metadata context |
| Approve / reject | ✅ complete | Flows through our ApprovalDocument page — we control the write |
| Edit (metadata / rename) | ✅ complete | Version history (`Editor`, `Modified`) + log row |
| Delete | ✅ good, 93 days | `/_api/web/RecycleBin` exposes `DeletedByUser`, `DeletedDate`, `DirName`. SCA sees the whole site. Blind after purge from 2nd-stage bin. |
| Share | ⚠️ prevent, don't log | See "Sharing" below |
| **Download** | ❌ **impossible** | See "Downloads" below |

### Downloads — why custom cannot work

SPFx runs **in the browser as the signed-in user**. To log a download we must be in the path of
every download. A custom document browser that logs on its own download button is trivially
bypassed: the user still holds Read on the library, so the native library URL, SharePoint search,
Open in Office, and OneDrive sync all reach the file without touching our code. We cannot remove
that Read access, because our own web part needs it to fetch the file on the user's behalf.

Closing the hole requires a server-side component with app-only credentials serving files the user
cannot reach directly — a different architecture needing admin consent, i.e. exactly the dependency
removed in the Share Guard retirement.

**Decision: do not build a download log.** A bypassable download log is worse than none — it
produces a report that looks authoritative and is not. Downloads require Purview or nothing.

Note for whoever builds the Purview side later: "download" is **not one event**. `FileDownloaded`
is the obvious one, but Open-in-Office logs `FileAccessed` and OneDrive sync logs a separate sync
event. A report keyed on `FileDownloaded` alone will under-report.

### Sharing — prevent rather than audit

Under the existing RBAC (Uploader = Contribute, Reader = Read, Approver = Design) **no ordinary
user holds Manage Permissions**, so nobody can grant access directly. The only remaining vector is
sharing *links*.

A **site owner** can close that: Site permissions → Sharing settings → *"Only site owners can
share files, folders, and the site."* No admin, no ticket.

**Decision: disable non-owner sharing.** This converts an unauditable event into a prevented one,
and the residual question ("who can reach this file?") is then fully answered by Part 2. Stronger
than a sharing log: not "here is a record of sharing that happened" but "sharing cannot happen;
here is exactly who has access and why."

### The log list

`DMS Activity Log` — a standard SharePoint list.

| Column | Type | Purpose |
|---|---|---|
| `Timestamp` | DateTime | Event time |
| `Actor` | Person | Who did it |
| `Action` | Choice | Upload / Approve / Reject / Edit / Rename / Delete |
| `FileName` | Text | Name at event time |
| `FilePath` | Text | Server-relative path at event time |
| `FileUniqueId` | Text (GUID) | **Stable join key** — survives rename and move within the site |
| `BusinessSegment` / `Department` / `Unit` | Text | Denormalized level labels |
| `DocumentType` / `YearPeriod` / `Confidentiality` / `Vendor` | Text | Denormalized metadata |
| `Details` | Note | Rejection reason, old→new name, error text |
| `Result` | Choice | Success / Failure |

Synthetic row:

```
Timestamp:        7/30/2026 14:32
Actor:            A. Kumar
Action:           Approve
FileName:         Invoice-0001.pdf
FilePath:         /sites/ClarenceDMSTesting/Staging/GHO/Group Finance/...
FileUniqueId:     00000000-1111-2222-3333-444444444444
BusinessSegment:  Group Head Office
Department:       Group Finance
Unit:             Accounts Payable
DocumentType:     Invoice
YearPeriod:       2026
Result:           Success
```

**Date format:** `Timestamp` is written as `M/D/YYYY` (US site locale) — CLAUDE.md rule #1. ISO
`YYYY-MM-DD` is rejected. Reuse `toSpDate()` from `Form.tsx`.

`FileUniqueId` is consistent with the existing architecture — DMS Folder Map already routes by
UniqueId.

### Design rule: denormalize, never look up

Copy metadata values **into** the log row at write time. Do not store a lookup to the document.

Three reasons, all of which will occur in production:
- the document may be **deleted** later — a lookup goes blank;
- the metadata may be **edited** later — the log must show what was true *then*, not now;
- the auto-route flow **moves** files from Staging to Documents on approval — a path-based lookup
  breaks by design.

A log row must be immutable and self-contained or it is not a log.

### Why custom beats Purview for these columns

Purview records raw paths and has no concept of Business Segment or Document Type. Answering
"all approvals in Group Finance where Document Type = Invoice" from Purview means parsing folder
paths out of URLs and reverse-engineering metadata — and that fails in this architecture because
approved files move, so the logged path may no longer exist.

The custom log writes the metadata at the moment of the event, so the client's requested columns
are just columns: filterable, groupable, exportable.

### Tamper resistance

If uploaders hold Contribute they can edit or delete their own log rows, which voids the log.
The list needs a **custom permission level with Add Items + View Items only** (no Edit, no Delete),
created by a site owner and assigned on `DMS Activity Log` with inheritance broken.

### Viewer

**v1 = no custom viewer.** It is a normal SharePoint list, so filtering, grouping, saved views,
Excel export and Power BI come free. Ship configured views — *Approvals by Segment*, *Deletions
this month*, *Activity by user* — and only build UI for what a list view genuinely cannot do.

Set at creation, painful to retrofit: **index** `Timestamp`, `Actor`, `Action`, `BusinessSegment`
before the list passes the 5,000-item view threshold, and give the default view a rolling date
filter.

### Honest labelling

Ship this as an **operational activity log**, not an audit log. It records events flowing through
our web parts; recycle bin and version history backstop some of the rest, but not all. To an
auditor, "audit log" means Purview. Do not use the word in the handover for this component.

---

## Part 2 — Access matrix

Buildable now, **no new permissions**, unaffected by whatever happens with Purview.

### Direction A — user → folders ("what can Anil reach?")

The client's literal ask. Mostly a read of data we already own:

1. Resolve the user's SP group memberships via `spGroups.ts` (`/_api/web/...`, no Graph).
2. Map groups → unit term GUID → folder via `DMS Group Map`.
3. Walk term ancestry (`loadTermPath`) to render Segment → Department → Unit.
4. Role (Uploader / Approver / Reader) from the group suffix (`_UPL` / `_APR` / base).

Fast, authoritative, and it reflects the intended design rather than accidental state.

### Direction B — folder/file → users ("who can reach this?")

Read `RoleAssignments` on the folder's `ListItemAllFields`, expanding `Member` and
`RoleDefinitionBindings`.

### Direction C — file-level direct shares

Scan for `HasUniqueRoleAssignments` on items and read their role assignments plus the hidden
`SharingLinks.*` groups. **Expensive** — walking the library live does not scale, so this wants a
scheduled crawl into a list rather than an on-demand scan.

If non-owner sharing is disabled per Part 1, this should return near-empty **by design**. That is
the correct answer and should be presented as such, not as a bug.

### Reconciliation value

Direction A shows *intended* access (Group Map); Direction B shows *actual* access (role
assignments). **Diffing them is the real product** — it surfaces drift, orphaned grants, and
folders whose inheritance was never broken. Worth building as a third view once A and B exist.

---

## Purview — what it does and does not buy

If the client pursues **Audit Reader**:

- ✅ Covers all five list-1 events including downloads and sharing, properly.
- ⚠️ Retention: 180 days (Audit Standard / E3), 1 year (Premium / E5), 10 years only with a paid
  add-on. If compliance means "produce this in three years," the licence matters.
- ⚠️ Latency: typically under an hour, SLA allows up to 24. Not a live view.
- ❓ **Unverified:** whether SharePoint *content approval* raises a distinct audit event or only a
  generic file-modified. Our own log captures approvals cleanly either way — one reason to keep
  the in-app log even if Purview arrives.
- ❌ **Does not answer Part 2.** Purview logs events, not current state. It can say Anil was
  granted access on 12 March; it cannot say what Anil can open today.

**Audit logging is on by default** for E3/E5 tenants (since early 2023), so SDG's history is
almost certainly accruing already. This makes the decision **deferrable at no cost** — Audit Reader
granted in three months still sees the prior 180 days. Verify it is actually enabled rather than
assuming.

---

## Decision

1. Build the **in-app activity log** (upload / approve / reject / edit / delete) — no new access.
2. **Disable non-owner sharing**; do not build a sharing log.
3. **Do not build a download log.** Downloads are Purview-or-nothing; say so plainly.
4. Build the **access matrix** (Directions A and B first, C only if sharing stays enabled).
5. Treat Purview as an **optional upgrade for download visibility**, on the client's own timeline.
   Nothing above is blocked on it.

### Framing for the client

> "Uploads, approvals, edits and deletions we can deliver in-app with no new access. Sharing we'll
> prevent outright rather than log, and the access matrix will show exactly who can reach what.
> Downloads are the one thing SharePoint only records in Purview — if you need them, someone on
> your side needs **Audit Reader** (read-only, least privilege). If you don't, we can start now."

---

## Open — to verify before committing

- [ ] **Site-collection audit reports** (`_layouts/15/Reporting.aspx?Category=Auditing`) — reachable
      by SCA with **zero grants**. May cover deletions/edits/views scoped to this site collection.
      Microsoft has been hollowing this page out for years; it may return empty or error.
      **The only zero-grant option — test before taking anything to the client.**
- [ ] Is Audit (Standard) actually enabled on the SDG tenant?
- [ ] Does content approval raise a distinct audit operation?
- [ ] Confirm a custom Add+View-only permission level behaves as expected for the log list.

## Not in scope

- Retention policy (backlog item 13) — separate concern, M365 Purview retention labels.
- Any server-side / app-only component. The package stays Graph-free.
- Log archival beyond the list itself.

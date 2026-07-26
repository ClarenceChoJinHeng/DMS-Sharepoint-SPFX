# In-Site Term Store Management — Design

**Date:** 2026-07-26
**Status:** approved decision — verified, not yet applied to production
**Supersedes the assumption in:** the earlier "site-collection–local term store pivot"
(memory `dms-term-store-site-collection-pivot`). Verification changed the recommended path from
"recreate everything locally" to "manage the existing tenant sets in-site."

## Problem

The client refuses to grant **M365 Admin Center** access (Global Admin / SharePoint Admin) on
security grounds, and does not want a security-team ticket every time a DMS metadata label
(Document Type, Department, Unit, etc.) needs adding or editing. The DMS **requires** the term
store — folder routing and Phase 2 cross-site matching key off term GUIDs, and Choice columns are
explicitly disallowed (CLAUDE.md rule #8). So we need a way to create/maintain term sets **without
the Admin Center** and **without a ticket per change**.

## Verified facts (2026-07-26, test site `ClarenceDMSTesting`)

1. The term store is manageable **from the site itself**, not the Admin Center, via either:
   - Classic: `.../sites/ClarenceDMSTesting/_layouts/15/termstoremanager.aspx`
   - Modern: `.../sites/ClarenceDMSTesting/_layouts/15/SiteAdmin.aspx#/termStoreAdminCenter`
     (reached via ⚙️ → Site information → View all site settings → Term store management)
   The URL stays under `/sites/...` — it never touches `admin.microsoft.com`.
2. Clarence could **create a root-level global group** (`DMS Metadata Native`) and a term set
   (`ZZ Test` → term `Hello`), and could **add a term to an existing global set**
   (`ZZ Term` under `Upstream Malaysia Head Office`). Both actions succeeded from the site page.
3. The DMS read path resolves in-site sets: `GET /sites/ClarenceDMSTesting/_api/v2.1/termStore/
   sets/{guid}/children` returned the `Hello` term as JSON — the exact endpoint `Form.tsx` and
   `FolderManager` reconciliation use.

## Decision

**Manage the DMS term sets in the existing global `DMS Metadata` group, in-site, using the term
store management page. No Admin Center. No per-change ticket.** Existing tenant GUIDs in CLAUDE.md
remain valid — **zero GUID churn, no column re-binding, Phase 2 cross-site reuse intact.**

### Access required (least-privilege ladder)

| Tier | Grant needed | Self-service? | Reusable across sites? |
|------|--------------|---------------|------------------------|
| **1 — Primary** | **Contributor** on the `DMS Metadata` global term group (one-time, ~30s, set in the group's General → Contributors/Group Managers) | ✅ forever | ✅ tenant-wide |
| **2 — Fallback** | **None** — Clarence is Site Collection Admin of the DMS site (he creates it), which auto-grants management of that site's **Site level term groups** | ✅ forever | ❌ site-local only (each segment site needs its own set) |
| 3 — Avoid | Hand admin a CSV to import | ❌ ticket per change | — |
| 4 — Rejected | Drop term store, use Choice columns | ✅ | ❌ breaks folder routing + Phase 2 |

- **Contributor** ≠ Admin Center, ≠ Global/SharePoint Admin, ≠ Term Store Administrator. It is
  scoped to a single term group and grants no access to sites, users, or tenant settings.
- **Contributor vs Group Manager:** Contributor edits terms within the group; Group Manager also
  adds other users. Contributor is sufficient for Clarence's work.
- One caveat: someone who already holds term store rights must make the Tier-1 grant once.
  Tier 2 needs no grant at all because Clarence is SCA (he is creating the site).

### Production ask to the client

> "I don't need M365 Admin Center or any admin role. Either (a) grant me **Contributor on the
> `DMS Metadata` term group**, or (b) I'll use the DMS site's own **site-level term group** as
> Site Collection Administrator — no grant needed. Both are managed inside the site; neither
> touches the Admin Center."

## Consequences / follow-ups

- Tier 1 keeps the current architecture unchanged: existing GUIDs, DMS Config `TermSetGuid`
  rows, `DEFAULT_MODES`/`RECON_MODES` fallbacks, and taxonomy column bindings all stay as-is.
- Only if the client forces Tier 2 do the term-store implications from
  `dms-term-store-site-collection-pivot` apply (per-site GUIDs, Phase 2 duplication). Re-run the
  make-or-break read test with a properly restricted account before relying on Tier 2 in prod.
- **Not verified yet:** whether a *restricted* (SCA-only, no term store role) account can manage
  a site-level group end-to-end. Clarence currently has broad rights on this tenant, so today's
  success does not prove Tier 2 under lockdown.

## Cleanup

Delete the throwaway probes: term `ZZ Term` (under Upstream Malaysia Head Office), term set
`ZZ Test`, and group `DMS Metadata Native`.

## Update 2026-07-27 — Tier 2 (site-level) chosen + GUID rewiring

Client is unlikely to grant Contributor on the global group, so we went **Tier 2: site-level term
group** (`DMS`, in `ClarenceDMSTesting`), managed as Site Collection Admin — no grant needed. The
4 shared sets + GHO tree were rebuilt there by CSV import (`docs/term-store-import/*.csv`, pulled
from the old tenant sets via the `/terms` read API). New per-site GUIDs captured below.

**Old (tenant) → new (site-level) GUID mapping — the source-code rewire:**

| Set | Old | New |
|---|---|---|
| Document Type | `0540e66e-7cb3-47ac-b0ef-4e3069387394` | `866c5754-258e-401f-8685-03d20ae59b1d` |
| Year | `f7c578a1-e0e5-42ff-9e0c-d748cba42ede` | `023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf` |
| Confidentiality | `032534ab-9285-4b42-98c6-5c7b0df1f066` | `0d6d1da8-27e5-477f-8684-e8cf169f8fb9` |
| Vendor | `cb3c0ab7-a959-4200-9b7b-d1e13397d240` | `eaafd0e5-03fd-4d33-b1b1-e4252bec430a` |
| GHO segment | `efa87c6a-9536-4f7c-910f-011bf7413b80` | `df4b9afa-d3b9-4c04-9097-50dcaf5d8036` |
| Upstream Malaysia HO | `5ab1c7c4-78d2-43b4-869f-3eab4b1c375c` | `16a52947-57a3-4217-9a49-b48cb8b0dd31` |
| Minamas HO | `6ba9a64c-a363-48fd-afd1-324897df781c` | `c6b26d32-1c3e-441f-b4ea-78f053e12990` |
| NBPOL HO | `21d7e6fe-8f71-4a56-bd2e-e4a2176995a7` | `303f2c38-086a-46ba-8ee0-85445f6bfa3a` |

**Code fallbacks rewired** (`Form.tsx` DEFAULT_MODES/DEFAULT_SETTINGS, `BulkUpload.tsx`,
`FolderManager.tsx` RECON_MODES + YEAR/DOCTYPE consts). `groupMapModel.test.ts` fixtures left as-is
(arbitrary test data). **DMS Config list** (runtime source of truth) must be updated in-UI to the
new GUIDs — code fallbacks only matter offline. **Taxonomy columns** (`Document_x0020_Type`,
`Year_x002f_Period`, `Confidentiality_x0020_Level`, `Vendor`) must be re-bound to the new site-level
sets when the Staging/Documents libraries are built.

## Not in scope

Actually populating the Department/Unit trees for the pending segments (Upstream Malaysia,
Minamas, NBPOL) — that is data entry covered by the segment onboarding plan, done once access is
confirmed.

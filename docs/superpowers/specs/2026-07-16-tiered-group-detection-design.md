# Tiered Group Detection — Design

**Status:** proposed (awaiting review)
**Supersedes** the free segment/level picker in the current `Form.tsx`.

## Problem

The current form exposes a **Business Segment picker** and **editable level dropdowns**. That
lets an uploader tag/route to a segment or unit they don't own — a direct violation of the RBAC
rule: *an uploader may upload only to their own segment → … → unit.* Location must be
**auto-detected from group membership and locked**, never chosen from the full term set.

## Core rule (confirmed)

Upload eligibility is a **hard check across every tier**:

- The user must be a **member** of the group at **each tier** from the top down, **and**
- Hold the **`_UPLOADER` role** group at the **leaf** (deepest) tier.

Being a plain member of a unit (e.g. `DMS_GF_TAX`) grants **view only**. Upload requires the
member groups at every tier **plus** the leaf uploader group (e.g. `DMS_GF_TAX_UPLOADER`). Miss
any tier → that path is not offered, even if the leaf uploader group is present.

Each unit must have its **own distinct** uploader group (own Object ID). Reusing one uploader
group across units would let a member upload everywhere — forbidden.

## The five modes (tier chains; leaf = uploader + permissioned folder)

| Mode | Toggle side | Tier chain (top → leaf) | Leaf |
|---|---|---|---|
| Group Head Office | Business Segment | Segment → Department → **Unit** | Unit |
| Group Upstream Operations | Business Segment | Segment → Region → **Estate/Mill** | Estate/Mill |
| Group SDGI Operations | Business Segment | Segment → Refinery → **Department** | Department |
| Group Innovation & Technology | Business Segment | Segment → **I&T Operating Unit** | I&T Operating Unit |
| Group-led Projects | Project | Project Name → Department → **Unit** | Unit |

Notes:
- For Business Segment modes the **Segment** is the term set itself (not a term); its membership
  is a group tied to the term-set GUID. The chain terms (Department/Region/Refinery/…, then leaf)
  are the term-store hierarchy.
- **Projects has no Segment tier** — its top tier is Project Name (the term set's top-level terms).
- The **leaf differs per mode** (Unit / Estate·Mill / Department / Operating Unit). Generically it
  is simply the deepest level of the mode's `Levels` chain.
- The file always lands in **`…/<leaf folder>/<Year>/<Document Type>/`**. Year and Document Type
  are the only user-selected values (they become on-demand subfolders, unchanged from today).

## Data model — DMS Group Map (generalised)

Today the list holds **unit-only** rows. Generalise it to **one row per (group × tier)**:

| Column | Meaning |
|---|---|
| `GroupId` | Entra Object ID — the match key |
| `GroupName` | cosmetic label |
| `Segment` | the mode's term-set GUID (scopes the row to a mode) |
| `TermGuid` | the term this group represents. **For a Segment-tier row, `TermGuid = the term-set GUID`** (segment membership). For every other tier it is the term's GUID. |
| `Role` | `MEMBER` (membership tiers, incl. Segment and intermediates), `UPL` (leaf uploader), `APR` (approver — used by the approver tooling, not this form), `GLOBAL` (global uploader — bypasses tier detection, may upload anywhere; `Segment`/`TermGuid` ignored) |

> `Level` is intentionally **not** a required column — the term ancestry walk determines each
> term's tier. A `Level` note column may be added for human maintainability but the algorithm
> ignores it.

Example (GHO / Group Legal, Risk & Compliance / GCO):

| GroupName | GroupId | Segment | TermGuid | Role |
|---|---|---|---|---|
| GHO Segment Members | ‹id› | ‹gho-set› | ‹gho-set› | MEMBER |
| LRC Dept Members | ‹id› | ‹gho-set› | ‹lrc-term› | MEMBER |
| GCO Unit Uploaders | ‹id› | ‹gho-set› | ‹gco-term› | UPL |

## Detection algorithm

Inputs: the user's Entra group Object IDs (`/me/memberOf`), all DMS Group Map rows, the modes,
and the site-admin flag.

0. **Privileged short-circuit:** if the user is a site admin **or** belongs to any `GLOBAL`-role
   group → skip steps 1–4, enable the full manual cascade (all modes + all terms).
1. `memberTerms` = { `TermGuid` of every row whose `GroupId` ∈ user's groups } (any role — a UPL
   row also proves membership of its leaf term).
2. `uploaderLeaves` = rows where `Role = UPL` and `GroupId` ∈ user's groups → candidate leaves,
   each carrying its `Segment` (term-set GUID).
3. For each candidate leaf:
   a. Find its mode by `Segment`. Walk the term ancestry in that set: `chain = [top … leaf]`.
   b. **Require every term in `chain` ∈ `memberTerms`** (membership at every tier of the chain).
   c. For **Business Segment** modes, **additionally require the term-set GUID ∈ `memberTerms`**
      (segment membership). Project mode skips this (its top tier is already `chain[0]`).
   d. If all pass → a **valid uploadable path** `{ mode, chain }`.
4. Collect all valid paths. These, and only these, are offerable.

## UI behaviour

- **Toggle (Business Segment | Project):** show a side only if the user has ≥1 valid path on it.
- **Per tier, top → leaf:** if the user's valid paths resolve a tier to **one** term → render it as
  a **locked read-only** value (breadcrumb). If **more than one** → render a dropdown containing
  **only those terms**, and filter the next tier to children of the chosen one.
- The resolved location shows as a breadcrumb, e.g.
  **Group Head Office › Group Legal, Risk & Compliance › Group Compliance (GCO)**.
- **No valid path** (and not an admin) → block upload with:
  *"Your account isn't fully provisioned to upload (you need membership at every level plus the
  unit uploader role). Contact your administrator."*
- **Year / Document Type / Confidentiality / Document Date / Vendor** stay user-editable.
- **Privileged users bypass detection** and get the full manual cascade (all modes + all terms),
  i.e. may upload anywhere. Privileged = a **site admin** *or* a member of a **`GLOBAL`-role**
  group in DMS Group Map. (Uploads are still constrained by SharePoint folder ACLs.) `GLOBAL` is
  purely data-driven — add/remove a list row, no code change.

## Folder routing (unchanged)

The leaf term → DMS Folder Map (`FolderUniqueId`) → resolve current server path → ensure-create
`Year` then `Document Type` subfolders (inherit the leaf's ACL) → upload. Rename-proof, exactly
as built.

## Metadata (unchanged shape)

One label + term-GUID pair per tier in the chain, plus the Business Segment column, via
`buildLevelFormValues` + config-driven `labelCol`/`tidCol`. The leaf and its ancestors are all
written from the resolved `chain` (not from free selection).

## Code impact

- `formModel.ts`: replace `matchUserPaths` (unit-only) with a tiered resolver returning valid
  `{ mode, chain }` paths; add unit tests (member-at-every-tier, missing-tier rejected,
  leaf-needs-UPL, multi-path choice, Project has no segment tier).
- `Form.tsx`: remove the segment picker; drive the cascade from resolved paths (locked vs
  choose-among-own); render the breadcrumb; keep admin full-cascade; block when no path.
- `dmsFolderMap.ts`: unchanged.
- `DMS Group Map` list + seed-data runbook: add Segment/intermediate rows; `TermGuid` replaces
  `UnitTermGuid`; `Role` gains `MEMBER`.

## Setup impact (tenant)

Groups and DMS Group Map rows are now required at **every tier** (segment, each intermediate,
each unit uploader) — materially more than the 6 unit rows in the sandbox today. The seed-data
runbook's DMS Group Map section will be expanded accordingly.

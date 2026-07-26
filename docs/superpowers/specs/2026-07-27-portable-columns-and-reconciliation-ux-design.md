# Portable Column Names + Reconciliation Progress/Throttle — Design

**Date:** 2026-07-27
**Status:** IMPLEMENTED 2026-07-27 (tsc green; 79 unit tests pass). Throttle params are code
constants (`RECON_WRITE_DELAY_MS`/`RECON_BATCH_SIZE`/`RECON_COOLDOWN_MS`) rather than DMS Config
rows — deferred as a follow-up; cap is an invisible auto-cooldown (no user click) per the client
UX request.
**Relates to:** the in-site term-store / portability work
(`2026-07-26-in-site-term-store-management-design.md`) and the reconciliation provisioner
(`2026-07-20-reconciliation-provisioner-design.md`).

Two independent enhancements agreed this session. Build **Feature 1 first**, then Feature 2.

---

## Feature 1 — Config-driven metadata column names

**Why:** the taxonomy/date + BusinessSegment column internal names are hardcoded in `FIELDS` /
`LEVEL_COLUMNS` (Form.tsx + BulkUpload.tsx). Moving the DMS to another site currently means editing
code to match that site's column internal names. Make them a config "map" so a new site needs
**zero code changes** — set the names in DMS Config instead. (Level columns Department/Unit are
already config-driven via DMS Config Levels JSON `labelCol`/`tidCol`; the site URL is already
dynamic via `context.pageContext.web.absoluteUrl` — no work needed there.)

**Design:**
- Extend `DmsSettings` with a `columns` block:
  `{ documentType, yearPeriod, documentDate, confidentiality, vendor, businessSegmentLabel, businessSegmentTid }`.
- Settings loader reads new DMS Config `setting` rows (Title → SettingValue) with fallback to the
  hardcoded defaults:
  `col_documentType, col_yearPeriod, col_documentDate, col_confidentiality, col_vendor,
   col_businessSegment, col_businessSegmentTid`.
- Replace `FIELDS.*` and the `LEVEL_COLUMNS.BusinessSegment` usages in Form.tsx + BulkUpload.tsx
  with the loaded values.
- **Defaults = the current ClarenceDMSTesting columns** so the site works with NO config rows:
  `Document_x0020_Type`, **`Year`** (client chose plain `Year`, not `Year_x002f_Period`),
  `DocumentDate`, `Confidentiality_x0020_Level`, `Vendor`, `BusinessSegment`,
  `BusinessSegmentTid`. Department/Unit tids stay via Levels JSON (their columns:
  `Department_x0020_Tid` / `Unit_x0020_Tid` on this site — set via Levels JSON, not code).
- ApprovalDocument.tsx `pick()` reads (display) — update the `Year` entry to plain `Year` too.

**Out of scope:** making the Levels JSON columns configurable (already are).

---

## Feature 2 — Reconciliation live progress + throttle-safe auto-continue

**Problems (verified in `runReconciliation`, FolderManager.tsx):**
1. `setLog(entries)` is called only at the END — no live feedback, so the run looks like it
   "snaps back to the button." Confusing.
2. The Year×DocType grid is created in **parallel** (`mapLimit(docTypeLabels, GRID_CONCURRENCY,…)`)
   → hundreds of near-simultaneous folder POSTs → SharePoint **HTTP 429 throttling**.

**Design:**

*Throttle safety (layered):*
- **429-aware writes:** wrap all SP write calls (createFolder, breakInheritance, addRoleAssignment,
  ensureFolder) so a `429` reads `Retry-After`, waits, and retries with exponential backoff (cap
  retries ~5). This is the robust fix.
- **Serialize the grid:** replace the parallel `mapLimit` with sequential creation.
- **Inter-write delay:** `await sleep(delayMs)` between folder ops. Default **500ms**.
- **Invisible cooldown batching:** every **N=150** folder ops, auto-pause `cooldownMs` (default
  ~4000ms) then continue — **fully automatic, no user click** (client-friendly: a click would
  annoy them). During the cooldown the spinner keeps running and the status text rotates
  (e.g. "Discombobulating…" → "Generating folders…") so the pause reads as work.
- Config-driven: DMS Config `setting` rows `recon_writeDelayMs`, `recon_batchSize`,
  `recon_cooldownMs` (fallback to 500 / 150 / 4000).

*Live two-panel progress (replaces silent run):*
- New `reconProgress` state, updated live as the loop runs (React repaints between the awaited
  delays). Shape: `{ folders: ProgressItem[], assignments: ProgressItem[], phase, counts }`.
- **Left panel "Folders":** each structural folder (segment → dept → unit) as its own line with a
  status: spinner while creating → `✓ created+locked` / `skipped (exists)`. The Year×DocType grid
  is a **single updating counter per unit** (`<unit> grid: 37 / 60`), not one line per folder.
- **Right panel "Group assignments":** each folder → group with `✓ assigned` / `✗ failed`, and when
  the SP group is missing → **"Group '<name>' not found — ask an administrator to create it"**
  (the existing admin-needs-to-create message).
- Header: running counts (folders done / total-ish, assignments) + rotating status text.

**Notes / gotchas:**
- Keep the run **idempotent** (already is) so the batched cooldown/continue is safe.
- Preserve the existing isolation rule (Staging = UPL/APR only; Documents = MEMBER only) and the
  ancestor-Read browse grants — only the *sequencing* and *reporting* change, not the ACL logic.
- `ensureFolder` is in `shared/dmsFolderMap.ts` — the 429 wrapper must cover it too (either wrap at
  the fetch level there or pass a retrying fetch).

**Verification:** `tsc --noEmit` green; a reconciliation run shows live two-panel progress, never
snaps back silently, survives a forced 429 (backoff), and auto-continues past 150 without a click.

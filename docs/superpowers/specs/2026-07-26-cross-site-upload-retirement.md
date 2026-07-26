# Cross-Site Upload / Multi-Site Storage — Retirement

**Date:** 2026-07-26
**Status:** approved — implementing now
**Retires:** `2026-07-20-multi-site-storage-phase2-design.md` (Phase 2 multi-site model)

## Why

The multi-site design existed to spread DMS storage across separate **site collections** (one per
segment), on the assumption that this distributed or saved storage. That assumption is wrong:

> A new SharePoint site collection draws from the **same tenant storage quota**. Creating more
> sites does not add or save space — it only adds ACL, term-store, and routing complexity across
> collections.

With no storage benefit, the cross-site model is pure overhead. The DMS stays **single-site**:
one site collection, unit folders + auto-route to the in-site Documents library (unchanged).

## Scope of removal

**Delete:**

| Path | Kind |
|---|---|
| `src/webparts/crossSiteBrowser/` | web part — cross-site library browser |
| `src/shared/siteMap.ts` | cross-site mapping helpers (imported ONLY by crossSiteBrowser) |
| `src/shared/siteMap.test.ts` | its unit tests |

**Edit:**

| File | Change |
|---|---|
| `config/config.json` | drop `cross-site-browser-web-part` bundle + `CrossSiteBrowserWebPartStrings` |
| `config/package-solution.json` | remove component id `b2d5f8a1-3c7e-4d90-a1f2-6e8c9b04d7a3`; bump solution version |
| `docs/superpowers/specs/2026-07-20-multi-site-storage-phase2-design.md` | mark superseded by this retirement |

**Docs (untracked):** `docs/cross-site-upload-flow.drawio` (+ `.pdf`, `.bkp`) — the client may
delete these; they are not in git.

## NOT in scope (stays)

- The **in-site auto-route** flow (approved files → Documents library in the *same* site). This is
  not cross-site; it is core to the pipeline. Unchanged.
- Unit-folder ACL model, DMS Group Map, upload/approval pipeline.
- Term store decision (`2026-07-26-in-site-term-store-management-design.md`). Note: dropping
  multi-site also removes the cross-site term-GUID-reuse concern that the term-store pivot worried
  about — with one site there is nothing to reuse across.

## Verification

- `npx tsc --noEmit` succeeds with no unresolved imports from the deleted modules.
- `grep -rn "siteMap\|CrossSiteBrowser" src/` returns nothing.
- No remaining `cross-site` references in `config/`.

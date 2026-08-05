# Configurable Name Prefix (DMS → CRS) — Design

**Date:** 2026-08-04
**Status:** DRAFT — agreed in principle 2026-08-04 ("everything, config-driven"). Not implemented.
**Branch:** `feat/folder-abbreviations`
**Related:** memory `dms-to-crs-rename-pending`, `2026-07-28-client-site-migration-runbook.md`,
`2026-07-27-site-entry-access-layer-design.md`

---

## 1. The problem, stated precisely

The client rebrands every `DMS`-named artefact to `CRS` when the solution is imported to their
site: four lists, a content type, a term-store group, and every SharePoint group. The code
hardcodes `DMS` in **105 places across 13 files**.

Two things make this worse than a find-and-replace:

**A missed reference is a runtime failure, not a compile error.** `getbytitle('DMS Config')`
compiles perfectly and returns HTTP 404 on the client's site. It surfaces during their UAT, not
in our build.

**The current failures are silent.** Two are already known:

- `spGroupsFilter.ts` keeps only groups whose title starts with `DMS_`. On a `CRS_` site it
  matches nothing, so the site-entry membership sync **adds nobody** and reports success.
- `SITE_ENTRY_GROUP_NAME` is the literal `"DMS_SITE_MEMBERS"`. On a `CRS_` site the site-entry
  pass creates a *second*, wrongly-named group and grants it site Read, while the group the
  client actually uses gets nothing.

Neither raises an error. The symptom in both cases is "users cannot open the site", which reads
as a permissions problem and sends the investigation to the wrong place.

Hardcoding `CRS` instead of `DMS` fixes today and rebuilds the same trap for the next rename.

## 2. What is named

| Kind | Today | Derived as |
| --- | --- | --- |
| Config list | `DMS Config` | `<P> Config` |
| Group map list | `DMS Group Map` | `<P> Group Map` |
| Folder map list | `DMS Folder Map` | `<P> Folder Map` |
| Abbreviation list | `DMS Term Abbreviation` | `<P> Term Abbreviation` |
| Deletion log list | *(new)* | `<P> Deletion Log` |
| Folder content type | `DMS Folder` | `<P> Folder` |
| Term-store group | `DMS Metadata` | `<P> Metadata` |
| SharePoint group prefix | `DMS_` | `<P>_` |
| Site-entry group | `DMS_SITE_MEMBERS` | `<P>_SITE_MEMBERS` |

`<P>` is the prefix: `DMS` or `CRS`, and whatever a future client chooses.

## 3. Resolution: probe per artefact, not once globally

### 3.1 Why the prefix cannot simply be read from config

The obvious design — a `namePrefix` row in the config list — cannot bootstrap itself: the config
list's own name depends on the prefix. Reading `DMS Config` to discover that the prefix is `CRS`
is circular.

So the prefix is **discovered**, not configured: probe the candidate prefixes in order against
`<P> Config` and take the first that responds.

```
CANDIDATE_PREFIXES = ["CRS", "DMS"]
```

`CRS` first, so a fully renamed site costs one request and the legacy site costs two. Extending
that list is how a third client gets onboarded.

### 3.2 Resolution is PER ARTEFACT, not one global answer

The tempting shortcut is to resolve once and derive every other name from that prefix. It is
wrong, for a reason specific to how this rename actually happens: **the client renames lists by
hand, one at a time, through the SharePoint UI.** There is necessarily a period — minutes, or a
coffee break, or a week if they are interrupted — in which `CRS Config` exists and
`CRS Group Map` does not.

A single global prefix turns that window into a total outage: the moment `Config` is renamed,
every other list is looked up under `CRS` and 404s. Per-artefact resolution makes a partial
rename merely partial.

So each name resolves independently, caching its answer:

```
resolveListTitle("Group Map")  →  probes "CRS Group Map", then "DMS Group Map"
```

**A consequence worth stating plainly:** the rename and the deployment become order-independent.
The client can rename before, after, or halfway through, and nothing needs coordinating. That is
the actual deliverable — not the string change.

### 3.3 Caching, and when it is wrong

Resolved names are cached for the page session. A rename performed *while* an admin has the
Folder Manager open is not picked up until refresh. Acceptable, and far cheaper than probing on
every call — but it means a reconciliation run must print the names it resolved, so "why is it
still writing to the old list" is answerable without a debugger.

### 3.4 The one thing that must never be probe-dependent

The **group prefix filter** accepts *every* candidate prefix, always — not just the resolved one.

`spGroupsFilter` decides which SharePoint groups are "ours", for search and for the site-entry
membership sync. If it accepted only the resolved prefix, a half-renamed site (some groups
`CRS_`, some still `DMS_`) would silently sync one half and ignore the other — producing exactly
the "only some users cannot open the site" report that is hardest to diagnose. Matching both
costs nothing: a group named for a different client's prefix will not exist on this site.

## 4. Where names come from after this

One module, `src/shared/naming.ts`:

- `resolveListTitle(suffix)` — `"Group Map"` → the live title
- `resolveContentTypeName()` — the `<P> Folder` content type
- `groupPrefix()` — the resolved `<P>_`, for *writing* new group names
- `matchesAnyGroupPrefix(title)` — for *reading*; accepts all candidates
- `siteEntryGroupName()` — `<P>_SITE_MEMBERS`

**No list-title literal may appear anywhere else.** That is the invariant the whole design rests
on, and it is worth enforcing mechanically rather than by review — see §5.

## 5. A test that reads the source

Because a missed reference cannot fail the build, one test scans `src/**/*.{ts,tsx}` for the
patterns `'DMS ` / `"DMS ` / `DMS_` and fails on any hit outside `naming.ts` and test fixtures.

This is unusual and deliberate. Every other guard in this codebase catches a wrong *value*; this
one catches a wrong *place*, which is the only failure mode that matters here. Without it the
refactor is "we think we got all 105", and the way we would find out is the client's UAT.

Allowed exceptions, listed explicitly in the test so adding one is a visible decision:
`naming.ts` (the candidate list), `*.test.ts` fixtures, and `Form-Copy.txt` /
`Form.reference.tsx` (not compiled).

## 6. Migration

Nothing to migrate. Both prefixes resolve, so:

- **Our test site** keeps its `DMS` names and carries on untouched.
- **The client's site** works whether they have renamed nothing, everything, or half.
- Neither needs a deployment window.

The `namePrefix` config row is deliberately **not** introduced. It would be a second source of
truth alongside the actual list names, and the two would disagree the moment someone renamed a
list without editing the row — restoring the silent failure this spec exists to remove. The live
names are the only truth.

## 7. Out of scope

- **Renaming anything on the client's behalf.** They rename; we adapt.
- **The `Staging` and `Documents` library names.** Already resolved from config, and not prefixed.
- **Term-set GUIDs.** Per-site already, and unrelated to naming.

## 8. Implementation order

1. `naming.ts` + unit tests — pure probe-order logic with an injectable fetch, so no tenant is
   needed.
2. The source-scanning test from §5, landed `.skip`ped, so it arrives as the definition of done
   rather than as a hundred red assertions.
3. Convert the shared modules: `dmsFolderMap.ts`, `allowedFileTypes.ts`, `formModel.ts`,
   `groupExportCsv.ts`, `groupMapModel.ts`, `spGroupsFilter.ts`.
4. Convert the web parts: `Form.tsx`, `BulkUpload.tsx`, `FolderManager.tsx`,
   `GroupMapBuilder.tsx`, `FolderMap.tsx`, `ApprovalDocument.tsx`. Un-skip the §5 test.
5. Log the resolved names as the opening lines of a reconciliation run (§3.3).

Steps 1–2 are safe to land alone. Step 3 onward changes live list reads, so each web part is
verified on the test site before the next is converted — a 404 from a mistyped suffix looks
identical to a missing list, and converting everything at once makes it unclear which change
caused it.

## 9. Verification

1. **Unrenamed site (ours).** Every web part loads, reconciliation runs, no behaviour change.
2. **Fully renamed site.** Rename all four lists, the content type and the term group to `CRS`,
   plus one group. Confirm every web part loads and reconciliation resolves `CRS` names.
3. **Half-renamed site — the case per-artefact resolution exists for.** Rename `Config` only.
   Confirm the other three lists still resolve to `DMS` and nothing 404s.
4. **Mixed group prefixes.** One `CRS_` group and one `DMS_` group, each with a member.
   Reconcile and confirm **both** members land in the site-entry group.
5. **Site-entry group name follows the prefix.** On a `CRS` site, confirm the pass finds or
   creates `CRS_SITE_MEMBERS` and does **not** create `DMS_SITE_MEMBERS` beside it.
6. **The §5 test fails when it should.** Reintroduce one `'DMS Config'` literal and confirm the
   suite goes red. A guard never seen to fail is not known to work.
7. **Resolved names appear in the log.** Confirm the run's opening lines name the lists resolved.

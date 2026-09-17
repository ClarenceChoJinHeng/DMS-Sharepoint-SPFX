# Adding a new folder level no longer migrates anything — old and new shapes coexist forever

**No file calls this document** — a design spec under `docs/superpowers/specs/`, read before code,
per this project's standing rule. **No data file is read or written by this document.** Confirmed via
`Glob`: `*structure*` matches `2026-08-06-configurable-folder-structure-chain-design.md` and
`2026-08-10-structure-manager-ui-design.md` — the two specs that established the CURRENT
migrate-everything model this document partially reverses, not duplicates of the new idea;
`2026-08-11-subtree-migration-design.md` is the migrator's own spec, referenced below rather than
repeated.

**Client, 2026-09-17, on the old "Change the folder structure" flow moving 200-300 files at once:**
*"client do not want to move the files one by one which is crazy, it make sense. So now the flow is
this… if 2026 is already there and admin adds a new layer of new folder, then let the old 2026 stays
there but the next new folder will continue to have 2026 underneath."*

**Status: DECLINED, 2026-09-17, same day it was written. Kept for the record, not deleted, so it is
not re-proposed later as though it had not been considered.**

⚠ **Client, 2026-09-17, reversing this within hours of it being confirmed:** *"I just got an update
about the folder structure design change, I think since we do not have time we have to considering
abandoning the new idea for now. What we can do instead is auto move the files for them, and they
decide to move the files that are not suppose to be there manually."*

**What replaces it is not a new build at all.** Re-reading the client's ORIGINAL complaint that
started this whole document — *"if there is atleast 200 to 300 files, client do not want to move the
files one by one which is crazy"* — against the tool as it already exists today: the current
`SubtreeMigrator` does not require moving files one by one. The admin picks ONE value for the new
tier (per library, since 2026-09-09), and the migrator **auto-moves every existing file in bulk**
into that chosen value's folder in a single run. The tedium the client was reacting to is not "the
initial move," it is the **post-move correction** — if some of those bulk-moved files actually belong
under a *different* value for the new tier, fixing those specific ones up is a manual step today, in
SharePoint, one folder at a time.

**The client has now accepted that manual correction step**, rather than having the app build the
per-upload "skip this tier / use the old path forever" mechanism this document specified. So: **no
code changes needed for this item.** The existing, already-built "Change the folder structure" flow
is what continues to be used, unmodified, for adding a new level — bulk auto-move to one chosen
value, any exceptions relocated by hand afterward.

Everything below this line is the abandoned design, kept as the record of what was considered.

---

## 1. Confirmed shape

- **Nothing ever moves.** Adding a level never migrates a single existing document. The old shape
  stays exactly where it is, forever, fully browsable and searchable.
- **The new tier gets its own fresh branch, not a wrapper around the old one.** Confirmed against the
  screenshots: `2024` sits as a sibling of the newly added `New folder level`, which then grows its
  own `2025` underneath — the new level is inserted as a **parallel** path, not a container the old
  folder gets moved inside.
- **The new tier is optional, per upload, forever** (client confirmed option **B** over A):
  every future upload gets an active choice — pick a value for the new tier (nested, new path) or
  explicitly skip it (flat, old path) — and this stays true indefinitely, not just during a
  transition window.

This is a genuinely different model from today's: currently a structure change stages to
`PendingLevels`, a migration scan runs, files move, and once clean the new chain goes fully live —
one shape, applied uniformly. Here there is **no migration step at all**, because there is nothing to
migrate; the new tier is additive, and the old shape keeps being a valid, ongoing destination.

## 2. Scenarios — the three insertion points, brainstormed as requested

In every case below, only the **one newly added tier** becomes optional. Every tier that already
existed in the chain keeps behaving exactly as it always has — mandatory if it always was, in the
same relative order to every OTHER unchanged tier.

### 2.1 Before Year (the shown example)

```
Old (unmoved, forever):  Unit / 2024 / DocType / file.pdf
New tier "TestingLayer" inserted directly under Unit, ahead of Year.

Upload picks a TestingLayer value:  Unit / <TestingLayerValue> / 2025 / DocType / file.pdf
Upload skips TestingLayer:          Unit / 2025 / DocType / file.pdf   (same flat shape as always)
```

Skipping continues writing into the SAME flat branch old files already live in — a 2025 file that
skips the new tier lands right beside the 2024 files, in the identical shape.

### 2.2 Between Year and Document Type

```
Old (unmoved, forever):  Unit / 2024 / DocType / file.pdf
New tier inserted between Year and DocType.

Upload picks a value:  Unit / 2025 / <TestingLayerValue> / DocType / file.pdf
Upload skips it:       Unit / 2025 / DocType / file.pdf
```

Year is unaffected either way — it is not the tier that became optional, so a 2025 upload always
gets a real `2025` folder regardless of which branch it ends up in beneath that.

### 2.3 After Document Type

```
Old (unmoved, forever):  Unit / 2024 / DocType / file.pdf
New tier inserted after DocType.

Upload picks a value:  Unit / 2025 / DocType / <TestingLayerValue> / file.pdf
Upload skips it:       Unit / 2025 / DocType / file.pdf
```

Symmetric with 2.2, just at the tail.

**The pattern is identical across all three**: exactly one tier in the chain gains a genuine
either/or at build time — include this segment in the path, or omit it and continue straight to
whatever tier comes next, at the depth it would have occupied before the new tier existed. Everything
else in the chain (before and after the new tier) is untouched by this change.

## 3. What this means for the upload form

The tier that was just added needs a distinction the cascade does not currently have:

- **Today's "optional" is a per-unit COMPUTED fact** — a below-Unit tier with no term children for
  THIS unit is skipped automatically (`decideTier` → `skip`), with no choice offered because there is
  genuinely nothing to choose from.
- **This is a per-upload CHOICE** — real options exist, and the uploader may still decline them. The
  dropdown needs an explicit "skip / file under the existing structure" option alongside the real
  term values, distinct from today's silent-skip-when-empty behaviour.

This is a new flag on the level (working name `coexisting?: boolean`, separate from `abbreviated` and
`permissioned`), read by the SAME cascade logic (`buildOnDemandSegments` in `folderChain.ts`) that
already decides tier-by-tier whether to render a dropdown.

## 4. What this means for the structure-change flow itself

⚠ **Proposed simplification, not yet confirmed with the client**: since nothing is being migrated
and there is no "two shapes exist at once, and that's a hazard" window (coexistence IS the intended
end state, not a transitional risk `PendingLevels`/staging exists to guard against), adding a
`coexisting: true` level may not need the pause-uploads → migrate → reconcile → resume machinery at
all. It could write straight to the live `Levels` chain and take effect immediately on Save — a much
lighter admin flow than today's structure change.

This is flagged as a question rather than assumed, because the client has previously preferred
consistency over efficiency in an adjacent decision (the "Rename or re-code a folder" flow gaining a
hard upload pause purely "so we don't confuse the client, they have to go back and forth" — even
though a plain rename needs no pause at all). The same instinct might apply here.

## 5. Reorder and remove are UNCHANGED, and stay on the existing migrator

Restated plainly, since the first phrasing of this question didn't land: **reordering** an existing
tier (swapping the order of two tiers that already have real values) and **removing** a tier
(collapsing sibling folders together) both deal with documents that already HAVE a real value for
that tier — there is no "old files have no value for a level that didn't exist yet" case for either
operation, so the whole premise behind "let old stay, diverge going forward" doesn't apply. My
reading is that reorder and remove keep using the existing `SubtreeMigrator` exactly as it does
today (genuinely moving/collapsing files) — **only ADDING a brand-new level adopts this new,
non-migrating, coexisting behaviour.** Please confirm or correct.

## 6. Metadata, search, approval screens — confirmed unaffected

Client, confirmed: *"if that file is following the folder path, then the metadata folder path also
follows"* — matching how this already works today. A document's tier values are written to its own
`<Base>Tid` columns at upload time, independent of folder depth; every downstream screen (approval
details panel, My Submissions, CRS Search) reads those columns, not the folder's position in the
tree. A document sitting in the old flat shape and one sitting in the new nested shape are read back
identically by everything that isn't the folder browser itself. No change needed here.

## 7. Open questions

1. **Is `coexisting: true` a per-level CHOICE the admin makes when adding a level** (a toggle,
   leaving the option to still request the old migrate-everything behaviour for some future case),
   **or does "add a level" simply stop migrating, full stop, for every future addition?** The
   screenshots and the client's own framing ("let the old 2026 stays there") read as the latter, but
   worth pinning down explicitly since it decides whether a toggle needs building at all.
2. **Does §4's simplification (skip the whole pause/migrate/reconcile flow for this operation) match
   what the client wants**, or should adding a coexisting level still go through the same staged
   ceremony as every other structure change, purely for consistency with the rest of that flow?
3. **Naming for the "skip this tier" option** in the upload cascade — something that reads clearly to
   an uploader who has no idea a structure change ever happened (e.g. "Not applicable" vs "Skip" vs
   leaving it blank with a hint).

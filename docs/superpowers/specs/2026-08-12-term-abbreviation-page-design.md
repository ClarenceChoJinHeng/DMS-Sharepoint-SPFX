# Term Abbreviation page, and one home for folder administration

**Date:** 2026-08-12
**Status:** BUILT 2026-08-12 — not yet site-tested

Builds the admin UI for the data specced in
[2026-07-30-folder-abbreviation-naming-design.md](2026-07-30-folder-abbreviation-naming-design.md),
which defined the naming scheme but left the list to be edited by hand.

---

## 1. Why

Folder names come from `DMS Term Abbreviation`, keyed by term GUID. A term with **no abbreviation is
skipped** by every reconciliation run — no folder, and that unit cannot upload. So the list is
load-bearing, and today it is maintained either in the SharePoint list view or by pasting a console
script. The client does not want to touch list pages, and after handover there is nobody to run a
script for them.

It is also the worst kind of data to hand-edit: keyed by **term GUID**, a value the client has no way
to obtain except from a term-store properties panel, one term at a time.

## 2. The page never shows a GUID

An admin picks a **segment** from a dropdown. The page walks that segment's term tree and lists every
term, indented by level, with its current abbreviation in an editable box. `TermGuid` is resolved
behind the scenes and never displayed or typed.

This replaces the client's own first suggestion — paste a parent term's GUID, list its children —
because the segment is already known from its `DMS Config` mode row, and showing the whole tree at once
is what makes the sibling check in §4 possible without a single request.

## 3. Compulsory stays compulsory — "Same as term name" is the escape hatch

The client asked whether a term set could be **exempt** from needing abbreviations (their example:
Year — 2024, 2025). Two answers:

- **Below-Unit tiers already need none.** Year, Document Type and SubUnit name their folders from the
  sanitized term label (`buildOnDemandSegments`, `folderChain.ts`). Nothing to build.
- For a **permissioned** tier, an exemption was designed as a per-level flag and then **rejected in
  favour of the client's own better idea**: a **"Same as term name"** button that copies the term label
  into the abbreviation box.

The button wins because it adds no second naming path. An exemption flag would give reconciliation two
ways to name a folder, one of them invisible in the abbreviation list — so the list would stop telling
you what a folder is called. With the button, the stored value is always the literal folder name, the
"no abbreviation, no folder" guard stays exactly as it is, and the sibling-collision check and
rename-on-change behaviour apply to it for free.

It could not have been keyed on the term set anyway: a segment's permissioned tiers (Region,
Estate/Mill) live in **one** term set, so a per-set flag cannot distinguish them.

Offered per row, and per level in bulk — one click for a whole tier of short names like Regions.

**No auto-suggested initials.** A guessed `GMB_PMB2C` looks authoritative and is wrong; the client's
codes are real data, not derivable. An empty box is honest, and "Same as term name" already removes the
tedium wherever the label is short enough to use.

## 4. The live sibling-collision check

Two siblings sharing an abbreviation is the worst outcome the naming scheme can produce, and nothing
about the resulting folder tree looks wrong:

```
Compliance & Operational Risk (CORU)  ->  CORU
Corporate Reporting                   ->  CORU     <- same parent
```

Both resolve to `/GHO/GF/CORU`. The second does not create a second folder — it lands in the same one.
Two units then share one folder and **one ACL**, so each reads the other's documents.

Reconciliation already detects this and aborts before creating anything. The page moves that failure to
typing time, where it is cheap: the person who runs reconciliation is usually not the person who typed
the code.

- **Both rows are flagged**, not just the second. Which one is wrong is not knowable.
- Each message **names the other term and their shared parent**.
- **Save is blocked** while any collision exists. Nothing is lost — the fix is on the same screen — and
  the page must never save something reconciliation would refuse.
- **Compared case-insensitively and after sanitizing**, because that is what the folder will be called:
  SharePoint folder names are case-insensitive, so `coru` and `CORU` are one folder, and `GC EP` /
  `GC  EP` collapse to the same segment.
- **Siblings only.** `Tax` under Group Finance and `Tax` under Minamas GA both stay `TAX`. Not an edge
  case — `Tax`, `Legal` and `PM` each repeat under several parents in the client's data.

## 5. The rest of the page's behaviour

- **Rows with no abbreviation are flagged loudly** — *"no folder will be created"* — because that is the
  state that silently blocks a unit, and it is invisible in the list view today.
- **A length warning** on the resulting folder name. "Same as term name" makes it one click to create a
  60-character folder segment, and long paths are exactly what the scheme exists to avoid: a deep path
  returned HTTP 400 at roughly 330 characters (gotcha #9).
- **Below-Unit tiers are listed but read-only**, marked *"uses the term name — no abbreviation
  needed"*. Omitting them would send an admin hunting for Year.
- **Changing an existing abbreviation warns that it renames a live folder** on the next reconciliation
  run. That rename is safe and already built — UniqueId, contents, ACL and approval status all survive,
  in both libraries.
- **Save writes rows ONLY.** No folder is created or renamed by saving; folders change only when someone
  runs Folder Reconciliation. Saving is therefore inert and repeatable, and the destructive step stays
  behind its own button. The order is always: fill in → save → reconcile.
- **Unsaved-changes guard**, as on the other structure screens.

## 6. One home for folder administration

Folder administration is currently spread over two web parts and a list. It becomes one tab bar:

| # | Tab | Source |
|---|---|---|
| 1 | Term Abbreviations | new (this spec) |
| 2 | Folder levels | Folder Structure web part |
| 3 | Move existing folders | Folder Structure web part |
| 4 | Folder Reconciliation | stays |
| 5 | New segment | Folder Structure web part |

The order is the order the work happens in: name the terms, shape the levels, move what is already
filed, reconcile. A new segment sits last because it is the rarest.

**Term Abbreviations opens first**, which changes the landing tab from the old `Staging` folder tree.
Reconciliation is the tab an admin visits most, but it is the one that *fails* when the first tab has
not been filled in — a term with no code gets no folder — so it is not the place to start.

The three moved tabs are **MOUNTED from `userAccess/components`, not copied**. `StructureManager`
rewrites `Levels` and creates columns in both libraries; `SegmentCreator` writes a `mode` row. A second
copy of either would drift, and the symptom of drift in those two is a wrong column name or a
half-written segment, both of which surface weeks later.

**`Staging` and `Documents` are removed** (client, 2026-08-12: *"I am honestly not using it"*). They
were a manual folder tree — rename, create and assign permissions per folder in one commit — which
reconciliation and the Folder Access page now cover from data. What goes with them: browsing a folder to
see who actually holds access to it, and hand-creating a folder. The second is arguably worth losing,
since a hand-made folder is one reconciliation does not know about.

**The `Folder Structure` web part stays registered** even though its tabs now appear here. It may
already be on a page, and an unregistered web part leaves a broken zone. Deleting it is a one-line
change once no page uses it.

The web part is **retitled `Folder Administration`** (manifest title, property-pane label and the
page heading). The component **id is unchanged**, so any page already hosting it keeps working — only
the name in the web part picker moves. "Folder Manager" described the folder tree that is now gone.

**The retired folder tree's CODE stays, unreachable.** Its render branch, its Refresh/Update bar and
their helpers still compile; deleting ~500 lines belongs in its own reviewable change, not buried in a
tab restructure. Two consequences were handled explicitly:

- The tree's controls were gated on `tab !== "Reconciliation"`, which after this change also means
  Abbreviations, Levels, Migrate and New segment — it would have rendered a second Refresh/Update pair
  over every mounted screen, each of which has its own Save. The test is now `treeTab`, named for what
  it actually asks.
- The mount-time tree crawl is **gated off**. It was dozens of requests down every branch of a library
  on every page load, for a view nobody can open, and its failure toast told the admin to check
  permissions on a tab that no longer exists.

## 7. Testing

- A term with no abbreviation: flagged on the page; reconciliation skips it and reports it.
- Two siblings given the same code: both rows flagged, save blocked, message names the other term.
- The same code under two **different** parents: allowed, saves, two distinct folders.
- `coru` against an existing `CORU`: caught (case-insensitive).
- "Same as term name" on a short label (`Perak`): folder is `Perak`, no warning.
- "Same as term name" on a 60-character label: allowed, with the length warning.
- Change an existing abbreviation, save, reconcile: the live folder is renamed in **both** libraries,
  documents still present, permissions unchanged, no re-approval needed.
- Save, then check the folder tree: **nothing changed** until reconciliation runs.
- Below-Unit tiers are visible and not editable.
- Folder Administration: five tabs, Staging and Documents gone, each moved tab behaving as it did.
- Type into a tab, then click another tab: the switch is refused, not confirmed. Save, and the refusal
  message clears itself.
- The page no longer crawls the library on load — the retired folder tree's read is gated off, so an
  admin with no library-root rights sees no "check your permissions" toast on arrival.

## 8. Related

- [2026-07-30-folder-abbreviation-naming-design.md](2026-07-30-folder-abbreviation-naming-design.md) —
  the naming scheme this edits: skip-not-guess, sibling uniqueness, rename on change
- [2026-08-02-term-guid-orphan-repair-design.md](2026-08-02-term-guid-orphan-repair-design.md) — why
  abbreviation rows are never auto-deleted, and how a dead GUID is repaired by level + label
- [2026-08-12-add-segment-design.md](2026-08-12-add-segment-design.md) — slice B, whose closing
  checklist points at this page

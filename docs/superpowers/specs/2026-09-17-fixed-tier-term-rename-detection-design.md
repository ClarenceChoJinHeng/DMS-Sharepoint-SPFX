# A term-store LABEL rename on a fixed below-Unit tier (Document Type, Year) currently strays the path

**No file calls this document** — it is a design spec under `docs/superpowers/specs/`, read before
implementation, per this project's own standing rule to write specs before code. **No data file is
read or written by this document itself.** Confirmed via `Glob` that no existing spec covers this
exact gap: `*rename*` matches `2026-07-12-rename-proof-upload-design.md` (upload-time filename
collisions, unrelated), `2026-08-28-library-rename-flow-repair-runbook.md` (a whole-library rename,
unrelated) and `2026-09-15-segment-rename-and-group-rename-design.md` (segment/group renames,
unrelated); `*term*` matches `2026-08-17-term-manager-design.md`, which is the closest relative and
is discussed and corrected in §1 below, and `2026-09-09-below-unit-abbreviations-design.md`, which
this spec is the direct complement of (that one covers CODED below-Unit terms; this one covers the
UNCODED, fixed ones it deliberately excludes).

**Client, 2026-09-17:** *"I basically went directly into the term store and rename Power point slide
deck into slide deck and reupload a file again and it created another folder Slide Deck instead of
renaming the current existing doctype power point slide deck folder into slide deck."*

Clarified the same session: this is about renaming a **term value** (a child term under the Document
Type term set, e.g. "PowerPoint Slide Deck" → "Slide Deck"), never about renaming the tier/column
name "Document Type" itself. *"Oh I think I might have reword this wrongly… I did not meant changing
the wording Document Type in term store to something else."*

**Status: NOT BUILT. Design only.**

---

## 1. Corrects a claim in the existing, unbuilt `2026-08-17-term-manager-design.md`

That spec's §4.2 ("Rename a term") says renaming is *"Allowed, encouraged, and safe… Folder names
come from **abbreviations**, not labels, so a rename does not rename a folder."*

⚠ **This is true for coded tiers (Department, Unit, and any below-Unit level with
`abbreviated: true`) and FALSE for the fixed tiers — Year and Document Type.** Those two are named
by `!builtIn` **enforcement**, deliberately excluded from ever carrying a code
(`2026-09-09-below-unit-abbreviations-design.md` §2.1: *"Not offered on a `fixed` level… nothing needs
a code for `2024`"*). Their folder name comes straight from the sanitized term **label**. §4.2's
blanket claim is exactly the live incident above: rename the label, and the folder name it drove is
now stale — with nothing tracking "what this folder used to be called" to compare against.

**This spec does not propose building `2026-08-17-term-manager-design.md`.** The client raised that
idea again this session (*"that would be the ideal way of having an interface, because in every
other flow in folder management client trips on that manual step of going into the term store"*),
then explicitly stepped back from it: *"hmm I think best to let it be manual."* The bigger in-app
term-authoring CRUD tool stays parked at its own 2026-08-17 spec, unbuilt, for a later decision — this
spec is the narrower fix the client actually asked for this time.

## 2. Why this is a real gap, not a labelling issue

Renaming a **coded** term (Department, Unit, an `abbreviated: true` below-Unit level) already
detects and warns correctly, because `CRS Term Abbreviation` rows are keyed on **`TermGuid`**, which
a label rename never touches — the existing "✎ was TAX → BE — live folder will be renamed" mechanism
(`AbbreviationManager.tsx`) simply compares the row's *stored* value against the *current* one for
that same GUID.

**Year and Document Type have no such row at all.** Per the 2026-09-09 decision, uncoded levels are
deliberately excluded from the abbreviation screen's term walk — *"An uncoded level names from the
label and needs no code; listing its terms would inflate `missing`."* So there is nothing anywhere
that remembers "this GUID used to be called PowerPoint Slide Deck" — the migrator can only ever
compare a folder's **current** name against the term's **current** label, and a renamed term simply
stops matching its old folder. The old folder becomes an unresolvable stray; the new label gets a
brand-new folder on the next upload. This is not a display bug in the migrator — there is genuinely
no persisted link for it to check.

## 3. The mechanism

**Extend the SAME tracked-row/rename-warning machinery to Year and Document Type terms**, using the
term's own **label** as the tracked value in place of a separate code (since these tiers have and
need no abbreviation).

1. `AbbreviationManager`'s walk stops excluding fixed-tier terms. Their rows carry no editable
   "abbreviation" field (there is nothing to type — the folder name IS the label), but the row still
   exists, keyed on `TermGuid`, storing the label at the time it was last seen.
2. On the next visit to this screen after a term-store rename, the stored label and the live label
   differ for that GUID — reusing the exact "✎ was X → Y — live folder will be renamed" warning
   already built for coded terms, just driven off the label field instead of a code field.
3. **The migrator does a plain folder RENAME, not a subtree move.** Nothing inside the folder needs
   to move — the documents stay exactly where they are; only the folder object itself is renamed
   (`folder.moveTo(newName)` in place). This is cheaper than the existing add/remove/reorder
   migrations, which genuinely relocate files.
4. **Client, explicitly:** *"let it be manual… ensure this flow is added into the current Rename or
   re-code, so the steps will be named to Rename Term Abbreviation or Full Name."* So the admin still
   does the rename in the classic Term Store tool themselves (no auto-triggering on save), then
   visits this now-relabelled step in the existing "Rename or re-code a folder" flow to confirm and
   apply it — same manual gate the coded-term rename already uses.

## 4. What changes on screen

- The Abbreviations step's title (wherever it renders inside the "Rename or re-code a folder" flow)
  changes to something naming both halves it now covers — client's own phrasing, **"Rename Term
  Abbreviation or Full Name"** — so an admin who only renamed a label (no code involved) still
  recognises this is the right screen.
- Year/Document Type rows on that screen render read-only (no code box — there is nothing to type),
  showing only the current label and, when it has changed since last seen, the same "✎ was X → Y"
  note coded rows already get.

## 5. Open questions

1. **Naming**: is "Rename Term Abbreviation or Full Name" the exact final wording, or a starting
   point? (Taken verbatim from the client's own message for now.)
2. **Does this also need to cover per-unit, uncoded SubUnit-style terms** that were authored before
   the 2026-09-09 codes feature and never opted in — or is the scope strictly Year/Document Type,
   the two tiers that can never be coded at all? Leaning toward the latter (Year/Document Type only),
   since an uncoded-but-codeable below-Unit term can simply be switched to `abbreviated: true` and
   get full rename-safety through the existing mechanism instead.
3. **First-run backfill**: on a site where Year/Document Type terms already have renamed-but-strayed
   folders (this client's live PowerPoint Slide Deck / Slide Deck pair), does turning this on
   automatically detect the existing drift, or does the first visit to the relabelled screen simply
   start tracking from whatever the CURRENT label is (meaning the already-strayed pair needs a manual
   one-time folder rename/merge by the admin, since there's no stored "before" value to compare
   against for a rename that already happened before this feature existed)? Almost certainly the
   latter — flag it plainly so the existing Slide Deck stray isn't expected to self-heal.

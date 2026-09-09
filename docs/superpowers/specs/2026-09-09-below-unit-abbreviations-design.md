# Term abbreviations for below-Unit levels (sub unit and shared folder terms)

**Client, 2026-09-09:** *"I just want to ensure that they can use term abbreviations with no errors
for both unit and shared folders."* Two reasons: **paths are getting long**, and **consistency —
everything else is coded**.

⚠ A cheaper alternative was offered and DECLINED: renaming the subunit terms shorter
(`Expatriate Formalities Subunit` -> `EF Subunit`) buys ~75% of the path saving and all of the
visual consistency for no code, no rows and no migration. The client heard that and chose codes.
Recorded so nobody re-proposes it as though it had not been considered.

**Status: pure modules BUILT (§3 steps 1-2). Screen, flow step and plumbing NOT built.**

---

## 1. What exists today

A below-Unit folder is named from the **sanitized term LABEL**, which is the entire reason those
tiers are cheap: no abbreviation row, no Folder Map row, no group, no reconciliation pass. Add a
term, get a folder.

Naming and matching live in **two pure modules**, which is what makes this tractable:

| Site | Module | Today |
|---|---|---|
| Upload (both web parts) | `folderChain.ts` -> `buildOnDemandSegments` | `sanitizeFolderSegment(picked.label)` |
| Which tier does this folder belong to | `subtreeMigration.ts` -> `assignSegments` | matches the sanitized LABEL |
| Is this folder at the right depth | `subtreeMigration.ts` -> `classifyChild` | same |
| Where should it go | `subtreeMigration.ts` -> `planLeaf` | reuses the EXISTING folder name |

`AbbreviationManager` caps its walk at `maxDepth = seg.levelNames.length` on purpose — *"anything
deeper belongs to a below-Unit tier, which names its folder from the term label and needs no code."*

---

## 2. The design

### 2.1 Codes are opted in PER LEVEL, never "all below-Unit terms"

⚠ **Requiring a code on every below-Unit term would require one for `Year`.** What is the code for
`2024`? And `Document Type` would need ~20 rows to turn `Agreement` into `AGR`, which nobody asked
for — and switching either on would move EVERY document in the system.

So each level carries `abbreviated?: boolean`.

⚠ **Absent means FALSE — the opposite of `permissioned`, deliberately.** `permissioned` defaults TRUE
because the wrong answer there is loud (an extra ACL'd folder). Here defaulting true would make every
existing segment instantly demand codes it does not have and refuse every upload. Absent must mean
"behave exactly as today".

⚠ **Not offered on a `fixed` level** (Year, Document Type). Nothing needs a code for `2024`.

**The flag lives in the chain, so turning it on IS a structure change** — it stages to
`PendingLevels`, shows `CHANGE PENDING`, and the migration applies it. No new machinery; it drops
into the flow the client already knows.

### 2.2 A term with no code REFUSES — it does not fall back to the label

Client, 2026-09-09: *"WHy not we force them to not be able to create untill they provide a term
abbreviation, the current design does the same, this should be more safe."* Right, and better than
the fallback first proposed here: a fallback means a tree holding `EFS` beside
`General Admin Subunit`, which defeats the consistency the codes are for.

⚠ **The force cannot live at term CREATION.** The Term Store is Microsoft's own UI; there is no hook,
and nothing built here can make it demand an abbreviation.

**It lives in the flow instead, and that gate already exists.** `blocksNext` holds the `abbreviations`
step while any term lacks a code, and the reconcile step has an `abbreviationsComplete` lock.

⚠ **The structure flow has no abbreviations step** (`pause -> levels -> migrate -> reconcile ->
resume`). One must be added between `levels` and `migrate`, or nothing forces the codes in.

### 2.3 An uncoded term is still OFFERED, and refused by name at submit

⚠ **Hiding an uncoded option from the dropdown is the dangerous version.** A unit whose subunits are
ALL uncoded would then have an empty option list, `decideTier` returns `skip`, the tier does not
apply, and **the document files one level shallower, silently** — the SDG defect of 2026-08-26, where
a blank term set hid Year and Document Type for a whole segment.

So the option is still shown, and the upload is refused naming the term:
*"'General Admin Subunit' has no folder code — ask your administrator to add one."*
Loud, actionable, cannot misfile.

**This also removes a separate decision.** A failed abbreviation READ is now just "no code" and takes
the same refusal path — one rule, no extra fail-closed branch to design.

### 2.4 The rename is the EXISTING "Move existing folders" run

Existing folders are named by label; coded ones would be named by code. Left alone that is two folders
per term, which is worse than either naming alone.

⚠ **It resolves with no new pass, provided two rules change TOGETHER:**

- `assignSegments` / `classifyChild` match **code OR label**, so an existing `EF Subunit` folder is
  still recognised as that tier rather than becoming a stray;
- `assignSegments` records the **canonical** name (code where coded) rather than the name the folder
  currently has, so `planLeaf` builds a destination that differs from where the folder sits.

The run then treats it as an ordinary **move**, which is what a rename is. Collisions raise the
existing rename form.

⚠ **Half of this is the failure that matters.** Destinations built from codes while matching still
keys on labels turns **every existing below-Unit folder into a stray**: the migration reports "needs a
value chosen" for everything and nothing moves. Buah's failure mode by a new route. Pinned by tests.

### 2.5 Sub units and shared levels are one rule, different shapes

| | Terms to code | Where the terms live | Sibling scope |
|---|---|---|---|
| Sub unit | children of EVERY unit — MHO ~120 across 49 units | one level deeper in the SEGMENT's own set | within each unit |
| Shared level | the set's top terms — 2 to ~20 | a DIFFERENT term set entirely | within the set |

⚠ **The screen work is the reverse of the volume.** A sub unit extends the existing walk one level;
a shared level needs `AbbreviationManager` to walk term sets it does not read today and group the
rows under each.

⚠ **Step 3 must read `PendingLevels ?? Levels`.** Step 2 stages the level; step 3 has to show that
level's terms before step 4 applies it. Same `row.pending ?? chain` the migrator already uses. Miss
it and a newly added shared level's terms never appear on the step meant to force their codes.

⚠ **A shared set may be used by more than one segment, and the abbreviation row is GLOBAL** (keyed by
term GUID). The FLAG is per level per segment, so one segment can be coded while another using the
same set stays on labels; each is internally consistent, and if the second switches on later the
codes are already there.

---

## 3. Build order

1. **`folderChain.ts`** — naming from code-or-label; report uncoded terms. **DONE**
2. **`subtreeMigration.ts`** — match either, assign the canonical name, destinations from the code.
   **DONE**
3. **`AbbreviationManager`** — walk one level deeper for a per-unit tier, walk the sets named by
   shared levels, off the pending chain.
4. **Folder levels** — the per-level toggle, not offered on `fixed` levels.
5. **New `abbreviations` step** in the structure flow, gated by the existing rule.
6. **Upload + migrator plumbing** — read the codes; refuse and name per §2.3.

Steps 1 and 2 are the whole correctness argument and are fully testable without a tenant.

---

## 4. What this cannot break

Below-Unit folders carry **no Folder Map row, no group and no unique permissions** — everything below
Unit inherits the Unit folder. Permissions, approval, routing and metadata are untouched. The two
realistic failure modes are §2.4 (folders stray — loud, fixed by correcting the matcher and
re-running) and a duplicate folder, which §2.3 prevents by refusing rather than falling back.

---

## 5. Do the cutover NOW, while it is small

MHO has **two** subunit terms today with almost nothing filed under them. That is 2 codes and a
handful of renames. Every document filed before the cutover makes step 4 longer. The same window as
renaming the terms: open now, closing as documents accumulate.

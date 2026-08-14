# Folder Management — guided flows instead of tabs

**Date:** 2026-08-14
**Status:** BUILT 2026-08-14, not yet site-tested.
**Replaces:** the five-tab bar from `2026-08-12-term-abbreviation-page-design.md` §6 as the *front door*
(the tabs survive behind "All tools")
**Host:** `folderManager/components/FolderAdmin.tsx` (today a 7-line shell)
**Rules:** `src/shared/folderFlows.ts` (new, pure, tested)

---

## 1. Why

> "right now it is showing tabs, its confusing and client doesnt know how it works. So here is my idea,
> remove the CRS Term Abbreviations. Folder Structure Management, Folder Reconciliation. Instead we
> replace with button such as. Adding a new segment, when user clicks into it, it will only shows the
> neccesary tabs for them to go through, the client will have to fullfill the requirements one by one
> before moving to the next step of tabs."

The tab bar was ordered by when the work happens, which helped and was not enough: it still presents five
equal doors and no indication that four of them are steps of one job. The order is the thing that breaks
— miss Term Abbreviations and reconciliation creates **nothing**, silently, with no error.

A flow states the order, and states what is left.

## 2. The five flows

Steps marked **(term store)** happen outside this tool entirely. They are shown as steps anyway, because
three of the four flows cannot start without them, and an admin who does not know that is stuck at step
one wondering why nothing works.

### Flow 1 — Add a new segment
1. **(term store)** Create the term set and its terms, nested to exactly the tier depth you will name.
2. **New segment** — name it, name its tiers, point at the term set. *Refuses on depth mismatch.*
3. **CRS Term Abbreviations** — a folder code for every term.
4. **Group Management** — create the unit groups.
5. **Folder Access** — map each group to its segment, tier and role.
6. **Folder Reconciliation** — build the folders and apply the permissions.

### Flow 2 — Add a department or unit *(the everyday one)*
1. **(term store)** Add the term under the right parent.
2. **CRS Term Abbreviations** → 3. **Group Management** → 4. **Folder Access** → 5. **Folder Reconciliation**

Steps 2–5 of flow 1, and **the job they will actually do most weeks**. Given its own button rather than
being reached by starting "Add a new segment" and skipping the first half — which is what an admin does
when the frequent case has no home, and it leaves them one wrong click from creating a duplicate segment.

### Flow 3 — Change a segment's folder structure
1. **Recommended: pause the Auto-route and folder-approval flows.** *(Power Automate — outside.)*
   **Not required** — see §4c. Advisory, never a gate.
2. **Folder Structure Management** — edit the levels. Writes `PendingLevels`; nothing moves yet.
3. **Move existing folders** — migrate, then it applies the new shape **as its own last step**.
4. Turn the flows back on if you paused them, and verify.

Reconciliation is **not** a step here: it walks the term tree, not `Levels`. Saying so in the flow
prevents the reasonable-but-wrong instinct to run it after a structure change.

### Flow 4 — Rename or re-code a folder
1. **(term store)** Rename the term. **Never delete and re-add** — a new GUID orphans the Abbreviation,
   Folder Map and Group Map rows at once.
2. **CRS Term Abbreviations** — change the code.
3. **Folder Reconciliation** — renames the live folder.

### Flow 5 — Retire a segment *(destructive)*
1. **Move existing folders** — get the documents out first.
2. **Segments → Delete** — removes the mode row and Group Map rows; folders are a separate typed opt-in.

Visually separated from the other four and never styled as a "get started" card. It is the only flow that
destroys anything.

## 3. Architecture — drive `FolderManager`, do not dismantle it

`FolderAdmin.tsx` becomes the host and renders one of three things:

- **the flow picker** (default),
- **a flow runner** — the step rail plus the current step's screen,
- **All tools** — `<FolderManager />` exactly as today.

`FolderManager` gains two optional props, `initialTab` and `hideTabs`, and nothing else changes in it.
Steps it already hosts (Abbreviations, Folder levels, Move existing folders, Segments, Reconciliation) are
rendered as `<FolderManager initialTab="…" hideTabs />`.

**Reconciliation is never extracted.** It lives inline in a 4,186-line file and is the most complex,
most site-verified code in the project. Extracting it to make it mountable would put the riskiest thing
here at risk for a navigation change. Driving the existing host costs one re-mount per step — the admin is
clicking through deliberately, so that is affordable.

**Group Management and Folder Access are mounted directly** from `userAccess/components`, because
`FolderManager` does not host them. They also stay standalone on the landing page under User Access
Management: **one component, two mount points, never a second copy** — the rule the three
`Folder Structure` tabs already follow.

## 4. Gating — order first, verification only where it is honest

The wizard's value is the **order**. Verification is a bonus, and claiming more of it than we have would
make the wizard a liar — worse than the tab bar, because a green tick is believed.

**Cheaply and reliably verifiable, from four reads already performed elsewhere:**

| Fact | Source |
|---|---|
| The segment exists | `DMS Config` mode rows |
| Its groups exist | site groups matching the segment's code prefix |
| Its Folder Access rows exist | `DMS Group Map` rows for that term set |
| Its folders exist | `DMS Folder Map` — reuse `segmentProvisionState()` |

**Abbreviation completeness is verifiable, but only once the step is open.** Knowing every term has a code
requires walking the whole term tree — and `AbbreviationManager` **already does that walk**
([:188-216](../../../src/webparts/folderManager/components/AbbreviationManager.tsx)), already counts the
terms with no code (`missing`, at :431) and already displays the count. So the number is exact and free
the moment the step is opened; what is unaffordable is pre-walking every segment on the picker screen
before one has been chosen (4–12 walks for a list being glanced at).

Hence: **"not checked yet" on the picker, an exact count once opened, and a real lock on reconciliation
from that point** — see §4a. This corrects an earlier draft of this spec which called the check
unaffordable and gave up the one gate worth having.

### 4a. The four real locks

Each rests on one definitive read. A **failed** read never locks — unknown falls through to allow.

| Lock | Backed by | Why it earns a lock |
|---|---|---|
| Flow 1: no step 3+ until the segment exists | `DMS Config` mode row | Nothing downstream has a subject without it |
| **Reconciliation blocked while any term lacks a code** | the tree walk `AbbreviationManager` already performs | **The silent failure.** Reconciliation skips an uncoded term and creates no folder, with no error anywhere |
| Flow 3: no "Move existing folders" until `PendingLevels` is set | mode row field | Migrating with nothing staged does nothing and looks broken |
| Flow 5: cannot delete a segment that is not there | mode row | — |

The abbreviation lock is available **only after the step has been opened once**, because the walk is what
produces the count (~115 requests for GHO). Before that the rail reads "not checked yet" — honest, and
free. Pre-walking every segment on the picker is the part that is unaffordable, not the walk itself.

### 4b. Flows 2 and 4 ask what you are doing

Both open by asking for a **subject**: "Adding `[Treasury]` under `[Group Finance ▾]`" (flow 2), or "which
term did you rename" (flow 4). That turns an unverifiable step into a checkable one — walk the segment's
tree, look for that label under that parent — and the walk is the same one the abbreviation step performs,
so it costs nothing extra.

It also earns its keep beyond gating: the rail can name the job instead of being generic, the abbreviation
step can point at **that** row rather than leaving the admin to find the blank one among fourteen, Folder
Access can pre-fill segment/tier/term, and Group Management can pre-fill the name `suggestGroupName`
already derives (`GHO_GF_TREASURY_UPL`).

**Matching must be normalised, and "not found" must never say "you have not done step 1".** GHO contains
`Group Legal, Risk ＆ Compliance` with a **fullwidth ＆** — SharePoint requires it — so a normal `&` fails
to match a term that exists perfectly well. Compare on a trimmed, case-folded label with `&`/`＆` and
whitespace folded, and word the failure as *"no term matching that name — check the spelling, or add it in
the term store"*. An unreadable tree stays unknown and passable.

### 4c. Power Automate: leave it running, but say so

The earlier draft made "turn off Auto-route and the folder-approval flow" a safety-critical step, on the
grounds that it could not be verified. **Re-reading the flow config, that overstated it.** Auto-route is:

```
When an item is created or modified   (approval library; trigger condition {IsFolder} is false)
Get item → Condition: ModerationStatus is equal to Approved
├─ True  → copy to Documents, stamp, delete source
└─ False → no-op
```

So a **pending or rejected** file being moved fires the flow, which reads the status and does nothing —
the case that covers almost everything in that library (§5.6 of the Auto-route spec already notes it fires
several times per upload and all but one run correctly takes the False branch). Folders the migrator
creates get approved twice, idempotently. `Documents` is wired to neither flow. And the 2026-08-13 run that
copied a file still *Waiting for Approval* was caused by the then-missing `{IsFolder}` trigger condition,
which is now in place.

Two residual effects, neither a reason to gate:

1. **A stale approved file left in the approval library WOULD be acted on** — copied out and the source
   deleted, which is the move that should already have happened. But the destination path is derived from
   the SOURCE path, so firing mid-migration can file it in `Documents` under the old shape. Rare (Auto-route
   deletes on success, so one only lingers if the flow was off or failed) and recoverable, because the
   migration records no progress and a second run re-derives what is left.
2. **Quota.** A large migration can burn Power Automate's daily action limit on no-op runs, and then the
   next genuine approval is not routed until it resets — **silently**, because a flow that never fires
   leaves no run history.

**So the step advises both ways:** you may leave them running, and pausing them is recommended for a large
run. Advisory, not typed, not blocking — we cannot verify it either way, and a gate on something
unverifiable is theatre. The durable fix is a moderation-status trigger condition on Auto-route (see the
Auto-route spec §5.6), after which moving pending files fires nothing at all.

### 4d. What no amount of checking will reach

**Existence is checkable; intent is not.** Every tick in this wizard means "this thing exists", never
"this is right". Specifically out of reach:

- **Whether the right people are in a group.** Membership is readable; correctness is a human judgement,
  and an admin will assume the wizard checked it.
- **Whether an abbreviation is the code they meant.** Existence and sibling-uniqueness are checked
  (`findCollisions`); `TRS` versus `TREAS` is not, and changing it later renames a live folder.
- **Whether a term went where they meant.** The named parent is checked. A wrongly named parent passes.
- **Whether the client finished in the term store** — they may add a department and forget its units.
- **Group Management having nothing to do.** Checkable via the naming convention, but the convention is a
  *suggestion* (`suggestGroupName` writes nothing), so a hand-named group defeats it. Marked, never locked.
- **Anything in Power Automate** (§4c).

Three rules, each following this codebase's existing discipline:

1. **A step is never hard-blocked by a failed read.** Unknown means *warn and allow*. One transient error
   must not strand an admin mid-flow with no way forward.
2. **Steps are enterable out of order**, with the out-of-order ones marked. The rail says what is
   incomplete; it does not padlock. An admin who knows better must not be trapped — and "fulfil before
   moving on", enforced literally, becomes unusable the first time a check is wrong.
3. **Progress is DERIVED on every load, never stored.** A saved step number goes stale the moment somebody
   does a step by hand, and a wizard confidently wrong about where you are is worse than no wizard.
   Deriving it also makes every flow resumable for free — which matters, because flows 1, 2 and 4 all
   leave the tool for the term store and will not be finished in one sitting.

## 5. Which segment?

Flows 2–5 act on an existing segment, so each opens with a segment picker; flow 1 creates one and
therefore does not. The picker shows each segment's readiness from the same facts, so "which of these is
half-built" is answerable before choosing — reusing `segmentProvisionState()` rather than inventing a
second notion of ready.

## 6. What this does not do

- **It does not create term sets or terms.** It cannot; that is the term store. Flows 1, 2 and 4 name it
  as step one instead of hiding it.
- **It does not replace the tabs**, only the front door. "All tools" keeps every screen reachable, which
  is also what keeps testing possible.
- **It does not pre-check abbreviations on the picker** — that would be 4–12 term-tree walks for a list
  being glanced at. It *does* gate reconciliation once the step has been opened (§4a), which is the gate
  that matters.
- **It does not require the Power Automate flows to be off** (§4c), and it cannot verify either way. It
  advises both: leaving them on is safe, pausing them is recommended for a large run.
- **It adds no new capability.** Every screen already exists; this is ordering only. Nothing about
  permissions, folder naming or reconciliation behaviour changes.

## 7. Risks

- **`FolderManager` re-mounts per step**, re-running its mount-time loads. Acceptable, and the reason the
  flow runner does not try to keep it alive across steps.
- **The step rail can disagree with reality** where verification is absent or a read failed. Mitigated by
  never padlocking, and by labelling "not checked" as exactly that.
- **Two mount points for Group Management and Folder Access** must not become two copies. Enforced by
  review, as with `StructureManager` today.
- **`FolderManager` is already 4,186 lines** and over the lint ceiling. This adds two props to it and puts
  the new code in `FolderAdmin.tsx` and `shared/folderFlows.ts`, so it does not grow meaningfully.

## 8. Testing

1. Each flow opens on its first incomplete step, and re-opens there after a reload.
2. Complete a step by hand outside the flow; the rail must notice on next load (derived, not stored).
3. Break the Group Map read; the rail must show "not checked" and still let you through.
4. Flow 2 on a segment with a new unit: abbreviation → groups → Folder Access → reconciliation, and the
   folder appears with the right ACL.
5. Flow 3 must **not** offer reconciliation, and must say why.
6. "All tools" shows all five tabs, behaving exactly as before.
7. The landing page's three Folder Management rows still resolve — they now land on flows or All tools, so
   the `#tab=` handling needs re-checking against §3.

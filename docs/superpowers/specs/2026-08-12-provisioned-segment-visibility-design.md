# A segment is offered only once its folders exist

**Date:** 2026-08-12
**Status:** built. Existence gate shipped, then **corrected to a per-user reachability probe** after a
live test on 2026-08-12 proved existence insufficient — see §2.1 before changing anything here.
**Client instruction (2026-08-12):** *"I also do not want the half built segment to show, the UI UX
experience should be that way, once the Segment and structure and group is properly assigned only
then the uploader should be able to upload."*

Supersedes the closing paragraph of §5 of
[2026-08-12-add-segment-design.md](2026-08-12-add-segment-design.md), which said a new segment would
be visible immediately and that a "draft" state should not be added. The client wants the opposite,
and — see §2 — the mechanism that delivers it is not a draft state at all.

---

## 1. Two symptoms, one cause

These were tracked as separate problems. They are the same one:

1. **A newly added segment appears in the upload form before it has any folders.** Slice B writes a
   `mode` row; the form reads mode rows live; the segment is offered.
2. **A user added to a unit group can pick their unit before reconciliation has run** (observed
   2026-08-11, and the origin of the HTTP 403 in the upload form). The dropdown offers the unit, the
   upload then fails at the destination.

The cause in both cases: **the form decides what to offer from the `DMS Group Map` and the term tree,
and never asks whether it can actually write there.** `resolveValidPaths` walks group membership
against the term ancestry and returns every leaf whose chain resolves. Nothing about the destination —
neither its existence nor its ACL — is consulted until the user presses Upload.

So the form offers a path built from *intent* and only discovers at write time whether that path is
*usable*. Both symptoms are that gap, at different tiers — and note that symptom 2 is about the ACL,
not the folder, which is why an existence-only gate did not fix it (§2.1).

## 2. The gate is DERIVED, not a flag

### 2.1 Existence is NOT the question — reachability is (corrected 2026-08-12, live test)

The first build of this gate asked *"does the folder exist?"* and shipped. A live test the same day
disproved it: two new groups (`NBPOL_T_T_UPL`, `NBPOL_S_C_UPL`) were created without running
reconciliation, and their units were **still offered** to a non-admin uploader.

The reason is in reconciliation's structure. It creates folders by walking the **term tree**, keyed on
`DMS Term Abbreviation` (`FolderManager.tsx` ~1836), and assigns group ACLs in a **separate pass**,
keyed on `DMS Group Map` (~2384). So:

- Creating a group creates **no** folder and removes none.
- Those units already had folders and Folder Map rows from an earlier NBPOL run.
- Existence answered "yes", correctly, and the gate offered a unit the user could not write to.

**Existence is a property of the folder. Being able to upload is a property of the folder AND the
user.** No amount of Folder Map reading can answer the second, which is why symptom 2 in §1 survived
the first fix. The gate now probes, per authorised path:

```
GET /_api/web/GetFolderById(guid'<UniqueId>')/ListItemAllFields/EffectiveBasePermissions
```

and keeps the path only if `AddListItems` is set. That subsumes existence — a missing folder answers
404 — and is the same question the upload itself will ask.

**It tests upload, not read.** Read is not sufficient and is genuinely reachable without upload: a PIC
who is also in their unit's base group holds Read on the unit folder through that group while their
`*_UPL` group is still ungranted (2026-08-08 §5.8). A visibility probe would call that ready and
deliver the identical 403.

Cost is one request per authorised path, in parallel, at mount. A PIC has one or two; the department
fan-out that yields many belongs to Documents-side viewers, who are not uploaders.

Statuses map deliberately: `401`/`403` → denied; `404` → missing (security trimming also answers 404,
so this is *not* proof of deletion — but both mean "cannot upload here" and carry the same fix); any
other non-OK status → **unknown**, which is KEPT (§5).

`Low` is read arithmetically, not with `&`: JavaScript's bitwise operators coerce to a **signed**
32-bit int and Full Control returns `Low = "4294967295"`. A single-bit `&` test happens to give the
right answer on a negative number, which is precisely the kind of correct-by-accident that breaks when
someone extends it to two bits.

### 2.2 The bulk-upload web part is existence-gated only

`BulkUpload.tsx` writes into **`Documents`**, not the approval library — it resolves the Staging unit
folder and swaps the library segment (`toDocumentsPath`). So the correct probe target there is the
*Documents* folder, where `UPL` is downgraded to **Read** by design (2026-08-09 persona change).

Probing `AddListItems` there would empty that form for every PIC. Strictly that is accurate — a PIC
genuinely cannot bulk upload — but it is a client-facing behaviour change on a tool the client is told
is temporary, and not a side effect to smuggle in while fixing the upload form. So BulkUpload keeps the
existence gate, with the reasoning recorded at its import site. If the client wants that form hidden
from PICs entirely, that is a decision to take on its own terms.

### 2.3 Why not a flag

Readiness is **"a `DMS Folder Map` row exists for the leaf term"** — the exact condition the upload
path itself checks (`lookupFolderMapping` in `Form.tsx`). Nothing new is stored.

A `Ready` / `Active` column on the mode row was the obvious alternative and is worse:

- It can be **wrong**. An admin ticks it when they believe the segment is finished; the gate then
  asserts something nobody verified. A derived gate cannot disagree with reality because it *is* the
  reality — the same row the uploader's write will look for.
- It is a **second staging mechanism**. `PendingLevels` already stages structure changes
  ([2026-08-11-subtree-migration-design.md](2026-08-11-subtree-migration-design.md)); a second,
  differently-shaped "not live yet" concept invites the two to disagree.
- It only fixes symptom 1. Symptom 2 is per-unit, not per-segment — a flag at segment level cannot
  express "these three units are provisioned and that one is not".

And it needs no migration: existing sites have Folder Map rows for every provisioned unit already, so
the gate is a no-op everywhere it should be.

## 3. Behaviour

Filtering happens at the **leaf**, and the tiers above disappear as a consequence:

| State | What the uploader sees |
|---|---|
| Some units reachable | Only those units. The tiers above list only what leads to them |
| Unit folder exists, ACL not granted to this user | **Absent** — the case existence-checking missed (§2.1) |
| No unit in a segment reachable | **The segment is absent** — from the Segment dropdown and, if it was the only one, from the Business Segment / Project toggle |
| No unit reachable at all | The empty state, with the message in §4 |
| Folder Map unreadable, or a probe inconclusive | **Everything is offered, as before** — see §5 |

This is why no per-segment logic is needed to satisfy the client's instruction. A half-built segment
has no provisioned units, so every one of its paths is filtered out and the segment vanishes on its
own. Hiding it is a *consequence* of the leaf gate, not a second rule.

**Site admins are unaffected.** `privileged` users get the full manual cascade over the term set and
no `validPaths` at all; they are the people building the segment and must be able to see and test it.
The Structure Manager's closing checklist is where an admin is told what remains (slice A §5.1).

## 4. The empty state must name the real cause

Today, an uploader with groups but no folders gets:

> Your account isn't fully provisioned to upload — you need membership at every level plus the unit
> uploader role. Contact your administrator.

Which is **wrong**, and wrong in an expensive direction: it sends the administrator to check group
membership that is already correct. So the two states are now distinguished by whether
`resolveValidPaths` returned anything *before* the folder filter:

- **No paths before filtering** → the existing message. It is accurate: this is a membership problem.
- **Paths existed, all withheld** → a new message naming the actual fix:

  > Your unit isn't ready to receive uploads yet. Your DMS administrator needs to run folder
  > reconciliation — and if the unit has no folder at all, give it an abbreviation in the DMS Term
  > Abbreviation list first.

Reconciliation **leads**, because after §2.1 the common case is an ungranted ACL on a folder that
already exists, and reconciliation is what grants it. The abbreviation is named second rather than
dropped: a unit lacking one is skipped by every run, so for that unit reconciliation alone silently
changes nothing — the same reason the upload-time message at `Form.tsx` names both.

It is shown only when `known` is true on **both** filters. An unreadable list or a single inconclusive
probe withholds nothing, and asserting "not ready" then would send an administrator to reconcile an
intact, correctly permissioned tree.

## 5. An unreadable Folder Map must NOT hide anything

If the Folder Map read fails, the gate is skipped entirely and every path is offered.

This is the rule already established for the stale-chain guard (gotcha 10b) and for
`AllowedFileTypes` (gotcha 11): **empty is not unknown.** A read failure is evidence of nothing, and
a transient error that silently empties every dropdown takes the whole form down for every uploader
on the site — far worse than the state the gate exists to prevent. The failure mode we accept is the
old one: the path is offered, and the upload fails loudly at the destination with a message that
names the fix.

The upload-time check at `Form.tsx:1373` therefore **stays**. It is now a backstop rather than the
only line of defence, and it is what makes the degraded path safe.

## 6. Implementation

- **`src/shared/segmentReadiness.ts`** (new, pure, unit-tested — 19 tests) —
  `filterProvisionedPaths(paths, mappedTermGuids)` for existence and
  `filterReachablePaths(paths, verdicts)` for per-user access. Generic over the path shape so both web
  parts share them, and free of SharePoint imports so they test without a tenant.
  - `mappedTermGuids === null` means *unknown* and returns the input untouched; that distinction is
    the whole of §5.
  - `verdicts` is **positional**, and a missing entry counts as `unknown` — a short array fails OPEN
    rather than silently hiding the tail.
  - GUIDs are normalised (lowercased, braces stripped) before comparison: a stored `{ABC…}` and a
    term store's `abc…` are the same term, and a case-sensitive compare would hide every unit on the
    site — presenting as a total outage rather than a comparison bug.
- **`src/shared/dmsFolderMap.ts`** — `probeFolderUploadAccess()` returning
  `granted | denied | missing | unknown`. Shorter retry ceiling than `probeFolderById` (3 attempts,
  8s cap) because this sits on the form's mount path and `unknown` already fails open: a throttled
  site must not hold the form on a spinner.
- **`Form.tsx`** — `loadFolderMapRows` joins the mount-time `Promise.all` with `.catch(() => null)`;
  the existence filter runs first (free — the rows are already in memory, and it skips a probe for a
  unit with no row at all), then the survivors are probed in parallel. `awaitingFolders` requires
  `known` on both.
- **`BulkUpload.tsx`** — existence filter only, for the reason in §2.2.
- Both web parts keep their duplicated `resolveValidPaths`; unifying those is a separate change and
  not one to make while altering behaviour.

## 7. Testing

- **A NEW GROUP on a unit whose folder already exists, before reconciliation** — the case that broke
  the first build (§2.1). The unit must be absent. This is the test to run first, because it is the
  one an existence check passes.
- A unit with a group but **no** Folder Map row: absent from the Unit dropdown; if it was the only
  one, the segment is absent and the §4 message names reconciliation.
- Either unit **after** reconciliation: present, upload lands correctly.
- A user holding **Read but not upload** on the unit folder (in the unit base group, `*_UPL` not yet
  granted): absent. A read-only probe would wrongly offer it.
- A segment with a `mode` row and no folders at all: absent for uploaders, **present** for a site
  admin.
- A user with **no** groups: unchanged membership message, not the new one.
- Folder Map read forced to fail: every path offered, upload still refuses loudly at the destination,
  and the §4 message is **not** shown.
- A Folder Map row whose GUID is stored in a different case: unit still offered.
- A site admin (`privileged`): every segment visible, including one with no folders — and **no probe
  requests**, since admins resolve no `validPaths`.
- The bulk-upload web part: unchanged behaviour for a PIC (§2.2).
- A provisioned unit in one segment and an unprovisioned one in another: only the first segment is
  offered, and the toggle hides the emptied side.

## 8. Related

- [2026-08-12-add-segment-design.md](2026-08-12-add-segment-design.md) — slice B, whose §5 this changes
- [2026-08-10-structure-manager-ui-design.md](2026-08-10-structure-manager-ui-design.md) — §5.1, the
  six-step checklist this gate makes enforceable rather than advisory
- [2026-08-11-subtree-migration-design.md](2026-08-11-subtree-migration-design.md) — `PendingLevels`,
  the staging mechanism this deliberately does not duplicate

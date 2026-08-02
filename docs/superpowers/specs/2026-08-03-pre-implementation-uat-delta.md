# Pre-Implementation UAT — Delta for 1.0.71 → 1.0.75

**Date:** 2026-08-03
**Branch:** `feat/folder-abbreviations` · **Build:** 1.0.75.0
**Purpose:** confirm the current build is sound before starting the access-scope work
(`2026-08-03-access-scope-mapping-design.md`).
**Not a replacement for** `2026-07-28-dms-uat-test-plan.md`, which still covers Approval
Document, Onboarding, the extensions and the Auto-route flow. That plan predates everything
below and its line counts are stale.

Every check needs a **hard refresh** — config is read once in a mount-time `useEffect`.

---

## 0. Before you start

- [ ] Deploy 1.0.75.0, hard-refresh.
- [ ] `legallyPrivilegedFor` is set in `DMS Config` to the Confidential term's GUID. Blank means
      the Legally Privileged tick never renders, and an empty row looks identical to a broken
      build.
- [ ] `AllowedFileTypes` has all five extensions ticked on the `allowedExtensions` row.

---

## 1. Upload Form

New since the last plan: three required name parts, the Document Date / Confidential Level /
Legally Privileged row, tooltips moved onto labels.

- [ ] **1.1** Upload with Document Name, Project Name or Vendor blank → refused, and the toast
      names exactly the blank ones. A single space counts as blank.
- [ ] **1.2** All three filled → saved file is
      `[Project] - [Vendor] - [Document Name] - [DDMMYY].ext`, and the "Saves as:" preview under
      Document Name matched it before upload.
- [ ] **1.3** Document Date, Confidential Level and the Legally Privileged tick sit on **one
      line**. With no Confidential level chosen, date and level split the row evenly and line up
      with Project Name / Vendor above.
- [ ] **1.4** The `ⓘ` beside **Confidential Level** opens on hover *and* on keyboard focus, and
      lists **only Confidential and Restricted** — Legally Privileged must NOT be in it.
- [ ] **1.5** The `ⓘ` beside **Legally Privileged** carries that definition, and its panel opens
      without spilling past the card's right edge.
- [ ] **1.6** Clicking either `ⓘ` does **not** open the Confidential Level dropdown.
- [ ] **1.7** Tick Legally Privileged, then change Confidential Level to something else → the
      tick disappears, and the uploaded item's `LegallyPrivileged` is **false**, not true. (The
      value is re-derived at upload; hiding the control does not clear the state behind it.)
- [ ] **1.8** File lands in `…/Unit/Year/DocumentType` with every metadata column populated,
      including `Remark`.
- [ ] **1.9** Upload to a unit with **no abbreviation row** → the message names the DMS Term
      Abbreviation list, not "re-run reconciliation".

## 2. Bulk Upload — rebuilt as a single selection

The whole screen changed in 1.0.71. Treat as new, not as a regression pass.

- [ ] **2.1** No batch UI anywhere: no "Add batch", no batch tiles, no side panel.
- [ ] **2.2** Empty state reads *"Choose multiple documents or drop them here"* with
      *"Max. 50 files."*
- [ ] **2.3** Pick 3 files → green bar reads "3 documents selected". Remove one → reads 2.
- [ ] **2.4** **Add more appends**, it does not replace. This is the one behaviour a user coming
      from the Form guesses wrong.
- [ ] **2.5** Try to exceed 50 → the message names the limit and says how many were added.
      Nothing is silently dropped.
- [ ] **2.6** **Drop** a `.exe` onto the zone — drag-and-drop, not the picker, because the
      picker's `accept` filter hides it and a drop bypasses that entirely → rejected **by name**,
      and the other files in the same drop survive.
- [ ] **2.7** Upload: per-file rows show `LOADING` then `READY`, completed rows stay visible and
      keep their order, and the header count matches the list.
- [ ] **2.8** Every file lands in `…/Unit/Year/DocumentType` carrying the **same** metadata, and
      **keeps its own filename** — no `[Project] - [Vendor] - …` composition here.
- [ ] **2.9** Re-upload a selection containing two names that already exist → the Replace dialog
      appears **for each in turn**; Yes overwrites, No records `skipped`, and files after the
      prompt still upload.
- [ ] **2.10** A fully clean run shows the same **Upload Successful** dialog as the Form, with
      Back to Document and Upload More.

## 3. Group mapping — personas

Model: `2026-07-16-highly-confidential-securing-design.md` §4.3 (on `feat/hc-libraries`).
Personas are membership combinations of atomic per-unit groups; no bundle groups are created.

| Persona | Groups |
| --- | --- |
| Head-of #1 | `{Unit}` + `_UPL` + `_APR` |
| Head-of #2 | `{Unit}` + `_APR` |
| Head-of #3 | `{Unit}` + `_APR` + `_DEL` |
| Head-of #4 | `{Unit}` + `_UPL` + `_APR` + `_DEL` |
| PIC #1 | `{Unit}` + `_UPL` |
| PIC #2 | `_HC` only |
| PIC #3 | `_UPL` only |
| SDG Employee (common) | `{Unit}` only |

Head of Department and Head of Unit use the **same four bundles** and differ only in scope: Head
of Unit covers their own unit, Head of Department every unit under the department. Department
scope does not work on any branch yet (§6 of that spec).

> **Scope limit.** `_DEL` and `_HC`, and the custom `DMS Approve` / `DMS Delete` permission
> levels, live on **`feat/hc-libraries`** and are not in this build. Head-of #3, Head-of #4 and
> PIC #2 cannot be tested here at all, and Head-of #2 cannot be tested *honestly* — see 3.6.

- [ ] **3.1 SDG Employee (common)** — `{Unit}` only. Can open the site and Home, reads their
      unit's approved documents, sees **no** Staging, cannot upload.
- [ ] **3.2 PIC #1** — `{Unit}` + `_UPL`. As above, plus upload; sees only their own pending
      items in Staging.
- [ ] **3.3 PIC #3** — `_UPL` only, no base group. Can upload; **cannot** reach Documents. Sees
      only their own pending items. Confirm they can still open the site — this persona has no
      base group, so site entry has to be coming from `DMS_SITE_MEMBERS`.
- [ ] **3.4 Head-of #1** — `{Unit}` + `_UPL` + `_APR`. Upload, and approve their unit's pending
      items.
- [ ] **3.5 Isolation** — every persona above sees **only** their own unit. Check one user
      against a second unit's folder by **direct URL**, not just by browsing; browsing is
      security-trimmed and will look correct either way.
- [ ] **3.6 Head-of #2** — `{Unit}` + `_APR`, no `_UPL`. Expected on this build: **the approver
      can still upload**, because `APR` maps to Design, which includes Add Items. That is the
      known gap, not a new defect. Record it and move on; the custom `DMS Approve` level fixes it.
- [ ] **3.7 User Access → People** — each mapped group lists its real members. Search a person's
      name: only matching people show inside each group, the hit is highlighted, groups
      auto-expand, and the summary line counts them.
- [ ] **3.8** A group with no members is called out in red, not shown as an empty success.

## 4. Folder Reconciliation

- [ ] **4.1 The ETA.** Run reconciliation **twice**. On the second (mostly "already there") run
      the estimate must track reality — this is the regression that produced "~197m left" on a
      three-minute run. Watch that the step counter advances on skipped folders too.
- [ ] **4.2** Four panels, grouped by library: Staging folders | Staging group assignments, then
      Documents folders | Documents group assignments. Each fills with its own library's work.
- [ ] **4.3** Log tabs — All, Documents, Staging, Warnings, Errors — with counts. Confirm an
      entry appears in **both** its library tab and its severity tab, and that All contains the
      prune / orphan-repair / site-entry lines that name neither library.
- [ ] **4.4 Full Name** is populated on every folder and visible in the details pane; the grid
      does **not** show the column.
- [ ] **4.5 Orphan repair.** Note a unit's term GUID, delete the term, re-create it with the
      exact same name, reconcile. Expect
      `✎ <ABBREV> — "<name>" was re-created; re-pointed to <guid> (+N group-map row(s))`.
      Then **check `DMS Group Map` actually holds the new GUID** — that write is what grants
      access, so do not take the log's word for it. Reconcile again; the folder is created.
- [ ] **4.6 Ambiguity.** Delete two same-named units (e.g. two `Tax`) and re-create both →
      reported as `AMBIGUOUS`, **not** repaired, with both candidate sets named.
- [ ] **4.7 `NO TERM`.** Leave a folder from 4.5 behind → reported, and **not** deleted.
- [ ] **4.8 Collision** — set two sibling units to the same abbreviation → the run aborts before
      creating anything.
- [ ] **4.9 Missing abbreviation** — delete one row → that term is reported `SKIPPED`, no folder
      is created, and 1.9 above shows the right message to the uploader.
- [ ] **4.10 Idempotency** — a third run changes nothing and reports no errors.

---

## 5. The decision this UAT should produce

Whether `feat/hc-libraries` lands **before** the access-scope work starts. Section 3 can only be
half-tested until it does, and the access-scope spec adds `Scope`/`Target` to the same
`DMS Group Map` list that branch adds `DEL`/`HC` roles to. Landing them in the wrong order means
reconciling two schema changes to one list against a live site.

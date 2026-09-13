# Porting the 21 Power Automate flows to SDG — what has to change, per flow

Built from the **actual exported definitions** (2026-09-09/10), not from memory. Every figure below
was counted out of `definition.json`, so it describes what is really in those flows rather than what
this project's docs claim.

⚠ **Power Automate is not in source control. This file and the exports in `PowerAutomateFlows/` are
the only record of these 21 flows.**

---

## 0. ⚠⚠ BEFORE THE FIRST IMPORT — the mistake that cannot be undone later

**Sign in as the SERVICE ACCOUNT, and pick its connection at import.** A flow's connection is baked
in at creation and cannot be reassigned afterwards; a flow imported under a person's account **stops
silently** the day that password changes, and presents months later as *"my colleague's folder
doesn't exist"*. Every prior runbook in this project says this and it is still the single most
expensive thing available to get wrong here.

⚠ **Import does NOT repoint anything SharePoint-specific.** It carries actions, expressions,
conditions and connection *references* — but **Site Address, List Name and every list GUID are
parameters inside the definition**. Nothing remaps them. That is what the rest of this file is for.

### ✅ The identity, settled 2026-09-10

**`crs@sdguthrie.com`** — the connection in SDG's own `SD Guthrie` environment. Every flow is imported
against it. **Do not import one against a personal account**, and check the `...` → *My connections*
line on each imported flow reads `crs@sdguthrie.com` before saving: a flow with one action on a
different connection is the silent-failure shape again.

### ✅ The import sequence that actually works — proven on `Audit — Documents deletions`

The first import **fails**, and that is expected rather than a problem:

```
Flow save failed with code 'DynamicOperationRequestClientFailure' …
'GetTable' failed with status code 'Unauthorized'
```

The package carries the exporting tenant's connection reference, which does not exist in SDG's
environment, so the importer cannot authenticate to validate the SharePoint actions. Take the
**"Save as a new flow"** link it offers; the flow lands as a **draft**.

Then, in this order — and the order is forced, not a preference:

1. **Resolve every "Invalid connection" placeholder to `crs@sdguthrie.com`.**
   ⚠ **The real actions are INVISIBLE until this is done** — they render as grey `Connections` blocks,
   so you cannot see or fix a single list reference first.
2. **Site Address → `https://sdguthrie.sharepoint.com/sites/CRS`** on every SharePoint action.
3. **Paste the SDG list GUIDs** from §2. ⚠ **The picker will NOT offer list names for an imported
   value** — the field is in custom-value mode, so it is paste-only. That is safe *only because* the
   GUIDs in §2 were read off SDG itself; never paste a GUID that came from inference.
   A wrong one shows as `GetTable … 'NotFound'` with `"List not found"`, which is the loud and
   therefore harmless failure.
4. **Check the trigger conditions survived.** They did on the first flow — but a dropped one is
   silent, and for the deletion flows it means folder deletions logged as destroyed documents.
5. Save → turn on → **test with a real deletion** and confirm the row lands.

⚠ **Only the SITE and the LIST GUIDs change. Library TITLES do not** — see §3.

---

## 1. What every flow needs, without exception

**All 21 reference exactly one site: `https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting`.**
So all 21 need their Site Address changed on **every SharePoint action**, not just the trigger.

⚠ **Changing the Site Address usually invalidates the List Name box** and forces a re-pick. That is
helpful — it converts a silent wrong-list into a visible empty field. **But it only happens for
CONNECTOR actions.** A list named inside an **HTTP request URI** (`getbytitle('…')`) is just text in
a string and will be carried across untouched and wrong. Those are listed in §3.

---

## 2. The nine list GUIDs, and what they are

Every `"table"` parameter in these flows stores a **list GUID**, and GUIDs are per-site — so all nine
must be re-pointed. Five are confirmed against CLAUDE.md; four are inferred from which flows use them
and are marked as such.

✅ **ALL NINE MEASURED ON BOTH SITES, 2026-09-10** (`scripts/dump-site-inventory.js`). This is the
substitution table — nothing here is inferred any more.

| List | Test site → | SDG |
|---|---|---|
| Approval for Document (`/ApprovalDocument`) | `a9342528-66be-458c-ba70-a6e3248a5133` | `eeb1bb19-ec53-4c41-8dc9-ecf255979c9b` |
| Approval for HC Document (`/HCApprovalDocument`) | `2311cd83-90aa-4077-a647-65d251a5f426` | `d935aa6d-dc48-4832-ba7e-eca485ecdff3` |
| Restricted & Confidential (`/Shared Documents`) | `e322e3a5-3687-4da3-94f8-e0b06c01dd7e` | `fbc062dd-fcd0-4458-9161-2bfc19c917fa` |
| Highly Confidential (`/HCDocuments`) | `e4fb3b2c-6f20-4bbb-a4f2-e9a5301d4080` | `122a4aa9-65fb-479a-90a7-3866d19f51b4` |
| Archive R&C (`/Archive`) | `8c28ca71-4edb-445b-8ebd-5711dacd71d4` | `647fd644-ab00-4358-9b3f-49031083fcf0` |
| Archive HC (`/HCArchive`) | `bfcd6735-7605-47e7-a0d6-bbfff41bab7b` | `fc301054-abf5-4e53-957a-6f8c58a410b8` |
| CRS Audit Log | `84ed065f-14e2-49b2-8b37-d40324587825` | `4ac219e5-b855-46a3-be4b-4e96bba546c5` |
| CRS Requests | `e47c85d5-f0ee-4190-99b8-b3c5575f95a4` | `1bcb1aa1-fbf1-4d1b-a6d2-88436824c297` |
| CRS Submissions | `f413972a-49ae-476a-8a8f-df4562349918` | `dd1bc5bd-daa7-491a-9dc7-b81212c391ab` |

⚠⚠ **AND IT IS A DIFFERENT TENANT.** SDG is `https://sdguthrie.sharepoint.com/sites/CRS`, not
`dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting`. So the **HOST** changes as well as the site
path — every hardcoded URL in an email body needs both. A find-and-replace of just
`ClarenceDMSTesting` leaves 22 links per routing flow pointing at the wrong tenant.

### What SDG already has — do not rebuild these

All seven lists exist; all five custom permission levels exist (`CRS Upload`, `CRS Approve`,
`CRS Delete`, `CRS Share`, `CRS Request`); 839 groups; `CRS_SITE_MEMBERS` present; **Site Pages and
Site Assets both INHERIT** — so neither the home-page lockout of 2026-08-21 nor the hero-banner
permission problem of 2026-09-03 applies there.

`CRS Requests` and `CRS Submissions` are both at **0 items**, which is consistent with the request
and submission-record features never having been exercised on SDG.

**Config is COMPLETE** (measured 2026-09-10): `hcConfidentialityLevel` = `Highly Confidential` — the
row whose absence shows the HC level to **every** uploader — plus `legallyPrivilegedFor`,
`tenantDomains` = `sdguthrie.com`, `allowExternalSharing` = `no`, `uploadsPaused` = `no`, the three
`termSet_*` GUIDs, three mode rows (GHO, MHO, NBPOL), and `AllowedFileTypes` ticked
`.pdf .docx .xls .xlsx .doc`.

⚠⚠ **SDG'S TERM SET GUIDs ARE DIFFERENT AND MUST NEVER BE COPIED FROM THE TEST SITE.** The term
store is per-site: SDG is `fd973a45…` / `3a262838…` / `7901b1dd…` against the test site's
`866c5754…` / `023a866a…` / `0d6d1da8…`. **A below-Unit level's `Levels` JSON embeds a term set
GUID**, so copying a chain between sites points a level at a set that does not exist there — and
that fails as one permanently empty dropdown, which reads as *"this metadata was never captured"*
rather than as a misconfiguration.

### The five non-flow gaps on SDG, measured

1. ⚠⚠ **The archive libraries are missing ELEVEN wanted columns each**, five more than the test
   site's: additionally `Business_x0020_Segment`, `BusinessSegmentTid`, `SubUnit`, `SubUnitTid`,
   `Full_x0020_Name`. **The 2026-09-02 archive-column fix was applied to the test site and never to
   SDG.** A `MoveTo` carries only columns that EXIST at the destination, so anything archived there
   loses them silently — including `LegallyPrivileged`, a legal marker. ⚠ The three taxonomy columns
   among them cannot be created by code: a wrong binding looks right and silently tags nothing, so
   they are added by hand, as they were on the test site.
2. ⚠ **Only ONE user custom action** where there should be two — the `+ New Folder` customizer and
   the bulk-approve command set. Neither provisions from `elements.xml` on this tenant, so one of
   them is inert and will need POSTing by hand. Establish which.
3. **`autoApproveOwnUpload` is absent**, so a Head of Unit's own upload does not self-approve. The
   setting fails closed, so absent = off. **A client decision, not a defect** — the feature was asked
   for on 2026-08-24, so confirm whether it should be on before adding the row.
4. **`tenantDomains` is DUPLICATED** (items 11 and 17, same value). Harmless while they agree;
   undefined the moment somebody edits one. Delete one.
5. **`Restricted & Confidential Document` keeps 100 major versions on SDG, not 500.** Versioning is
   what makes Auto-route's Replace behaviour recoverable, so it is worth knowing the history there is
   shallower — 100 is still ample, and this is a note rather than a fault.

---

## 3. Library TITLES inside HTTP request URIs — ✅ MEASURED, AND THEY DO **NOT** CHANGE

⚠⚠ **THIS SECTION SAID THE OPPOSITE UNTIL IT WAS MEASURED, AND FOLLOWING IT WOULD HAVE BROKEN FOUR
FLOWS.** It claimed *"SDG's libraries were never renamed… SDG's are the original names"* and marked
five titles as **CHANGES**. **That was inferred from a CLAUDE.md line about which libraries EXISTED
on SDG in August, not about their current titles** — and the client renamed both sites. Editing
those five strings would have pointed `Auto-route`, `HC Auto Route` and both archive movers at
libraries that do not exist, failing as `NotFound`.

**Verified 2026-09-10 by running `scripts/dump-site-inventory.js` on both sites. The titles are
IDENTICAL:**

| URL segment | Title on BOTH sites |
|---|---|
| `/ApprovalDocument` | `Approval for Document` |
| `/HCApprovalDocument` | `Approval for Highly Confidential Document` |
| `/Shared Documents` | `Restricted & Confidential Document` |
| `/HCDocuments` | `Highly Confidential Document` |
| `/Archive` | `Archive Restricted & Confidential Document` |
| `/HCArchive` | `Archive Highly Confidential Document` |

**So every `getbytitle('…')` string stays exactly as it is** — `Restricted & Confidential Document`,
`Highly Confidential Document`, `Approval%20Document`, `CRS Submissions`, `CRS Group Map`. Leave
them alone. The one exception is the `Approval%20Document` inside **`HC Auto Route`**, which is wrong
on *both* sites for a different reason — §4.

⚠ **THE LESSON, since this is the failure this project keeps repeating: a claim about live state
inferred from a doc is not a measurement.** The inventory script exists precisely so this table
carries numbers rather than recollection. Re-run it rather than trusting this table if either site
is touched again.

---

## 4. Two bugs to fix WHILE porting, not after

- ⚠⚠ **`HC Auto Route` contains `getbytitle('Approval%20Document')` — the NORMAL library.** That is
  `Get_Reject_Comment`, and item ids are per-LIST, so on an HC document that id belongs to a
  different document or none. It is **inert today only because nothing reads its output**. Porting is
  the moment to either delete the action or repoint it to the HC approval library. **Do not carry the
  fault to a live tenant.** Full write-up in CLAUDE.md, 2026-09-10.
- ⚠ **`CRS — Audit replacements` has a TRAILING NEWLINE in a site address** (`…ClarenceDMSTesting\n`).
  Fifth occurrence of the invisible-newline trap in this project's flows. Retype the value rather
  than editing around it.

---

## 5. Per-flow checklist, heaviest first

Counts are references found in the definition. "hard" = occurrences of the literal
`ClarenceDMSTesting`, most of which sit in email bodies and HTTP URIs.

| Flow | list refs | titles in URIs | hard | Notes |
|---|---|---|---|---|
| `Auto-route` | 3 | 3 | **22** | Email bodies carry site URLs. Also the `Copy file` conflict setting — confirm it reads **Replace** |
| `HC Auto Route` | 4 | 3 | **22** | Same, **plus the `Get_Reject_Comment` bug above** |
| `CRS — Archive after seven years` | 2 | 1 | 11 | ⚠ Restore the rolling cutoff — see §6 |
| `CRS — HC Archive after seven years` | 2 | 1 | 11 | ⚠ Cutoff is ALREADY rolling in the export — but the 100 cap is not in effect. See §9 |
| `NotifyApprovers` | 1 | 1 | 7 | `ApprovalDocument.aspx` links in the body |
| `HCNotifyApprovers` | 1 | 1 | 7 | Same, and every link needs `&lib=hc` |
| `CRS — Approval reminder` | 1 | 1 | 6 | Approve/Reject links in the body |
| `CRS — HC approval reminder` | 1 | 1 | 6 | ⚠ **exists** — CLAUDE.md says it does not. Verify its Group Map filter is `APRHC`, not `APR` |
| `Audit — approval activity` | 3 | — | 5 | One ref is the **title** `Approval for Document`, not a GUID |
| `CRS — Notify request activity` | 1 | 1 | 4 | `SiteOrigin` + `RequestsPageLink` composes |
| `Audit — approval deletions` | 2 | — | 4 | |
| `Audit — HC approval activity` | 2 | — | 4 | |
| `Audit — HC approval deletions` | 2 | — | 4 | |
| `CRS — Audit replacements` | 2 | — | 4 | ⚠ trailing newline |
| `Audit — Documents deletions` | 2 | — | 3 | |
| `Audit — HC Documents deletions` | 2 | — | 3 | |
| `CRS — Audit request activity` | 2 | — | 3 | |
| `CRS — Approve new folders in Approval Document` | 1 | — | 2 | ⚠ trigger condition `{IsFolder} = true` |
| `CRS — Auto-approve bulk imports` | 1 | — | 2 | ⚠ needs `BulkImport` **visible** on the library |
| `HC auto-approve` | 1 | — | 2 | Same |
| `HC folder approval` | 1 | — | 2 | ⚠ trigger condition `{IsFolder} = true` |

⚠ **`HC Auto Route` also carries the title `Approval for Highly Confidential Document` as a `table`
value** — a connector action storing a title rather than a GUID. Easy to miss because it looks like
a list name in the designer and is a string in the definition.

---

## 6. Flow-specific traps, all previously paid for

- ⚠ **The two archive movers are currently built against a FIXED TEST CUTOFF**, not the rolling
  seven years. Their `Get items` filter reads `CutOff` — restore it to
  `addDays(utcNow(), -2557, 'yyyy-MM-ddTHH:mm:ssZ')` and re-enable the 100-file cap before they go
  anywhere near live data. A scheduled run against a stale literal archives by a fixed date, silently.
- ⚠ **`Get items` Pagination is a SEPARATE setting from `Top Count`.** Off, it silently caps the run.
- ⚠ **Trigger-condition polarity is load-bearing and both mistakes have been made:** `false` on the
  routing flows (files only), `true` on the folder-approval flows. Set it the wrong way and a flow
  never fires — leaving **no run history at all** to inspect.
- ⚠ **Type every numeric comparison through the `fx` editor.** A `0` typed into the value box is
  TEXT while `length()` returns an Integer; that mismatch failed all three conditions in the approver
  notification flow on 2026-08-25.
- ⚠ **Never press Enter in an fx box.** See §4 — it has now happened five times.

---

## 7. Order to work in

1. Run `scripts/dump-site-inventory.js` on **both** sites. Resolve all nine GUIDs and get SDG's real
   library titles. **Nothing else starts until this is done** — every step below depends on it.
2. Sign in as the **service account**.
3. Import the four **audit deletion** flows first — 2–4 references each, no email bodies. They are
   the cheapest way to prove the import-and-repoint loop works before touching a 22-reference flow.
4. Then the audit activity flows, then the folder-approval and auto-approve pairs.
5. Then the four notification flows (`NotifyApprovers` ×2, reminders ×2) — email bodies, so check the
   rendered mail rather than a green run: **a green `Send an email` says the send worked, never that
   the content is right.**
6. Then `CRS — Notify request activity`.
7. **`Auto-route` and `HC Auto Route` LAST.** They are the heaviest, they carry the approval emails,
   and they are the two whose failure loses documents rather than notifications.
8. Archive movers last of all, with the cutoff restored and **switched off** until deliberately
   scheduled.

⚠ **Leave every imported flow OFF until its references are checked**, and turn them on one at a
time. Twenty-one flows switched on together over a live library is not a state anyone can diagnose.

---

## 8. PORT PROGRESS, AND WHAT IS STILL UNVERIFIED (live — update as you go)

⚠ **A SAVED FLOW PROVES ITS REFERENCES PARSE. IT PROVES NOTHING ELSE.** Every flow below is
"imported and repointed" until a real event has produced a real row. This project has twice recorded
a flow reading **On** while every run since creation had failed in about two seconds.

### Imported and repointed

| Flow | Repointed | Trigger condition checked | Tested with a real event |
|---|---|---|---|
| `Audit — Documents deletions` | ✅ | ? | ✅ (proved the import loop) |
| `Audit — approval deletions` | ✅ | ⬜ | ⬜ |
| `Audit — HC approval deletions` | ✅ | ⬜ | ⬜ |
| `Audit — HC Documents deletions` | ✅ | ⬜ | ⬜ |
| `Audit — approval activity` | ✅ (5 fields — see below) | ⬜ | ⬜ |

### ⚠⚠ CLIENT'S OWN REQUEST, 2026-09-10 — DO NOT LET THIS SLIDE
> *"Later once all the audit flow is done remind me to show you this three flow to ensure it is
> working."* — the THREE deletion flows: `Audit — approval deletions`, `Audit — HC approval
> deletions`, `Audit — HC Documents deletions`.

**RAISE IT the moment the audit batch is finished.** What "working" has to mean for each, because a
green run satisfies none of these on its own:

1. **A row actually lands** in `CRS Audit Log` after a real deletion — `EventType: Deleted`.
2. **`Source` and `LibraryName` name the RIGHT vertical.** Every HC clone in this project has shipped
   with a library reference nobody swapped, so `Flow:ApprovalDeletionsHC` must be on the HC flow and
   `Flow:DocumentsDeletionsHC` on the approved-side one — a normal-library value on an HC flow is the
   fingerprint of a half-swapped clone.
3. **A deleted FOLDER produces NO row.** That is the `{IsFolder} = false` trigger condition, it is
   invisible in the designer, and its absence is silent — reconciliation, migrations and stray
   cleanup all delete folders in bulk, so one structure change would bury every real deletion.
4. **The path is recovered.** `DirName` comes from the recycle bin, so `ItemPath` should carry the
   unit folder. A `Deleted` row with no path cannot say WHOSE document was destroyed.
5. **`ItemUniqueId` is EMPTY, and that is correct** — the delete trigger returns none and the bin's
   `Id` is its own. Do not "fix" it. A deletion is the one event that cannot be threaded into
   *History of this file*.

### ⚠ The five fields `Audit — approval activity` needed, recorded because the next clone needs the same
Found live 2026-09-10 with three of five still on test-site values after the import:

1. trigger → List Name `eeb1bb19…`
2. **`GetUniqueId` → the GUID INSIDE the Uri** (`_api/web/lists(guid'…')`) — **nothing prompts for
   this one.** A Site Address change invalidates a *connector* action's List Name and forces a
   re-pick; a GUID inside an HTTP URI is just text and is carried across untouched and wrong.
3. `Get item` → **Site Address** as well as List Name. It was left on
   `dcidigitalcom…/ClarenceDMSTesting` while its List Name already held SDG's GUID — a pair that
   matches nothing. ⚠ The shape to fear is the near-miss: **a flow on SDG reading from our test site.**
4. `Already logged` → `4ac219e5…`
5. `Create item` → `4ac219e5…`

A GUID in `Get item` rather than the title `Approval for Document` is fine and slightly better — both
resolve for a connector action, and a GUID cannot be broken by a rename.

### ⚠⚠ PREREQUISITE BEFORE `CRS — Audit request activity` GOES ON
**SDG's `CRS Requests` is missing three columns** — the list sits at 0 items and nothing has ever
pressed *add missing columns* there. Open `Requests.aspx` on SDG as an admin and press it. Absent:
- **`SubmissionFileId`** — the identifier behind the 2026-09-10 fix. Without it a pending-stage
  request whose document routes before the approver decides fails **permanently**.
- **`RevokedBy`** — `EventKind`'s `Revoked` branch keys on it being non-empty, so absent means **a
  revocation logs as `RequestApproved`, naming the approver.** A false claim on the audit log.
- **`Stage`** — decides whether a request is against a pending or an approved document.

### PORT STATE, updated 2026-09-10 — 12 of 21 repointed, 2 tested with a real event

| Flow | Repointed | Trigger cond. | On | Tested |
|---|---|---|---|---|
| `Audit — Documents deletions` | ✅ | ? | ? | ✅ (proved the loop) |
| `Audit — approval deletions` | ✅ | ⬜ | ⬜ | ⬜ |
| `Audit — HC approval deletions` | ✅ | ⬜ | ⬜ | ⬜ |
| `Audit — HC Documents deletions` | ✅ | ⬜ | ⬜ | ⬜ |
| `Audit — approval activity` | ✅ | ⬜ | ⬜ | ⬜ |
| `Audit — HC approval activity` | ✅ | ⬜ | ⬜ | ⬜ |
| `Auto-route` | ✅ + conditions restored | ✅ | ⬜ | ⬜ |
| `HC Auto Route` | ✅ | ⬜ | ⬜ | ⬜ |
| `CRS — Approve new folders in Approval Document` | ✅ | ✅ `true` | ⬜ | ⬜ |
| `HC folder approval` | ✅ | ✅ `true` | ⬜ | ⬜ |
| `CRS — HC Archive after seven years` | ✅ (§9) | n/a | ⬜ | ⬜ (loop unexercised — see §9's last entries) |
| `CRS — Notify request activity` | ✅ | ✅ (`EventKind`/`WhichApproved` routed correctly) | ✅ | ✅ — deletion raised, approved, `DeleteApproved` branch fired, `SendEmailV2` returned 200 to the right address. **Delivery itself is blocked by the `simedarbyplantation.onmicrosoft.com` external-mail transport rule — see CLAUDE.md, 2026-09-10 — not a flow defect. Re-verify actual delivery once that rule has an exception, or by testing to an `@sdguthrie.com` address in the meantime.** |
| `CRS — Approval reminder` | ✅ (reported done 2026-09-11) | ⬜ | ⬜ | ⬜ — not yet cross-checked against its export; will be verified at the full batch export |

⚠ **`NotifyApprovers` AND `HCNotifyApprovers` SHOW IN SDG'S FLOW LIST, BUT ARE UNVERIFIED.** They were
built before this checklist existed and were not cross-checked against their exports field by field —
the client chose to keep building and export everything for one batch cross-check at the end rather
than confirm each one individually. **Do not read their presence in the "My flows" list as
"repointed" or "correct."**

### ✅ ALL 21 FLOWS NOW IMPORTED ON SDG (2026-09-11)
`CRS — HC approval reminder`, `CRS — Auto-approve bulk imports`, `HC auto-approve`,
`CRS — Audit request activity` and `CRS — Audit replacements` are all now confirmed present, on top
of the 16 already tracked above. **This closes the import phase.** Every flow got its site + list
GUID(s) repointed against the field-by-field checklists in this runbook.

⚠ **"IMPORTED AND REPOINTED" IS NOT THE SAME AS "VERIFIED."** Per §8's own opening warning, a saved
flow proves its references parse — nothing more. `NotifyApprovers`/`HCNotifyApprovers` were reported
done before this checklist existed and have never been cross-checked; every other flow's fields were
walked through here but **not one of the 21 has had its export re-pulled and diffed after the fact**
to confirm the edits actually landed as described. The planned full batch export + cross-check (client,
2026-09-10/11: *"I continue to build the rest then I export all for you to cross check"*) is what
closes that gap — do it before relying on any of this in production.

### ⚠⚠ `CRS — Audit replacements`'S TRIGGER WATCHED THE WRONG LIST — FOUND AND FIXED, 2026-09-11
First live test after import: replaced a file, nothing landed in `CRS Audit Log`, and the flow's run
history showed **zero runs at all** — not an error, nothing. Ruled out in order: the flow was ON, the
connection read `crs@sdguthrie.com` with no warning, `CRS Submissions`' own `ReplacedAt` column
genuinely held a value (confirmed from My Submissions showing `1 replaced`), so the trigger condition
was genuinely satisfied and still nothing fired.

- **THE TRIGGER'S `List Name` READ `CRS Requests`, NOT `CRS Submissions`.** So the trigger was polling
  an entirely different list — one a file replace never touches — which is why it produced no run at
  all rather than a failed one. Site Address was correct throughout
  (`https://sdguthrie.sharepoint.com/sites/CRS`); only the list picker was wrong.
- **LIKELY CAUSE: built immediately after `CRS — Audit request activity`**, whose trigger correctly
  watches `CRS Requests` — the two flows were done back to back and the list picker most likely
  defaulted to (or was left on) the previous flow's selection.
- **FIXED: changed the trigger's List Name to `CRS Submissions`.** Re-tested, verified live — two
  `Replaced` rows landed in `CRS Audit Log` with the correct filename and full path for each.
- ⚠ **WORTH CHECKING ON EVERY OTHER FLOW BUILT IN THE SAME SESSION, not just this one.** A wrong list
  picked from a dropdown produces NO visible symptom until something is actually tested against it —
  it saves clean, the connection shows valid, and the only tell is a real event producing zero runs.
  Any flow ported back-to-back with another one that shares a similar list name is worth a second look
  at its trigger specifically, not just its `Create item`/`Already_logged` actions.

### ⏭ WHAT IS STILL OPEN, NOW THAT IMPORT IS DONE
1. **The full batch export + cross-check**, covering all 21 — the one verification step that actually
   proves the field-by-field edits took, AND would have caught the wrong-trigger-list bug above without
   needing a live test to surface it.
2. **`CRS — Audit replacements`'s `Already_logged` filter — `EventType eq ''`** (found reading its
   export, 2026-09-11). Every row `Create_item` writes has `EventType = "Replaced"` hardcoded, so this
   filter can never match a real logged row and the dedupe never actually dedupes — a re-modification
   of a `CRS Submissions` row while `ReplacedAt` stays set would write a second audit row. **NOT
   exercised by the fix above** — that test replaced two DIFFERENT files once each, which never
   revisits `Already_logged` for the same item twice. Still open. **Decide whether to fix
   (`EventType eq 'Replaced'`) before or after testing that specific case** — flagged to the client,
   not yet decided.
3. **The normal archive mover's cutoff and cap**, same treatment as the HC one already got in §9 —
   restore the rolling `addDays(utcNow(), -2557, ...)` formula through the fx editor, confirm
   Pagination is off.
4. **Turn each flow ON one at a time and test with a real event**, per §7's closing instruction —
   twenty-one flows switched on together over a live library is not a state anyone can diagnose.
   §8's per-flow "Trigger cond. / On / Tested" columns are still mostly `⬜` and need working through.
5. **The client's own standing request, still not actioned**: *"Later once all the audit flow is
   done remind me to show you this three flow to ensure it is working"* — the three deletion audit
   flows (`Audit — approval deletions`, `Audit — HC approval deletions`,
   `Audit — HC Documents deletions`), with the five things "working" has to mean, listed at §8.

**Remaining flows still needing their site/list references confirmed CORRECT (imported, but not yet
walked through a checklist here):** `NotifyApprovers`, `HCNotifyApprovers`,
`CRS — Archive after seven years` (the normal one — cutoff/cap still to restore, per item 3 above).

### NEXT UP: `CRS — HC approval reminder`, field by field (measured from its export, 2026-09-10)

Already a correct HC clone on the test site — HC list GUID, `Role eq 'APRHC'` filter, all three links
already carry `&lib=hc`. Only the site+list repoint is needed, same shape as the normal one:

1. **`GetStale`** — `dataset` → SDG site, `table` → `d935aa6d-dc48-4832-ba7e-eca485ecdff3`
   (`Approval for Highly Confidential Document`).
2. **`GetApproverGroup`** — `dataset` only. `getbytitle('CRS Group Map')` unchanged; its filter is
   already `Role eq 'APRHC'` — leave it.
3. **`GetMembers`** — `dataset` only.
4. **`Send an email (V2)`** — body links already carry `&lib=hc`; swap only the host/site portion to
   `https://sdguthrie.sharepoint.com/sites/CRS/SitePages/ApprovalDocument.aspx`, keeping `&lib=hc`,
   `&decision=approve`/`reject` exactly as they are.

Same 1 GUID + 1 unchanged title + 3 site refs + 3 body links = 6 hard references, matching the
runbook's count. Same trailing-whitespace note as its sibling flow — leave it alone.

### NEXT UP: `NotifyApprovers`, field by field (measured from its export, 2026-09-10)

Matches the runbook's "1 list ref / 1 title in URI / 7 hard" count for this flow exactly.

1. **Trigger `When_an_item_is_created_or_modified`** — `dataset` → SDG site, `table` →
   `eeb1bb19-ec53-4c41-8dc9-ecf255979c9b`. ⚠ **Check its three trigger conditions survive the
   import**: `{IsFolder} = false`, `BulkImport ≠ true`, `{ModerationStatus} = 'Pending'`.
2. **`GetDoc`** — same `dataset` + `table` pair as the trigger.
3. **`GetApproverGroup`** (inside `HasUnit`) — `dataset` only; its `parameters/uri` reads
   `getbytitle('CRS Group Map')`, a TITLE, which per §3 stays unchanged on both sites.
4. **`GetMembers`** (inside `HasApprover`) — `dataset` only; its URI is `sitegroups(...)`, web-relative,
   no list name involved.
5. **`Send_an_email_(V2)`** (inside `HasRecipients`) — **no `dataset`/`table`**, but its `Body` carries
   **three** hardcoded links to
   `https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting/SitePages/ApprovalDocument.aspx` (the
   "View Document", "Approve", "Reject" links) — all three need the SDG host AND site path:
   `https://sdguthrie.sharepoint.com/sites/CRS/SitePages/ApprovalDocument.aspx`. **Verify the real
   page name on SDG first** — this client renames pages.

That accounts for all 4 `dataset` occurrences + 3 body links = the 7 hard references. `CRS Group Map`
(a title) is the "1 title in URI" and is left alone.

⚠ **NOT A PORTING CONCERN, BUT NOTICED READING THE EXPORT: `HasRecipients`'s condition carries a
trailing `\r\n` inside its left operand** — `"@length(coalesce(body('PickEmails'), createArray()))\r\n"`
compared against `"@0"`. This is the SAME invisible-newline shape flagged repeatedly elsewhere in this
project's flows. **Do not "fix" it while porting** — this exact expression is what this flow has run
successfully with for weeks on the test site (verified: `HasRecipients` gating an empty group has been
tested and works), so whatever Power Automate does with the trailing whitespace here is already the
proven, working behaviour. Importing carries it across unchanged, which is correct. Flag it only if a
populated approver group on SDG somehow fails to receive an email despite `PickEmails` returning rows.

`HCNotifyApprovers` is the same shape with `getbytitle('CRS Group Map')`'s filter on `Role eq 'APRHC'`
instead of `'APR'`, and every link needs `&lib=hc` appended — confirm both before importing it next.

### ✅ MEASURED 2026-09-10 — the column prerequisite is CLEARED for the routing flows
`ApprovedBy`, `SubmissionId`, `BatchId`, `SubmissionFileId` are ALL present, all `SP.FieldText`, on
**both** `Approval for Document` (`eeb1bb19…`) and `Approval for Highly Confidential Document`
(`d935aa6d…`). So the worry that SDG's reconciliation ran only half way does **not** affect
`Auto-route` / `HC Auto Route`. ⚠ **`CRS Requests` is a DIFFERENT list and is still unchecked** — it
needs *add missing columns* pressed on SDG's `Requests.aspx` before `CRS — Audit request activity`.

### ⚠⚠ THE NEW DESIGNER EMPTIES CONDITION LEFT-OPERANDS AND THE SAVE PERSISTS IT (2026-09-10)
`Auto-route` was opened in the **new** designer and saved. **Three conditions lost their left operand
and the damage was written to the server:**

| Condition | After the save | Restored to |
|---|---|---|
| `Condition` | `"" equals "Approved"` | `body('Get_item')?['{ModerationStatus}']` |
| `Condition_1` | `"" equals "Denied"` | `body('Get_item')?['{ModerationStatus}']` |
| `Condition_2` row 1 | `not(equals("", true))` | `body('Get_item')?['BulkImport']` |

- **WHICH ONES BREAK IS PREDICTABLE: the DYNAMIC-CONTENT references.** `DestExist` and `WasReplaced`
  survived untouched — they are hand-typed `@length(coalesce(...))` expressions. All three that broke
  were connector-field tokens, two of them field names carrying **braces** (`{ModerationStatus}`),
  which the new renderer cannot draw.
- **⚠ THE FAILURE WAS SAFE ONLY BY LUCK OF SHAPE.** `"" equals "Approved"` is always false, so the
  flow took the else branch, where `"" equals "Denied"` is *also* always false — **the flow became
  inert rather than routing unapproved documents.** Do not rely on that: a condition emptied the
  other way round would have been an approval bypass.
- **`HC Auto Route` was checked action-by-action in the CLASSIC designer and all three conditions
  survived.** So this is avoidable: **stay in the classic designer for any flow with conditions**, and
  read Code view BEFORE saving, not after.
- **THE RECOVERY IS THE EXPORTS.** `PowerAutomateFlows/*.zip` → `Microsoft.Flow/flows/*/definition.json`
  carries every original expression. That is what the exports are FOR, and it is the only reason this
  was a twenty-minute repair rather than a hand-rebuilt approval condition.

### ⚠⚠ TWO PRE-EXISTING BUGS THE EXPORTS EXPOSED — both live on the TEST site too, both deferred
Client, 2026-09-10: *"lets keep building first, we can fix it later."*

1. **THE 8-HOUR `Created` SKEW FIX WAS NEVER APPLIED TO `Auto-route`.** CLAUDE.md records it as applied
   to BOTH routing flows on 2026-09-05. The two exports, taken four minutes apart on 2026-09-09,
   disagree:
   - `HC Auto Route`: `formatDateTime(convertFromUtc(body('Get_item')?['Created'], 'Singapore Standard Time'), 'M/d/yyyy h:mm tt')` ✅
   - `Auto-route`: `formatDateTime(body('Get_item')?['Created'],'M/d/yyyy h:mm tt')` ❌
   So **every document routed through the NORMAL library is stamped 8 hours early**, and anything
   approved before 08:00 local lands on the previous day. `Created` is what My Submissions sorts by
   and what CRS Search displays. **Fix on SDG and on the test site.**
   - ⚠⚠ **ON SDG THE HC FLOW — WHICH DOES CONVERT — IS 15 HOURS AHEAD (measured 2026-09-10).** Five
     files routed at `07:05Z` read `Created 22:04Z`. Site Regional Settings are correctly
     `(UTC+08:00) Kuala Lumpur, Singapore` (bias −480), so the site is not the cause.
     **Working theory, not yet confirmed:** `validateUpdateListItem` parses the date string in the
     CALLING USER's regional settings, and the service account `crs@sdguthrie.com` has its own time
     zone set to Pacific (UTC−7 in September). `3:04 PM` read as PDT is exactly `22:04Z`, and
     SGT − PDT is exactly 15 h. On the test site the flows run as a person whose settings follow the
     site, which is why the same expression was correct there.
     - If right, the NORMAL `Auto-route` on SDG (no `convertFromUtc`) is off by **+7 h**, not −8 h.
     - Fix at the account, not the flow: as `crs@sdguthrie.com`, open
       `/sites/CRS/_layouts/15/regionalsetng.aspx?Type=User` and make it follow the site's time zone.
       Then add `convertFromUtc` to `Auto-route` as §8.1 already says.
     - Nothing backfills: files already routed keep their wrong `Created`.
2. **⚠ `Get_Reject_Comment` IS NOT INERT — IT KILLS THE REJECTION EMAIL.** CLAUDE.md says it is
   harmless "because nothing reads its output". Its output is not read, but
   `Send_an_email_(V2)` carries `runAfter: { Get_Reject_Comment: ["Succeeded"] }` — so when it 404s the
   email never sends and the run reports Failed. It queries `getbytitle('Approval%20Document')`, **a
   title that stopped existing when the libraries were renamed on 2026-08-28**, in BOTH routing flows.
   **So rejection emails have been silently dead on the test site for two weeks.**
   - Fixed on SDG's `HC Auto Route` by retitling to `Approval for Highly Confidential Document`
     (verified: that title resolves to `d935aa6d…`, `ItemCount` 190).
   - **Still to do:** `Auto-route` → `Approval for Document`, on SDG and on the test site. The cleaner
     fix is to delete the action and clear the email's `runAfter`, since nothing reads its output.
   - **THE LESSON: "nothing reads its output" is not the same as "nothing depends on it."** A
     `runAfter: Succeeded` is a dependency that consumes no data.

---

## 9. `CRS — HC Archive after seven years` — measured from the export, 2026-09-10

Read out of `CRS—HCArchiveaftersevenyears_20260909181215.zip`, not from §6.

- ✅ **`CutOff` is ALREADY `addDays(utcNow(), -2557, 'yyyy-MM-ddTHH:mm:ssZ')`** and `Get_items` reads
  it through `trim()`. §6's "restore the rolling cutoff" does not apply to this flow's export.
- ⚠⚠ **THE 100-FILE CAP IS NOT IN EFFECT.** `Top Count` is 100, but Pagination is ON with a threshold
  of 5000 (left over from the 2026-09-02 bulk test). With pagination on, Top Count is the page size
  and the threshold is the ceiling — so one run can move 5000. **Turn Pagination OFF** to restore the
  cap.
- **Nine actions carry the test site** and all need `https://sdguthrie.sharepoint.com/sites/CRS`.
  Four more values change:

| Action | Field | Test site | SDG |
|---|---|---|---|
| `Get_items` | List Name | `Highly Confidential Document` (a title) | `122a4aa9-65fb-479a-90a7-3866d19f51b4` |
| `Create_new_folder` | List Name | `bfcd6735-…` | `fc301054-abf5-4e53-957a-6f8c58a410b8` |
| `Send_an_HTTP_request_to_SharePoint` (the `moveto`) | Uri, **twice** | `/sites/ClarenceDMSTesting/` | `/sites/CRS/` |
| `GetArchivedFileId` | Uri, `lists(guid'…')` | `bfcd6735-…` | `fc301054-abf5-4e53-957a-6f8c58a410b8` |

- ⚠ **The last two are text inside HTTP URIs, and nothing prompts for them.** A Site Address change
  makes a connector action's List Name go blank; it does nothing to these.
- **Unchanged, deliberately:** `RelPath`'s `'HCDocuments/'`, `Find_the_moved_item`'s `'/HCArchive/'`,
  the two `getbytitle('Archive Highly Confidential Document')` URIs and both `getbytitle('CRS
  Submissions')` URIs. URL segments and titles are identical on both sites (§3), and `CRS Submissions`
  is SHARED by both verticals and must not be repointed.
- **Before switching it on, confirm nothing on SDG is already older than the cutoff.** A migration
  tool that preserves original `Created` dates would make the first run archive real documents at
  once — the "nothing is seven years old until 2033" argument only holds for documents uploaded
  through CRS.
- ✅ **PRE-CHECK PASSED ON SDG, 2026-09-10** (cutoff `2019-09-10T08:05:38Z`): **zero** HC documents
  older than the cutoff; `ArchivedAt` (DateTime) present on `CRS Submissions`; `Archived` (Boolean)
  and `SubmissionFileId` (Text) present on `Archive Highly Confidential Document`. So the first run
  moves nothing, and every column the flow writes exists.
- ✅ **IMPORTED AND REPOINTED ON SDG, 2026-09-10.** Verified from Code view: the `moveto` Uri reads
  `/sites/CRS/` both times, `GetArchivedFileId` reads `fc301054-…`, `HasStamp` survived the new
  designer intact, and `StampArchivedRecord` is a POST/MERGE over a correctly bound `StampLoop`. A
  manual run went green with `Get_items` returning nothing — which also proves the SDG site and list
  GUID resolve. **The loop body has NOT run on SDG and cannot until a document is seven years old**;
  it was proven on the test site. ✅ Pagination OFF — `Get_items`' Code view carries no
  `runtimeConfiguration` block, so `$top: 100` is the real cap.
- ⚠⚠ **A CORRECT URI ON A TEST-SITE SITE ADDRESS STILL CALLS THE TEST SITE — found on the first
  forced run, 2026-09-10.** The `moveto` Uri read `/sites/CRS/…` both times and was checked, but that
  action's **Site Address** was still `dcidigitalcom…/ClarenceDMSTesting`. The request went to OUR
  tenant carrying SDG's paths, and `crs@sdguthrie.com` answered **401** with an empty body (`{}`) —
  44 of 44. Nothing moved: `Create_new_folder` ran first and succeeded, so SDG's archive holds empty
  folders only. **An HTTP action has two independent places a site lives — the Site Address
  dropdown and the Uri — and checking one says nothing about the other.** The tell is the `source`
  field in the raw OUTPUT, which names the host actually called.
- ⚠ **`CutOff` WAS RESTORED AS PLAIN TEXT on the first attempt** — typed into the Compose box, so its
  output was the literal string `addDays(...)` and `Get_items` failed `Creating query failed`. It
  must go in through the **fx** editor; Code view then shows a leading `@`. Re-entered 2026-09-10;
  the next run went green with `Get_items` returning nothing, as it should for a seven-year cutoff.
  ⚠ The `moveto` Site Address fix is therefore **still unexercised** — the loop has not run since.
- ⚠ **The new designer has no whole-flow Code view** — its search box covers ONE action. So "search
  the flow for `ClarenceDMSTesting`" has to be done per action, or by re-exporting the flow and
  grepping `definition.json`.

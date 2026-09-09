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
| `CRS — HC Archive after seven years` | 2 | 1 | 11 | Same |
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

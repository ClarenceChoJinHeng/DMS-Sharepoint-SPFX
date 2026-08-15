# Highly Confidential — a separate library pair

**Status:** design, agreed 2026-08-15. Supersedes `2026-07-16-highly-confidential-securing-design.md`.

---

## 1. What the client asked for

> "there will be an uploader group that is for highly confidential, so if the group is highly
> confidential then they can also upload not just to Highly confidential but also Confidential and
> Restricted. HIghly confidential Group can also go to normal Approval library, also under the same
> unit and group, for example: `GHO_GF_CORU_UPL` — only access to approval library;
> `GHO_GF_CORU_UPL_HIGHLY_CONFIDENTIAL` — access to both approval library and highly confidential
> library under the same unit."

And, on who may read an approved HC document (2026-08-15):

> "the same HOD and the same C level segment and global can read, no need a separate HC for them,
> only for Approver HOU and PIC uploader needs a separate group."

---

## 2. Why the July design is superseded

`2026-07-16-highly-confidential-securing-design.md` kept HC files in the **same** library, in a
`…/Document Type/Highly Confidential/` subfolder, secured by breaking permissions on that folder.

That design needed an elevated, HTTP-triggered Power Automate flow running as a service account with
Full Control, because:

- an uploader holds Contribute, not Manage Permissions, so the web part could not secure the folder;
- an auto-created folder inherits the *visible* unit ACL, so there was a window — however brief —
  in which a peer could see the file;
- once secured, the ordinary uploader could not write into it either.

It also embedded the flow's trigger URL — a bearer secret — in the SPFx bundle.

**A separate library removes the entire problem rather than solving it.** Ordinary uploaders hold no
permission on the HC library at any moment, so there is no window to close, no elevation to perform
and no secret to ship. Every mechanism in that spec exists to compensate for HC files living
somewhere their peers can already reach; none of it is needed once they do not.

Nothing from the July spec is retained. It stays in the repository as the record of a rejected
approach, so the elevated-flow idea is not re-derived from scratch.

---

## 3. The model

### 3.1 Confidentiality becomes routing — for one level only

Today the confidentiality level is a pure label: *"Confidentiality is metadata, not a permission. Any
PIC may upload at any confidentiality level."* That stays true for every level **except** the one
named by a new config row.

A `DMS Config` setting row, `hcConfidentialityLevel`, holds the level label that routes to the HC
pair. Default `Highly Confidential`.

**A row rather than a constant, and a label rather than a term GUID.** The levels come from the
`confidentiality` term set, which the client owns and renames; a constant would need a redeploy, and a
GUID cannot be typed by an administrator who is looking at a label. Compared case-insensitively and
trimmed, because that is how the label will be re-typed.

**One level, not a general map.** A `level → library` map was considered and rejected as unearned: it
is one config row either way, and a map invites a second HC-like library that nothing else in this
design supports — separate groups, separate flows, separate columns. When *"Confidential should be
separate too"* is actually asked for, this row becomes a map and the rest of §4 is repeated for it.
Until then, a single row cannot be half-configured.

### 3.2 Four libraries, two pairs

| Pair | Approval side | Approved side |
|---|---|---|
| Normal | `Approval Document` | `Documents` |
| Highly Confidential | `HC Approval Document` | `HC Documents` |

Same folder tree, same abbreviations, same content approval, same draft-item security, same columns.

**Created without spaces and then retitled**, as `Approval Document` was — a list's URL is fixed at
creation and never moves on rename (gotcha #12). So `HCApprovalDocument` / `HCDocuments` are the URL
segments, and the titles carry the spaces.

**Both halves are resolved at runtime, never hardcoded.** The client renames everything at import
(memory `dms-to-crs-rename-pending`) and has already said these names may change. `primeNames` gains
the HC pair, probing `HC Approval Document` → `HCApprovalDocument`, and the title/URL pair is resolved
together from the response that carried both — never one derived from the other. A wrong title 404s
loudly; a wrong URL segment fails **silently**, and that is the failure that matters.

### 3.3 `LibTarget` gains two members

`type LibTarget = "Staging" | "Documents" | "StagingHC" | "DocumentsHC"`.

`Staging` remains a **logical key** — it is the stored `Target` value on library-scope Group Map rows
and must keep being written and filtered on unchanged. `StagingHC` and `DocumentsHC` are logical keys
in exactly the same way, translated to real titles at the API boundary by `libApiTitle()`.

`libApiTitle` currently special-cases one key; it becomes a lookup over all four. **The translation
stays at the API boundary**, never in stored data — the lesson of `StagingAccess`, which was broken
for precisely this reason on 2026-08-14.

---

## 4. Groups, roles and who sees what

### 4.1 Two new roles

`UPLHC` and `APRHC`. Group-name suffixes: `_UPL_HC` and `_APR_HC` canonically; the parser also accepts
`_UPL_HIGHLY_CONFIDENTIAL` and `_APR_HIGHLY_CONFIDENTIAL`, because that is the form the client wrote
and an administrator will type it.

**The name builder suggests the LONG form** — `GHO_GF_CORU_UPL_HIGHLY_CONFIDENTIAL`, exactly as the
client wrote it. An earlier draft of this spec said the short one, which was wrong on both counts:
`suffixForRole` already emits the long spelling for every other role (`_UPLOADER`,
`_DELETER_DOCUMENTS`), so the short form would have made HC the one inconsistent case, and it is not
the form the client used.

**Precedence between the spellings is handled by the existing length sort, not by declaration order.**
`_UPL_HC` (7 characters) is tested before `_UPL` (4) and before `_HC` (3), so a group named
`…_UPL_HC` can never parse as a plain uploader. Worth stating because getting it wrong grants HC
clearance to nobody while the Group Map still reads entirely correctly.

**Both roles are supersets.** An HC uploader files at *any* level; an HC approver approves in *either*
library. The separation the client asked for runs the other way: plain `UPL` and `APR` reach the HC
libraries **not at all**.

This is the client's own rule for uploaders, stated verbatim in §1, applied symmetrically to
approvers. The alternative — an HC approver who can approve only HC — needs two groups for the common
case of one person doing both, and gains nothing: *"not every approver sees HC"* is delivered entirely
by plain `APR` having no HC grant.

### 4.2 `LIBRARY_ROLES`

```
Staging:      UPL, APR, DELS, UPLHC, APRHC
Documents:    MEMBER, DEL, GLOBAL, SEGVIEW, UPL, APR, SHARE, UPLHC, APRHC
StagingHC:    UPLHC, APRHC, DELS
DocumentsHC:  DEL, GLOBAL, SEGVIEW, SHARE, UPLHC, APRHC
```

Read against the existing table, three absences are load-bearing:

- **`UPL` and `APR` are absent from both HC rows.** That single omission is the whole feature. This
  table exists precisely so a new role reaches no library until someone names it here, and the same
  property protects HC.
- **`MEMBER` is absent from `DocumentsHC`.** An SDG Employee reads their unit's approved documents;
  they are not HC-cleared, and nothing in §1 grants them HC.
- **`SEGVIEW` and `GLOBAL` are absent from both approval-side rows**, HC included. A segment-wide
  viewer on an approval library reads every unapproved draft in that segment — the existing isolation
  rule, unchanged and now doubly important.

`DOCUMENTS_READ_ONLY_ROLES` extends to `DocumentsHC`: `UPLHC` and `APRHC` downgrade to `Read` there,
exactly as `UPL`/`APR` do on `Documents`. Without it the flat `ROLE_TO_PERMISSION` grants `CRS Upload`
on the approved HC archive.

### 4.3 Personas

Two new personas, `pic_hc` and `hou_hc`, mirroring `pic` and `hou`:

| Persona | Roles |
|---|---|
| PIC — Highly Confidential | `UPLHC`, `DELS` |
| Head of Unit — Highly Confidential | `APRHC`, `UPLHC`, `DELS`, `DEL`, `SHARE` |

`personaTouchesStaging()` must classify both under the approval-library half of the Folder Access
picker — derived as it is today, never hand-listed.

### 4.4 What HoD and C-Level inherit, stated rather than assumed

The client's instruction is that HoD and both C-Level roles use their **existing** groups on the HC
pair. Their existing roles carry their existing powers:

- **`DEL` is view *and delete*.** A Head of Department can delete an approved HC document.
- **`SHARE` is Manage Permissions** at that scope. A Head of Department or C-Level can share an
  approved HC document, with no request and no approval.

Both follow directly from *"the same HOD and the same C level can read"*. Neither is a side effect of
the implementation, and both belong in front of the client before deployment rather than after.

---

## 5. The safety rule everything hangs on

**If the HC library cannot be resolved, the Highly Confidential level is not offered at all.**

Not offered-and-then-failing. Not falling back to the normal library. Absent from the dropdown.

This deliberately breaks the codebase's fail-open habit, for the same reason `canOfferFolderDelete`
and `allowExternalSharing` do. Elsewhere the cost of a failed read is a form out of service. Here,
falling back files an HC document into the library **every PIC in the unit reads once it is
approved** — silently, under a green success toast, with nothing on screen or in any log to say the
document went somewhere other than intended.

### 5.1 The dropdown is filtered by a WRITE probe, not by group membership

`AddListItems` against the user's resolved HC folder, exactly as `filterReachablePaths` does for
segments today. Holding `GHO_GF_CORU_UPL_HC` and being able to write into that unit's HC folder are
different questions: the group can exist before reconciliation has granted anything, which is the
origin of the upload form's original HTTP 403.

Read `Low` **arithmetically, never with `&`** — JS bitwise coerces to a signed 32-bit int, and Full
Control returns `Low = "4294967295"`.

An **inconclusive** probe hides the level. That is the opposite of the segment-visibility rule, and
the difference is the cost of being wrong: an over-hidden segment blocks an upload someone retries,
an over-offered HC level publishes a secret.

### 5.2 The upload re-checks

The destination is decided at submit time from the level, and the write itself is the final authority.
A user whose clearance was removed between page load and upload gets a refusal, not a misfile — the
same stale-page reasoning as gotcha 10b, and the reason `chainSignature` exists.

---

## 6. What changes, by area

| Area | Change |
|---|---|
| `shared/naming.ts` | HC pair in `primeNames`; `libApiTitle` over four keys |
| `shared/groupMapModel.ts` | `UPLHC`/`APRHC` roles, two personas, suffix parsing for both spellings |
| `FolderManager.tsx` | `LibTarget` ×4, `LIBRARY_ROLES`, `DOCUMENTS_READ_ONLY_ROLES`, reconciliation over four libraries |
| `Form.tsx`, `BulkUpload.tsx` | Filter the level list by probe; route to the HC pair |
| `ApprovalDocument.tsx` | Read the HC approval library for HC approvers |
| `MySubmissions.tsx` | Four libraries instead of two |
| `SegmentCreator.tsx`, `StructureManager.tsx` | Tier columns created in **four** libraries |
| `Requests.tsx` | HC deletions and shares decided by the HC approver |
| `shared/auditLog.ts` | No new event types — `library` already names which library was acted on |

### 6.1 Reconciliation roughly doubles

It already walks the term tree per segment and grants per folder; it now does so for four libraries.
The existing progress UI and the one-row-per-run audit rule both hold. Worth measuring before the
client runs it across twelve segments.

### 6.2 Columns must exist in four libraries

The standing warning applies with twice the surface: **`Documents` MUST carry the same columns under
the same internal names**, and for months it silently did not. Two silent failures came from that one
gap — bulk upload tagging nothing, and auto-route dropping metadata on every approved document.

`ensureColumn` already lives in `shared/spColumns.ts` and is shared by SegmentCreator and
StructureManager; it gains the HC pair. Provisioning must diff
`/fields?$select=Title,InternalName` across **all four**.

---

## 7. Power Automate — two new flows, and nothing works without them

The existing Auto-route and folder-approval flows are bound to the normal approval library. They will
never see an HC file.

1. **HC Auto-route** — as Auto-route, split on the HC URL segment, copying into `HC Documents`.
   Trigger condition `@equals(triggerOutputs()?['body/{IsFolder}'], false)`.
2. **HC folder approval** — as the existing one, trigger condition `…{IsFolder}'], true)`.

**The polarity is the whole thing**, and both mistakes have already been made once. `true` on
auto-route means no file is ever routed, silently, because a flow that never fires leaves no run
history.

Both must be created while signed in as the **service account**: a flow runs under its connection, the
connection is created implicitly by the first action, and the builder's account is then baked in for
life.

`HC Documents` must have **content approval OFF**, verified per site — the same trap as `Documents`,
presenting as a permissions bug that is not one.

---

## 8. Migration — documents already labelled Highly Confidential

Any file already carrying that label sits in the normal libraries today and is readable by its whole
unit. Deploying this does not move them, and afterwards the label implies a protection those files do
not have — which is worse than before, because the label now looks like it means something.

A one-off sweep is required: find items whose `Confidentiality Level` is the HC label, in both normal
libraries, and move them to the HC pair. The subtree migrator's verified facts apply — a file move
preserves approval status, and a folder move carries its subtree and UniqueId — so destination folders
can be ensure-created and stamped Approved where the library moderates.

**Out of scope for this build, but it must not be silently skipped.** Until it runs, the honest
statement to the client is that HC protection applies to documents filed *after* deployment.

---

## 9. Out of scope, deliberately

- **A general `level → library` map** — §3.1.
- **A separate HC term set, or an HC-specific folder structure.** Same tree, same abbreviations — the
  client's own decision on 2026-08-11: *"different library but same structure"*.
- **HC-specific personas for HoD and C-Level.** Explicitly declined by the client — §1.
- **Per-uploader isolation inside an HC unit folder.** Unchanged and still impossible in an approved
  library; the unit folder remains the smallest confidentiality boundary, HC included. Two HC-cleared
  people in one unit see each other's approved HC documents.

---

## 10. Open questions for the client

1. **Is a Head of Department deleting an approved HC document intended?** (§4.4)
2. **Is a Head of Department or C-Level sharing an approved HC document intended** — with no request
   and no approval? (§4.4)
3. **When should the migration sweep run**, and against which sites? (§8)
4. **Does every unit get HC groups, or only the units that handle HC?** A unit with no HC groups
   never offers the level, which is the safe default.

---

## 11. Test plan

1. A plain PIC does not see `Highly Confidential` in the dropdown.
2. An HC PIC does, and also files a `Confidential` document into the **normal** library.
3. An HC PIC files an HC document; it lands in `HC Approval Document`, pending.
4. A plain PIC of the same unit sees nothing of it, in either library.
5. A plain HoU (`APR`, no HC) cannot see it either — the case that distinguishes this from a
   folder-level scheme.
6. The HC approver approves it; HC Auto-route moves it to `HC Documents`.
7. HoD, C-Level segment and C-Level global each read it there. `MEMBER` does not.
8. My Submissions shows the HC uploader their HC file across both HC libraries.
9. Reconciliation grants all four libraries in one run and reports counts per library.
10. With the HC library renamed, everything still resolves — nothing hardcoded.
11. With the HC library **deleted**, the level disappears from the dropdown and no upload routes to
    the normal library.
12. Delete and share requests against an HC document reach the HC approver, not the plain one.

# Access requests — the Head of Unit approves, not the site owner

**Date:** 2026-08-19
**Status:** design agreed, NOT implemented
**Extends** `2026-08-15-deletion-and-share-requests-design.md` — the request list, the routing and the
approve-in-the-approver's-session mechanism all come from there, unchanged.
**Depends on** `2026-08-19-derived-page-access-design.md`: the page grants below are derived from
roles rather than granted by hand.
**Does not reopen** `2026-07-23-share-guard-retirement.md`. SharePoint's own Share button still cannot
be intercepted, and nothing here tries.

---

## 1. Why this exists

The client's instruction, 2026-08-19: access to a document in `Documents` should be approved **by the
Head of Unit, by group** — not by a site owner, and not by email.

Turning **Access Requests** off (added to the migration runbook the same day) removes the wrong route.
A PIC holds only Read on `Documents` — the `UPL` downgrade on approved-side libraries — so pressing
SharePoint's Share button falls through to an access request emailed to the site owner. With the
setting off, that press is inert.

**But inert is not answered.** Someone who genuinely needs a document then has no route at all, and
will ask a colleague to email them a copy — which is worse than what we just switched off, because it
leaves no record anywhere. This is the replacement route.

## 2. The shape

**A third `RequestType`, in which the recipient IS the requester.**

That one property is what makes the design safe, and it is worth stating plainly: **nothing is typed.**
The existing share flow asks the requester to type recipient addresses, and a Microsoft 365 group or
distribution list has an address — `finance-all@…` is a valid `Key` for `SP.Web.ShareObject` — so one
approval can reach two hundred people while rendering on the approver's screen as a single recipient
(§7, an open finding on that flow). Here the recipient is resolved from the signed-in session, so that
class of mistake cannot occur at all.

Scope: **`Documents` and `HC Documents` only.** There is no sharing in the approval libraries. `SHARE`
is absent from their role tables, confirmed by the 2026-08-19 run in which `→ CRS Share` appears under
`Documents` and `HC Documents` and never under `Approval Document`. A document there is a draft: there
is nothing to share yet, and Draft Item Security means a peer cannot see it anyway.

## 3. The flow

### 3.1 Raising it

The requester hits a denial page — SharePoint's, which still carries the item id — and pastes that link
into **`Request-Access.aspx`**.

The page resolves the link **as them**, which is the safe direction: without permission, the only thing
readable is the path. From the path it derives segment → department → unit:

```
A document in  Group Head Office → Group Finance → Tax
File name:     not shown — you do not have access to it yet
Goes to:       the Head of Unit for Tax
Reason:        ______
```

**The file NAME is withheld until access is granted.** A name can be confidential by itself —
`Redundancy list Nov 2026.xlsx` discloses plenty without opening anything. The approver sees it; the
requester does not.

One row is written: requester from the session, item, unit, reason, `Status = Pending`.

### 3.2 Routing

Unchanged from 2026-08-15. `documentUnit()` derives the tier chain from the document's own `<Base>Tid`
fields; `routeToApprover()` takes the **deepest tier that actually has an `APR` mapping** — never
simply the deepest, because a below-Unit tier such as SubUnit carries a Tid column exactly like a
permissioned one, so routing on depth alone files the request where nobody can see it.

When nothing matches, the request is **still sent** and the requester told an approver mapping is
missing. An unreadable Group Map and a genuinely absent approver are indistinguishable from here, so it
fails towards being visible.

### 3.3 The approver's decision

```
Chin Wei Ling (chinweiling@…) requests access to
Tax Computation FY2026.xlsx — Group Finance → Tax
"Working on the group consolidation with Finance Ops"

⚠ This person is not in your unit.
  ○ Grant them this one document        (view only / can edit)
  ○ Add them to GHO_GF_TAX_EMPLOYEE     — they would see everything approved in Tax
  ○ Reject, with a reason
```

**The middle option performs the add** (client, 2026-08-19: *"Yes add them to the unit group, should be
easier"*). `addMemberWithSiteEntry` exists and already handles the site-entry group, so the mechanics
are done.

- **The group is the unit's `_EMPLOYEE` group** — view-only. Never the uploader or approver group,
  which is what "add them to the unit" could be misread as granting.
- **The confirmation states the real scope**, because it is far wider than the request:
  > Adding **Chin Wei Ling** to `GHO_GF_TAX_EMPLOYEE` gives them **every approved document in Tax**,
  > not only this one — in `Documents` now, and in any library this group is later mapped to. It does
  > **not** give them Highly Confidential access.
- **It is usually the right answer, and the screen should say so.** A one-file grant creates a
  permanent unique permission scope — the count that walks towards SharePoint's 50,000 ceiling, which
  is what killed per-uploader ACLs on 2026-08-06 — and someone who needs Tax's documents once will need
  them again next month and raise another request. Group membership costs no scope at all.
- **A requester from another segment is surfaced, not blocked**: *"this person is in Group
  Sustainability"*. Legitimate, and the approver should see it before deciding.

### 3.4 Approval

Executes **in the approver's own browser session, as them** — unchanged from 2026-08-15 and the reason
it is acceptable at all: the grant can never exceed the approver's own rights, it fails loudly rather
than over-granting silently, and the audit row names who actually did it. It completes only while they
are on the page.

**A failed grant records `Failed`, never `Approved`.** A row reading "access granted" when the grant did
not apply is how someone stops looking.

### 3.5 Afterwards

- A one-file grant appears on the live-grants list with **Revoke** (§6).
- A group add has nothing to revoke and nothing to expire — removing them from the group on Folder
  Access ends it immediately.
- ⚠ **The requester may need to sign out and back in.** A new grant does not reach a session already
  open; that cost an hour on 2026-08-18 with every permission reading back correct. The notification
  must say so, and must **not** say access is immediate.

## 4. The page split

Chosen by the client over filtering rows in code, and it is the right instinct: **the page grant is the
boundary, rather than a filter inside the page.**

| Page | Policy | Audience |
|---|---|---|
| `Requests.aspx` | `["APR"]` | the approver queue — Heads of Unit |
| `Request-Access.aspx` | **open to everyone with site access** | the request form |
| `My-Submissions.aspx` | `["UPL"]`, unchanged | where a PIC tracks their own requests |

### 4.1 ⚠ Key the queue on `APR`, never on `SHARE`

Verified against the whole `PERSONAS` array rather than assumed:

- **`APR` is held by exactly one persona, `hou`.** So `APR` *is* Head of Unit — and because `hou`'s
  `namingRole` is `APR`, the group is literally `GHO_GF_TAX_APPROVER`. **There is no separate HoU
  group**; the approver group is the Head of Unit.
- **`SHARE` is held by three** — `hou`, `clevel_global` and `clevel_segment`.

`SHARE` is the intuitive key for a sharing screen and it is **wrong**: it would quietly admit both
C-Level personas to the approval queue. Pin both holder sets by test, as the HC clearance work did —
"role X is exclusive to persona Y" is a claim about the whole array, and it has been wrong before.

### 4.2 A PIC's own requests move to My Submissions

Narrowing `Requests.aspx` to `APR` takes visibility from PICs, who raise deletion and share requests
there today. Those move to **My Submissions** — the page where they raised them, and which exists for
exactly this question: what happened to the thing I submitted.

### 4.3 `PagePolicy` needs a third state

Today it expresses *"these roles"* or *"administrators only"*. It cannot express **"everyone who can
open the site"**, which is what the request form needs: the people who need it hold nothing, by
definition.

Leaving the page unmatched gives the right outcome for the wrong reason — an unmatched page is left
inheriting, so everyone can open it, but the next person tightening the patterns breaks it without
knowing. So **`openToAll: true`**, meaning *derivation deliberately leaves this page alone*, logged as a
decision rather than an omission.

## 5. What it deliberately does not do

- **No search.** CRS Search applies ACLs at query time, so a document the requester cannot read does not
  appear — by design. A link is the only way in, and that is correct: browsing for things you cannot see
  is a discovery channel of its own.
- **No HC path.** An uncleared person requesting an HC document is refused at the form with *"contact
  your Head of Unit"*, never routed. A one-file exception into a library they hold nothing in is exactly
  the leak the separate HC library pair exists to prevent.
- **No auto-approval**, even for someone already in the unit. If they are in the unit and still cannot
  see it, the interesting question is why — usually HC — and that deserves a human.
- **It does not intercept the Share button.** Settled 2026-07-23. This is the sanctioned route, made the
  *only working* route by turning Access Requests off — stronger than a guard, because there is nothing
  left to walk around.

## 6. Revoke — designed here, three decisions open

`2026-08-15` §5.3 specifies a Revoke action and a `ShareRevoked` audit event. Neither was built —
verified: no revoke path exists in `Requests.tsx`. Without it every approved grant is permanent until
someone edits SharePoint by hand, and the unique-scope count only grows.

**It is two mechanisms, not one**, and this is the part that must not be got wrong:

- **Internal recipient** — a role assignment on the file. Revoke = `removeroleassignment` by principal
  id, exactly as the access screens already do.
- **External recipient** — a **sharing link**: a different object, a different API. Removing a user's
  assignment does not kill the link, and anyone holding the URL still gets in. A revoke handling only
  the first case would **report success while leaving an external recipient with working access** — the
  worst outcome available to this feature.

| Decision | Recommendation |
|---|---|
| Reset inheritance when the last grant on a file goes? | **Yes** — reclaims the unique scope, so revoke cleans up the cost as well as the access. |
| Who may revoke? | **Any Head of Unit for that unit**, not only the approver who granted it. People leave, and a grant nobody can end is the problem being solved. |
| Expiry as well? | **Defer.** Expiry needs something to enforce it — a scheduled flow, since nothing here runs server-side — and that is a second unbuilt dependency on a feature already waiting on flows. Revoke runs in the approver's browser like everything else. |

## 7. Open finding on the EXISTING share flow

Not introduced here, and it should be fixed alongside. **A share recipient is typed as an email
address** (`MySubmissions.tsx` → `parseRecipients` → `{Key: …}`), and a Microsoft 365 group or
distribution list has one. `SP.Web.ShareObject` accepts it, so one approval can grant a document to
everyone in that list — **including people who join it later** — while rendering on the approver's
screen, and in the audit row, as a single recipient.

Recommended: **resolve each recipient before the approver decides, and label it** —
`finance-all@… — a group (214 people)`. It blocks nothing and makes the decision informed. Refusing
group recipients outright was considered and rejected: it removes a legitimate use, and people will
list the addresses individually instead.

⚠ Confirm first whether SDG's tenant permits sharing to a group at all. If it does, revoke must handle
it too — revoking a group share removes the group's assignment, not each person's.

## 8. Files

| File | Change |
|---|---|
| `src/shared/pageAccessPolicy.ts` | `openToAll`; `Requests.aspx` → `["APR"]`; a rule for the request form |
| `src/shared/pageAccessPolicy.test.ts` | `APR`-only queue pinned; `SHARE` must NOT open it; both holder sets asserted over `PERSONAS` |
| `src/shared/pageGrants.ts` | derivation skips an `openToAll` page, and logs that it did |
| `src/shared/requests.ts` | the `Access` request type and its state machine |
| `src/webparts/requests/components/Requests.tsx` | the decision UI, the group-add option, revoke |
| `src/webparts/requests/` *(new page)* | the request form |
| `src/webparts/mySubmissions/components/MySubmissions.tsx` | a PIC's own request status |

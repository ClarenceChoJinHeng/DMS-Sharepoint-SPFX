# Share expiry enforcement — Power Automate runbook

**Date:** 2026-08-30
**Status:** DESIGN — nothing is built.
**Flow name:** `CRS — Expire shares`
**Client instruction:** *"2. Build the enforcement, lets build the flow."*

---

## 0. Why this exists

`CRS Requests` has an **`ExpiresAt`** column. The uploader picks a date, the approver sees
*"until 2026-08-30"*, and the confirmation dialog says access lasts
*"until … or 2026-08-30 passes"*.

**None of that is true today.** `performShare` posts to `SP.Web.ShareObject` with `url`,
`peoplePickerInput`, `roleValue`, `groupId`, `propagateAcl`, `sendEmail`,
`includeAnonymousLinkInEmail`, `emailSubject`, `emailBody` and `useSimplifiedRoles` — **no
expiration field of any kind.** `expiresAt` reaches the list column and three display strings and
stops there.

> ⚠ **AND IT CANNOT BE FIXED IN THE SHARE CALL.** SharePoint has no per-grant expiry for a direct
> user grant. Expiring access is either an anonymous-link property (not what this uses — these are
> `role:` grants to named people) or a tenant-level Entra guest-lifecycle feature. **So the promise
> on screen can only be kept by something that comes back later and removes the grant.** That is
> this flow.

Until it exists, the honest alternative is to stop promising it — relabel the field *"review by"*.
The client chose enforcement.

---

## 1. Shape

| | |
|---|---|
| Trigger | **Recurrence**, once daily |
| Runs as | **the service account** — see §7 |
| Reads | `CRS Requests` |
| Writes | the file's ACL, then the request row |
| Downstream | the existing `CRS — Audit request activity` flow logs it |

---

## 2. Trigger

**Recurrence**, Frequency `Day`, Interval `1`, at **02:00** in the site's timezone.

Nightly is deliberate. An hourly run would revoke a share partway through the day it expires, and
nobody's expectation of *"until the 30th"* is *"until 09:00 on the 30th"*.

---

## 3. Find the expired shares — `GetExpired`

**Get items**, list `CRS Requests`.

**Filter Query** (the date is an expression — build it in the fx editor, not as typed text):

```
RequestType eq 'Share' and Status eq 'Approved' and ExpiresAt ne null and ExpiresAt lt '<startOfDay(utcNow()) formatted yyyy-MM-ddTHH:mm:ssZ>'
```

> ⚠⚠ **`ExpiresAt ne null` IS LOAD-BEARING.** A blank expiry means *"no expiry"* — a legitimate and
> common choice — and an OData comparison against a null column is not reliably false. Omit this
> clause and **the flow revokes every share that has no expiry date at all**, on its first run,
> silently. That is the single most destructive mistake available here.

> ⚠ **`startOfDay`, NOT `utcNow()` — this is a real off-by-one.** `ExpiresAt` is written from the
> browser as a date at midnight, so a share expiring `2026-08-30` is stored `2026-08-30T00:00:00Z`.
> Comparing against `utcNow()` makes it expire **the moment the day begins** — the user picks the
> 30th and loses access at 00:01 on the 30th. `startOfDay(utcNow())` revokes only once the expiry
> date is strictly in the past, which gives the full day the user asked for.

> ⚠ **`Status eq 'Approved'` and nothing wider.** `Revoked` rows are already done; `Failed` rows
> never granted anything; `Pending` rows have not been decided, and revoking one would carry out a
> decision nobody has made.

Set **Top Count** to 100 and leave pagination off. If a site ever has more than 100 expiring in one
night, the next night's run takes the rest — the flow is idempotent, so lagging is safe and
paginating a destructive loop is not.

---

## 4. Per row — `For each expired`

### 4.1 ⚠ First, check nothing else still grants this recipient — `OtherLiveShares`

**Get items** on `CRS Requests`:

```
RequestType eq 'Share' and Status eq 'Approved' and ItemUniqueId eq '<ItemUniqueId>' and Id ne <Id> and (ExpiresAt eq null or ExpiresAt ge '<startOfDay>')
```

> ⚠⚠ **TWO SHARES CAN EXIST ON ONE FILE, AND ONE OF THEM MAY NOT HAVE EXPIRED.** Revoking without
> this check takes access away from somebody whose share is still valid — with nothing on any screen
> explaining it, because their request still reads `Approved`. Observed live on 2026-08-30: one test
> file carried **two** share rows for the same recipient, one revoked and one approved.
>
> If any live row names the same recipient, **skip that recipient** and leave the grant alone. The
> expired row still gets its note written, so it is not reprocessed every night.

### 4.2 Resolve each recipient — `EnsureUser`

`ShareWith` is a semicolon/comma list of addresses. Split it, and for each:

```
POST {site}/_api/web/ensureuser
Body: { "logonName": "<address>" }
Headers: Accept: application/json;odata=nometadata
```

Read `Id` from the response.

> ⚠ **`ensureuser` RESOLVES, IT DOES NOT INVITE** — learned 2026-08-27, after a wrong diagnosis sent
> the investigation to the site's sharing setting. For an address that was shared with, the
> principal exists; that is what the share created. A `could not be found` here therefore means the
> person is **already** gone, which is a SUCCESS for this flow: mark the recipient done and carry
> on. **Do not fail the run.**

### 4.3 Remove the grant — `RemoveGrant`

```
POST {site}/_api/web/GetFileById(guid'<ItemUniqueId>')/ListItemAllFields/roleassignments/removeroleassignment(principalid=<Id>)
```

> ⚠ **`removeroleassignment` PER RECIPIENT — NEVER `resetroleinheritance`.** The app's *Revoke all*
> button resets inheritance because a human chose "everyone"; here the flow is acting on ONE expired
> request, and resetting would drop every other share on that file, including the live ones §4.1
> exists to protect. It would also silently discard anything an administrator granted by hand.

> ⚠ **A 404 ON THE FILE IS A SUCCESS, NOT A FAILURE.** The document may have been deleted or
> archived since. There is nothing to revoke, and failing the run would leave the row `Approved` and
> retry it every night for ever.

### 4.4 Write the row — `MarkRevoked`

**Update item** on the request row:

| Field | Value |
|---|---|
| `Status` | `Revoked` |
| `RevokedBy` | **the service account's address** |
| `DecisionNote` | the existing note **plus** a new line (below) |

```
Access for <recipients> ended automatically on <yyyy-MM-dd> — the expiry date on this request passed.
```

> ⚠ **APPEND, NEVER OVERWRITE.** That note holds the approver's reason for granting access in the
> first place; replacing it destroys half the trail. Same rule the in-app revoke follows.

> ⚠ **`RevokedBy` TAKES A REAL ADDRESS, NOT THE WORD "SYSTEM".** `CRS — Audit request activity`
> reads `RevokedBy` into `ActorEmail`, and an audit row whose actor is not a resolvable principal is
> worth less than one naming the account that actually did it. *Why* it happened belongs in the
> note, which is prose.

> ⚠ **`Status` MUST ONLY BECOME `Revoked` IF EVERY RECIPIENT WAS ACTUALLY REMOVED.** If §4.1 spared
> one, the request is partly live: leave `Status` as `Approved` and append the note anyway. Without
> some marker it is picked up again tomorrow and every night after; the appended note is the
> simplest one, and can be filtered out in §3 if the repetition ever becomes noisy.

---

## 5. What happens downstream, for free

`CRS — Audit request activity` triggers on **created or modified** on `CRS Requests`, so writing
`Status = Revoked` produces a `ShareRevoked` audit row naming the service account. No change is
needed there — **provided the `Revoked` branch has been added to its `EventKind` expression** and
`ActorEmail` prefers `RevokedBy`. Both were done on 2026-08-30 and verified live.

The requester's **My Submissions** then reads *"This share was approved, and the access has since
been taken back"*, with no name attached — which is correct, because no person took it back.

---

## 6. Notification — deliberately NOT in v1

Nobody is emailed when a share expires. The recipient simply stops being able to open the link.

Worth raising rather than assuming: an expiry is what the uploader asked for, so a warning three
days before is a courtesy, while a mail *after* the fact tells somebody about a door that has
already closed. If the client wants it, it is one extra branch over the same query with `ExpiresAt`
between today and today + 3.

---

## 7. Building it

> ⚠ **SIGN IN AS THE SERVICE ACCOUNT BEFORE CREATING THE FIRST ACTION.** A flow runs under its
> connection, and the connection is created implicitly by the first action — so the builder's
> account is baked in for life. Built as a person, this stops **silently** when that password
> changes, and presents months later as "shares are not expiring", which nobody is watching for.

> ⚠ **THE SERVICE ACCOUNT NEEDS MANAGE PERMISSIONS ON THE FILES.** `removeroleassignment` is a
> permissions change. Site collection administrator is the simplest guarantee; the custom
> `CRS Share` level also contains Manage Permissions, which is the whole reason it exists.

Set **Concurrency control = 1** on the `For each`. Two runs revoking the same file at once is not
dangerous — the second gets a 404 and treats it as success — but serial execution keeps the run
history readable. ⚠ Concurrency cannot be changed once enabled.

---

## 8. Testing

1. Raise a share, approve it, confirm the recipient can open the file.
2. Edit the request row by hand: set `ExpiresAt` to **yesterday**.
3. Run the flow manually.
4. **The recipient can no longer open the file** — the decisive check. Ask them, or use
   Check Permissions on the file.
5. The row reads `Revoked`, `RevokedBy` is the service account, and the note carries the new line.
6. `CRS Audit Log` gains a `ShareRevoked` row naming the service account.
7. **Run it a second time and confirm it does nothing.** The row is no longer `Approved`, so it
   falls out of the query. A destructive flow that is not idempotent is a nightly liability.
8. ⚠ **Then the one that matters most: a share with a BLANK `ExpiresAt`.** Run the flow and confirm
   it is **untouched**. That is the §3 null clause, and its failure mode is taking away access
   nobody agreed to give up.

---

## 9. Open questions for the client

1. **Should the recipient be warned before expiry?** (§6)
2. **Should an expired share be re-requestable in one click**, or does the uploader raise a fresh
   request? Today it is the latter, which is probably right.
3. **Should `ExpiresAt` be mandatory?** It is optional now, so "no expiry" is the default outcome of
   leaving the field blank — the opposite of what an expiry feature implies.

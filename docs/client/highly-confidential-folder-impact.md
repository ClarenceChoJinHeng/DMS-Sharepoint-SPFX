# Highly Confidential Folders — What You Should Know Before We Build

**Date:** 31 July 2026
**For:** SD Guthrie DMS stakeholders
**From:** Trinergy Digital

You have asked for Highly Confidential documents to be stored in a separate,
locked folder underneath each Document Type. We can build that, and we are
proceeding with it. This note sets out three consequences so there are no
surprises later — one of which needs a decision from you.

---

## 1. What we are building

Every Document Type folder gets an `HC` sub-folder with its own lock:

```
TREASURY / 2026 / Invoice /
                    ├── HC/               ← locked: HC members + approver only
                    └── (ordinary files)  ← the unit's uploaders
```

- Only people in that unit's **HC group** can upload a Highly Confidential
  document. Everyone else will not even see the option in the upload form.
- The unit's **existing approver** approves HC documents. No separate HC
  approver is needed.
- **Confidential** and **Restricted** stay exactly as they are today — labels on
  the document, no folder, no access change.
- The locked folders are created **in advance** by an administrator, never by
  the person uploading. This is what guarantees a Highly Confidential file is
  never visible to colleagues, not even momentarily.

---

## 2. Consequence one — system performance

Each locked folder is what SharePoint calls a **unique permission scope**.
Microsoft's guidance is to keep these under **5,000 per library**, with a hard
limit of 50,000.

Placing the HC folder under Document Type creates one lock for every
combination of unit, year and document type:

| | Locked folders | Per library |
| --- | --- | --- |
| **HC under Document Type** (your request) | ~14,600 | **~7,300** |
| HC directly under the Unit (alternative) | ~270 | ~135 |

At roughly 7,300 per library we are about **1.5× over Microsoft's recommended
figure**, though still well inside the hard limit. The practical effect is
slower page loads in the document libraries and slower permission checks,
becoming more noticeable as document volume grows.

We consider this acceptable. It is not a failure condition, but it is a
performance cost that will not go away.

---

## 3. Consequence two — setup and maintenance time

Folder setup is run by an administrator through the Folder Manager tool. Our
measured baseline is **6,090 folders in 111 minutes** for 48 units.

Adding HC folders under every Document Type adds roughly **73,000 additional
operations**, because each locked folder must be created, unlocked from its
parent, and then have permissions applied.

| Task | Estimated time |
| ---- | -------------- |
| Full setup today (133 units) | ~5 hours |
| Full setup with HC folders | **~20+ hours** |

This is a background process that can run unattended, but it does need to
finish before Highly Confidential uploads will work.

**The part that matters most day to day:** adding a new **Year** or a new
**Document Type** creates new folders that have no HC lock yet.

- Adding a **new year** — for example moving into 2027 — blocks all Highly
  Confidential uploads for that year until a full setup run completes.
- Adding a **new document type** does the same for that type, across every year.

The upload form will refuse the upload with a clear message rather than putting
the file somewhere unsafe, so nothing is ever exposed. But **please plan a setup
run whenever a year or document type is added**, ideally ahead of when people
need it, and allow up to a day for it.

Each new document type also adds roughly **2,400** locked folders, and each new
year roughly **3,500**. Over several years this accumulates towards Microsoft's
50,000 ceiling, so it is worth reviewing periodically.

---

## 4. Consequence three — a decision we need from you

This one has no default answer, and it affects one specific group of users.

You described a PIC who **can see Highly Confidential documents but must not see
Confidential or Restricted ones**. That person belongs only to the HC group.

The difficulty is how they reach their HC folder. To click down through
SharePoint to `Unit → 2026 → Invoice → HC`, they need at least view rights on
each folder along the way — and the **Invoice folder is where the ordinary
Confidential and Restricted files are kept**. Giving them enough access to walk
through it also lets them see the files inside it.

SharePoint has no "walk through this folder without seeing what is in it"
permission. There are two options:

| | What the user experiences | Effect on your requirement |
| --- | --- | --- |
| **Option A** | Browses the library normally, like everyone else | They **will** see Confidential and Restricted files in the document types they pass through |
| **Option B** | Cannot browse to it; reaches the HC folder only through a direct link or the DMS web parts | Requirement fully met |

Option B works well if these users always come in through the DMS screens we
have built, which link directly to the right place. It does not work if they
expect to navigate SharePoint folders themselves.

This only arises because the HC folder sits under Document Type. If it sat
directly under the Unit, the path a user walks through would contain only other
folders — no documents — and the question would not come up at all.

**We need you to choose Option A or Option B before this user type can be
built.** Everything else described here proceeds regardless.

---

## 5. The alternative, for completeness

Placing the HC folder directly under the Unit instead:

```
TREASURY / HC / 2026 / Invoice /   ← locked once, at the unit
TREASURY / 2026 / Invoice /        ← ordinary files
```

Each document type still has its own HC folder, and Highly Confidential files
are still separated by year and type exactly as you want. The only difference is
**where the lock sits** — one lock on the whole HC wing, rather than a separate
lock on every room inside it.

This would give:

- ~270 locked folders instead of ~14,600
- ~5 hour setup instead of ~20+ hours
- No delay when a new year or document type is added
- The section 4 decision would not be needed at all

We are **not** recommending a change of course — we are building what you asked
for. We are recording the alternative so the trade-off is on the record and can
be revisited if the performance cost becomes a problem in practice.

---

## 6. What we need from you

| # | Item | Needed by |
| - | ---- | --------- |
| 1 | Decide **Option A or B** in section 4 | Before the HC-only PIC role is built |
| 2 | Create two permission levels on the site: **DMS Approve** and **DMS Delete** (we will supply exact settings) | Before setup runs |
| 3 | Confirm that **approvers can read Highly Confidential documents** — unavoidable, since approving requires opening the file | Before go-live |
| 4 | Confirm that Highly Confidential documents in the **approved** library should be locked the same way | Before setup runs |

Item 3 is worth stating plainly: because you have asked the unit's normal
approver to approve HC documents rather than a dedicated HC approver, that
approver can read every Highly Confidential document in their unit. If that is
not acceptable, a separate HC approver group would be required and we should
discuss it now rather than after go-live.

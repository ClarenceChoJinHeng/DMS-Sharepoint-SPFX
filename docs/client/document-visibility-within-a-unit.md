# Who can see whose documents

**For:** SD Guthrie — CRS
**Last updated:** 9 August 2026

A plain-language answer to the question *"can we make each person see only their own files?"*
Written to be shown or read aloud as-is.

---

## The short answer

**Everything is shared within the unit — including work still waiting for approval.**

Changed **2026-08-19** at the client's request. Until then, pending and rejected submissions were
visible only to their author and the Head of Unit. They are now visible to everyone who can open the
library for that unit.

|  | Approval Document | Documents |
|---|---|---|
| Your **pending** submission | you, your Head of Unit, **and the other uploaders in your unit** | — |
| Your **rejected** submission | you, your Head of Unit, **and the other uploaders in your unit** | — |
| An **approved** document | *(removed — it has moved)* | everyone in the unit |

**Nobody outside the unit sees any of it.** The people who can open a unit's approval library are its
uploaders and its Head of Unit — nobody else, at any stage. And the Highly Confidential libraries stay
separate: only people in the unit's **Highly Confidential** uploader group can open those at all. A
plain uploader in the same unit cannot see a Highly Confidential document, or the library it lives in.

Why the change: a folder created by one uploader used to be invisible to their colleagues, so the
next person could not reach their own file inside it. Making everything visible within the unit fixes
that directly, and removes two background processes that existed only to work around it.

So a colleague in your unit cannot see what you are still working on, what you submitted
yesterday, or anything that was sent back to you. They can see what has been **approved** — and an
approved document is one your Head of Unit has signed off as the unit's record.

## Why approved documents are shared

Two reasons, one practical and one about the system.

**Practical.** The point of the archive is that the unit can find its own documents. If each person
could only see their own, then when someone is on leave — or leaves the company — their documents
become unreachable by the people who need them. The unit would have a filing cabinet where every
drawer has a different key.

**Technical.** SharePoint's smallest privacy boundary is a **folder**. It can hide a folder from
someone, but it cannot hide one person's files from another person inside a shared folder — not
without also making your Heads of Department, Heads of Unit and C-Level able to **edit and change**
documents. In SharePoint the permission to see hidden items cannot be separated from the permission
to change them. Since those roles were specifically asked to be view-only, that route is closed.

## Could it be done at all?

Yes — but it means changing the folder structure.

It would require a **folder for each person** inside every unit:

```
Group Head Office / Group Finance / CORU / Ahmad / 2026 / Tax Return /
Group Head Office / Group Finance / CORU / Siti  / 2026 / Tax Return /
```

That works, and it is secure. But it changes what the archive *is*:

- Documents would be filed **by person first**, then by year and type.
- "Show me all of CORU's 2026 Tax Returns" would no longer be a folder you open — you would have to
  search or filter across every person's folder.
- When someone leaves, their folder remains, named after them, holding the unit's documents.

That is a decision about how SD Guthrie wants its records organised, not a technical limitation. The
current structure was chosen so documents belong to the **unit**; a per-person folder makes them
belong to the **individual**.

## If two people genuinely must not see each other's work

Then they belong in **different units**. That is fully supported and costs very little — one new
unit in the term list, one abbreviation, and its access groups. The system is designed for it.

> **A note on SubUnits.** SubUnit folders organise a unit's filing; they do **not** separate access.
> Anyone who can see a Unit can see every SubUnit inside it, and every document in them. So if two
> people must not see each other's documents, putting them in different **SubUnits** will not achieve
> it — it has to be different **Units**.

This is usually the right answer, because "these two must not see each other's documents" almost
always means they are doing genuinely separate work.

---

## Summary for a meeting

> Everything in a unit is shared within that unit — work in progress as well as approved documents —
> because the unit needs to be able to find its own records. Nobody outside the unit sees any of it,
> and Highly Confidential documents stay separate: only the people cleared for them can open those
> libraries at all.
>
> Making documents private per person is possible, but only by giving every person their own folder
> inside each unit — which changes the filing structure from unit-based to person-based. If two people
> must not see each other's work, the supported answer is to put them in separate units.

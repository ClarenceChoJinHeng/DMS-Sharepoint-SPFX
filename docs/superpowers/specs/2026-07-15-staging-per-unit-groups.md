# Staging Library — Per-Unit Group & Permission Build Sheet

**Date:** 2026-07-15
**Scope:** Staging library only. Documents library + Home page = later pass.
**Goal:** Per-unit isolation with progressive-disclosure browsing (Model B).

---

## Decisions locked

| Decision | Choice |
|---|---|
| Hierarchy | **Business Segment → Department → Unit** |
| Isolation boundary | **Per Unit** (each unit private) |
| Visibility model | **Model B — progressive disclosure** (break every level; browse down by membership) |
| Membership | **Nested groups** — assign user to their **unit** group only; segment/dept visibility flows up |
| Roles (Staging) | **Uploader** (Contribute) + **Approver** (Design). Reader = Documents library, later pass. |
| Naming | `DMS_BS_<SEG>` · `DMS_DPT_<DEPT>` · `DMS_<DEPT>_<UNIT>_UPL` / `_APR` |

> Reader role (`DMS_..._RDR`) is **not** part of Staging — readers only touch the Documents library. Deferred.

---

## Naming convention

| Level | Group | Example |
|---|---|---|
| Business Segment | `DMS_BS_<SEG>` | `DMS_BS_GHO` |
| Department | `DMS_DPT_<DEPT>` | `DMS_DPT_GF` |
| Unit — uploaders | `DMS_<DEPT>_<UNIT>_UPL` | `DMS_GF_TAX_UPL` |
| Unit — approvers | `DMS_<DEPT>_<UNIT>_APR` | `DMS_GF_TAX_APR` |

**Token legend (confirm real values):**
- `GHO` = Group Head Office (Business Segment)
- `GF` = Group Finance (Department)
- `TAX` = Tax (Unit)

---

## Nesting (assign once, visibility flows up)

```
DMS_BS_GHO
  └─ DMS_DPT_GF                (member of DMS_BS_GHO)
       ├─ DMS_GF_TAX_UPL       (member of DMS_DPT_GF)   ← add uploaders HERE
       ├─ DMS_GF_TAX_APR       (member of DMS_DPT_GF)   ← add approvers HERE
       ├─ DMS_GF_TREAS_UPL
       └─ DMS_GF_TREAS_APR
```

Add a person to their **unit** group only. Transitive membership (SharePoint honors nested security groups) grants them the dept + segment groups automatically → they can browse GHO → Finance → their unit.

---

## Permission per level (Model B)

Break inheritance at **every** level and grant that level's group:

| Folder level | Group granted | SharePoint level | Why |
|---|---|---|---|
| `Staging/GHO` (Segment) | `DMS_BS_GHO` | **Read** | Browse into segment; can't add files here |
| `Staging/GHO/Finance` (Dept) | `DMS_DPT_GF` | **Read** | Browse into dept; can't add files here |
| `Staging/GHO/Finance/Tax` (Unit) | `DMS_GF_TAX_UPL` | **Contribute** | Upload + tag |
| `Staging/GHO/Finance/Tax` (Unit) | `DMS_GF_TAX_APR` | **Design** | Review + approve |

Security trimming does the rest: a user sees only the branches they're a member of; everything else is hidden from the listing.

---

## Per-folder apply checklist (every 🔒 folder)

1. **Break inheritance** — Manage access → Advanced → Stop Inheriting.
2. **Strip broad groups** SharePoint copied in — remove any site-wide `DMS_UPLOADER`, Members, Visitors. Leave only Owners/Admins.
3. **Grant** the level's group at the level above:
   - Segment/Dept folder → its group = **Read**
   - Unit folder → `_UPL` = **Contribute**, `_APR` = **Design**
4. Repeat down the tree. (Automatable via the Folder Onboarding web part — run it per level.)

---

## WORKED SAMPLE (replace names with real org data)

> This is illustrative structure using fragments seen so far. **Swap in the client's real
> Business Segments / Departments / Units.**

### Tree

```
Staging/
  Group Head Office/                 🔒 DMS_BS_GHO = Read
    Finance/                         🔒 DMS_DPT_GF = Read
      Tax/                           🔒 DMS_GF_TAX_UPL = Contribute · DMS_GF_TAX_APR = Design
      Treasury/                      🔒 DMS_GF_TREAS_UPL = Contribute · DMS_GF_TREAS_APR = Design
    Account/                         🔒 DMS_DPT_ACC = Read
      Account-Team-1/                🔒 DMS_ACC_T1_UPL = Contribute · DMS_ACC_T1_APR = Design
      Account-Team-2/                🔒 DMS_ACC_T2_UPL = Contribute · DMS_ACC_T2_APR = Design
```

### Grant table

| Folder (server-relative under Staging) | Group | Level |
|---|---|---|
| `/Group Head Office` | `DMS_BS_GHO` | Read |
| `/Group Head Office/Finance` | `DMS_DPT_GF` | Read |
| `/Group Head Office/Finance/Tax` | `DMS_GF_TAX_UPL` | Contribute |
| `/Group Head Office/Finance/Tax` | `DMS_GF_TAX_APR` | Design |
| `/Group Head Office/Finance/Treasury` | `DMS_GF_TREAS_UPL` | Contribute |
| `/Group Head Office/Finance/Treasury` | `DMS_GF_TREAS_APR` | Design |
| `/Group Head Office/Account` | `DMS_DPT_ACC` | Read |
| `/Group Head Office/Account/Account-Team-1` | `DMS_ACC_T1_UPL` | Contribute |
| `/Group Head Office/Account/Account-Team-1` | `DMS_ACC_T1_APR` | Design |
| `/Group Head Office/Account/Account-Team-2` | `DMS_ACC_T2_UPL` | Contribute |
| `/Group Head Office/Account/Account-Team-2` | `DMS_ACC_T2_APR` | Design |

---

## REAL BUILD — Group Finance (Phase 1)

Segment = **Group Head Office** (`GHO`) · Department = **Group Finance** (`GF`).
Only this branch is built now; other segments/departments later.

### Unit token map

| # | Unit (folder name) | Token |
|---|---|---|
| 1 | Org Culture & Workplace Services - Admin & Sports | `OCWS` |
| 2 | Finance Operations | `FINOPS` |
| 3 | Finance Upstream Malaysia | `FUM` |
| 4 | Tax | `TAX` |
| 5 | Treasury | `TREAS` |
| 6 | Finance Land & SDGRE | `FLAND` |
| 7 | Global Strategic Procurement | `GSP` |
| 8 | Group IT & Projects | `GITP` |
| 9 | Compliance & Operational Risk (CORU) | `CORU` |
| 10 | Strategy & Price Risk Management | `SPRM` |

### Groups to create (22)

**Parent/container (2):** `DMS_BS_GHO`, `DMS_DPT_GF`
**Per unit (20):** `DMS_GF_<TOKEN>_UPL` (Contribute) + `DMS_GF_<TOKEN>_APR` (Design)

```
DMS_GF_OCWS_UPL    DMS_GF_OCWS_APR
DMS_GF_FINOPS_UPL  DMS_GF_FINOPS_APR
DMS_GF_FUM_UPL     DMS_GF_FUM_APR
DMS_GF_TAX_UPL     DMS_GF_TAX_APR
DMS_GF_TREAS_UPL   DMS_GF_TREAS_APR
DMS_GF_FLAND_UPL   DMS_GF_FLAND_APR
DMS_GF_GSP_UPL     DMS_GF_GSP_APR
DMS_GF_GITP_UPL    DMS_GF_GITP_APR
DMS_GF_CORU_UPL    DMS_GF_CORU_APR
DMS_GF_SPRM_UPL    DMS_GF_SPRM_APR
```

### Nesting

```
DMS_BS_GHO
  └─ DMS_DPT_GF
       ├─ DMS_GF_OCWS_UPL / _APR
       ├─ DMS_GF_FINOPS_UPL / _APR
       ├─ DMS_GF_FUM_UPL / _APR
       ├─ DMS_GF_TAX_UPL / _APR
       ├─ DMS_GF_TREAS_UPL / _APR
       ├─ DMS_GF_FLAND_UPL / _APR
       ├─ DMS_GF_GSP_UPL / _APR
       ├─ DMS_GF_GITP_UPL / _APR
       ├─ DMS_GF_CORU_UPL / _APR
       └─ DMS_GF_SPRM_UPL / _APR
```
Add a person to their **unit** `_UPL`/`_APR` group only — dept + segment visibility flows up.

### Folder + grant table (12 folders to break inheritance)

Base path: `.../Staging/Group Head Office/Group Finance/<Unit>`

| # | Folder | Group | Level |
|---|---|---|---|
| 1 | `/Group Head Office` | `DMS_BS_GHO` | Read |
| 2 | `/Group Head Office/Group Finance` | `DMS_DPT_GF` | Read |
| 3 | `.../Org Culture & Workplace Services - Admin & Sports` | `DMS_GF_OCWS_UPL` / `_APR` | Contribute / Design |
| 4 | `.../Finance Operations` | `DMS_GF_FINOPS_UPL` / `_APR` | Contribute / Design |
| 5 | `.../Finance Upstream Malaysia` | `DMS_GF_FUM_UPL` / `_APR` | Contribute / Design |
| 6 | `.../Tax` | `DMS_GF_TAX_UPL` / `_APR` | Contribute / Design |
| 7 | `.../Treasury` | `DMS_GF_TREAS_UPL` / `_APR` | Contribute / Design |
| 8 | `.../Finance Land & SDGRE` | `DMS_GF_FLAND_UPL` / `_APR` | Contribute / Design |
| 9 | `.../Global Strategic Procurement` | `DMS_GF_GSP_UPL` / `_APR` | Contribute / Design |
| 10 | `.../Group IT & Projects` | `DMS_GF_GITP_UPL` / `_APR` | Contribute / Design |
| 11 | `.../Compliance & Operational Risk (CORU)` | `DMS_GF_CORU_UPL` / `_APR` | Contribute / Design |
| 12 | `.../Strategy & Price Risk Management` | `DMS_GF_SPRM_UPL` / `_APR` | Contribute / Design |

> Folder names keep the full display text (spaces, `&`, `-`, `()` are all valid SharePoint folder chars).
> Segment/Dept folders (rows 1–2) get **Read** only; unit folders (rows 3–12) get **Contribute + Design**.

---

## Open items (deferred, not blocking Staging)

- Documents library — same tree, roles = UPL/APR = Read, add `DMS_..._RDR` = Read (if read-only population exists).
- Home page navigation / entry point.
- Term-store new data — must mirror this Segment/Dept/Unit tree.
- Confirm token abbreviations with client once real names arrive.

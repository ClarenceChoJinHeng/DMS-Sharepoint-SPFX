# Highly Confidential Securing — Design

**Status:** proposed (retention rule decided; implement after base build is verified)

## Goal

A file marked **Highly Confidential (HC)** must be hidden from everyone in its unit **except**
the unit's HC-cleared members — with **no exposure window** at any point. Other confidentiality
levels are plain labels with no access effect.

## Confirmed requirements

- **Placement is mandatory:** HC files live in `…/Unit/Year/Document Type/Highly Confidential/`
  — the HC folder sits **under the Document Type folder**.
- **Any unit uploader** may create an HC file (not only HC-cleared members).
- **Single file** per upload.
- **No window:** no peer may ever see the file, not even for seconds.
- **Retention:** the HC folder grants **HC group (`Role = HC`) + the unit's approver group
  (`Role = APR`) + site admins** (automatic). Everyone else is removed. *Configurable* — the
  client may later switch `APR` for a dedicated HC-approver role; that's a flow edit, no redeploy.

## Why this needs an elevated pre-upload flow

- Hiding a file = a permission break on its **folder** (grant only HC + approvers). The uploader
  has **Contribute**, not **Manage Permissions**, so the web part cannot secure anything.
- "No window" means the folder must be secured **before** the file is written. Auto-created
  folders inherit the (visible) unit ACL, so an uploader-created HC folder is briefly public.
- Because the secured folder grants no access to the general uploader group (peers are uploaders
  too — they'd see it), a normal uploader can **neither secure the folder nor write into it**.
  Both privileged actions therefore move to an elevated service.
- The "grant add-but-not-read" trick fails: the metadata-tagging step needs read/edit on the item.

**Conclusion:** for HC files the whole submission is handed to an elevated, HTTP-triggered
Power Automate flow (a service-account connection with Full Control on Staging) that
**secures → uploads → tags**. Normal (non-HC) files keep the current direct web-part path.

## Flow (HTTP-triggered, called synchronously by the web part)

Input (POST body from the web part): file (base64) + filename, `stagingLibrary`, the resolved
HC folder server-relative path, `PermTermGuid` (leaf unit term), and the full metadata set
(Confidentiality, per-level label/tid pairs, Year, Document Type, Document Date, Vendor).

1. **Ensure** `…/Document Type/Highly Confidential/` exists (`folders/AddUsingPath`, idempotent).
2. **Resolve retained groups** — GET DMS Group Map items where `TermGuid eq '{PermTermGuid}'`
   and `Role` in (`HC`,`APR`) → their `GroupId`s.
3. If the folder is **not already secured** (`ListItemAllFields.HasUniqueRoleAssignments = false`):
   - `ensureuser` each group (claim `c:0o.c|federateddirectoryclaimprovider|{GroupId}`),
   - `breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)` on the folder item,
   - `addroleassignment` each: HC group → Contribute (or Read), approver group → Design.
     (Site-collection admins retain access implicitly.)
4. **Upload** the file into the secured folder (`GetFolderByServerRelativeUrl(@f)/Files/Add`).
5. **Tag** the item (`items({id})/validateUpdateListItem` with the same `formValues` the web part
   builds for normal files — Confidentiality, level label+tid pairs, Year, Doc Type, Date, Vendor).
6. Return success (or a typed error the web part surfaces as a toast).

**No-window guarantee:** the file is written only *after* the folder is secured, and only ever
exists inside that secured folder. Peers never have a moment of visibility.

**Idempotency:** `HasUniqueRoleAssignments` short-circuits re-securing; the first HC file per
document type secures the folder, later ones just upload+tag.

## Web part changes

- **`PermTermGuid`** — a hidden Staging text column the web part fills on **every** upload with
  the **leaf unit's term GUID** (mode-independent key so the flow/other tooling can resolve the
  unit regardless of which segment's leaf column applies).
- **HC branch** — when the chosen Confidentiality label is **Highly Confidential**, the web part
  builds the HC folder path (`…/Document Type/Highly Confidential/`) and **hands the whole
  submission to the flow** instead of uploading directly; shows the flow's result.
- Normal confidentiality levels → unchanged direct upload.
- The trigger label ("Highly Confidential") should be configurable via DMS Config setting
  (default `Highly Confidential`) to avoid hardcoding.

## Data / setup impact

- **DMS Group Map:** add `Role = HC` rows (per unit: HC group Object ID → unit term GUID). `APR`
  rows already exist from the tiered model and are reused for retention.
- **HC groups:** one per unit (`DMS_<unit>_Highly_Confidential`). The **Onboarding web part** can
  provision these alongside unit folders/groups.
- **The flow:** one generic HTTP-triggered flow (service-account connection, Full Control on
  Staging). The trigger URL is a bearer secret embedded in the SPFx bundle — acceptable for an
  internal tenant; validate inputs in the flow. For **large** HC files (tens of MB), host the same
  logic in an **Azure Function** (AAD-authenticated) instead of Power Automate.

## Sequencing

Implement **after** the current 1.0.5.0 build (multi-segment + tiered detection) is verified live.
This feature adds an external service dependency; it should sit on a proven base.

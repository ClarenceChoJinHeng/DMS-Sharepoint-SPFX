# Tenant Seed-Data Runbook — Multi-Segment DMS

**Purpose:** everything a tenant needs so the deployed `sd-gatrie.sppkg` "just works" with
**data insertion only — no code changes.** The web part is fully data-driven: modes, level
chains, and the real Staging column internal names all come from SharePoint lists / config.

Do the steps in order. Where you see `‹…›`, paste the real GUID / internal name you captured.

---

## 0. Deploy the package

1. Upload `sharepoint/solution/sd-gatrie.sppkg` to the site App Catalog and deploy.
2. Grant the pending **Graph API permissions** (SharePoint admin → API access):
   `User.Read.All`, `Sites.Read.All`, `GroupMember.Read.All`, `Group.Read.All`.
3. Add the **Form** web part to the upload page and the **Reconciliation** web part to an
   admin-only page.

---

## 1. Term Store — create 5 term sets (group: *DMS Metadata*)

Each **Business Segment** set holds its own hierarchy; the leaf term = the permissioned Unit.

| Mode | Term set | Hierarchy (top → leaf) | Example |
|---|---|---|---|
| gho | **Group Head Office** | Department → Unit | Group Legal, Risk & Compliance → Group Compliance (GCO) |
| upstream | **Group Upstream Operations** | Region → Estate/Mill | Sabah → Lavang Estate |
| sdgi | **Group SDGI Operations** | Refinery → Department | Pasir Gudang → Finance |
| it | **Group Innovation & Technology** | I&T Operating Unit | Digital Delivery |
| projects | **Group-led Projects** | Project Name → Department → Unit | Blue Sky → Finance → Treasury |

> The **leaf** level of each set is what gets a folder + permissions. Intermediate levels are
> metadata + folder path only.

Capture each set's **Unique Identifier** — you need it for DMS Config `TermSetGuid`.

| Set | Term-set GUID |
|---|---|
| Group Head Office | `efa87c6a-9536-4f7c-910f-011bf7413b80` *(sandbox — verified)* |
| Group Upstream Operations | `‹upstream-guid›` |
| Group SDGI Operations | `‹sdgi-guid›` |
| Group Innovation & Technology | `‹it-guid›` |
| Group-led Projects | `‹projects-guid›` |

---

## 2. Staging library — create the level columns

All **Single line of text**. Create a **label** column and a **Tid** (term-GUID) column per level.
Also create the two Business Segment columns once (shared by every Business Segment mode).

| Level (used by) | Label column | Tid column |
|---|---|---|
| Business Segment (all BS modes) | `Business Segment` | `Business Segment Tid` |
| Department (gho, sdgi, projects) | `Department` | `Department Tid` |
| Unit (gho, projects) | `Unit` | `Unit Tid` |
| Region (upstream) | `Region` | `Region Tid` |
| Estate/Mill (upstream) | `Estate Mill` | `Estate Mill Tid` |
| Refinery (sdgi) | `Refinery` | `Refinery Tid` |
| I&T Operating Unit (it) | `IT Operating Unit` | `IT Operating Unit Tid` |
| Project Name (projects) | `Project Name` | `Project Name Tid` |

> **Avoid `/` and `&` in the display name when creating** (they force ugly `_x002f_`/`_x0026_`
> internal names). Use `Estate Mill`, `IT Operating Unit` — the user-facing label comes from
> the term set + the DMS Config `Levels` `label`, not the column name.

### Read back the frozen internal names (critical)

SharePoint freezes an internal name at creation and it is **not predictable** (a single word
stays clean, `Business Segment` may become `Business_x0020_Segment`, `Business Segment Tid`
became `BusinessSegmentTid` in the sandbox). Fetch the real names:

```
‹site›/_api/web/lists/getbytitle('Staging')/fields?$select=Title,InternalName&$top=500
```

Record label + Tid internal names per level — you paste them into DMS Config `Levels` (step 3).

Sandbox verified so far: `Business_x0020_Segment` / `BusinessSegmentTid`,
`Department` / `DepartmentTid`, `Unit` / `UnitTid`.

---

## 3. DMS Config list — mode rows + setting rows

List: **DMS Config**. Columns: `Title`, `ConfigType` (`mode`|`setting`), `ModeLabel`,
`Side` (`BusinessSegment`|`Project`), `TermSetGuid`, `StagingFolder`, `SortOrder`,
`Levels` (multi-line **plain text**), `SettingValue`.

### 3a. One `mode` row per segment (`ConfigType = mode`)

`Levels` is a JSON array; each level carries `label`, `column` (logical key) **and** the real
Staging internal names `labelCol` / `tidCol`. The internal names here are authoritative — the
code uses them verbatim, so this is what makes the solution portable with zero code changes.

| Title | ModeLabel | Side | TermSetGuid | StagingFolder | SortOrder |
|---|---|---|---|---|---|
| gho | Group Head Office | BusinessSegment | ‹gho-guid› | Group Head Office | 1 |
| upstream | Group Upstream Operations | BusinessSegment | ‹upstream-guid› | Group Upstream Operations | 2 |
| sdgi | Group SDGI Operations | BusinessSegment | ‹sdgi-guid› | Group SDGI Operations | 3 |
| it | Group Innovation & Technology | BusinessSegment | ‹it-guid› | Group Innovation & Technology | 4 |
| projects | Group-led Projects | Project | ‹projects-guid› | Group-led Projects | 5 |

`Levels` values (paste, substituting your captured `labelCol`/`tidCol`):

```jsonc
// gho
[{"label":"Department","column":"Department","labelCol":"Department","tidCol":"DepartmentTid"},
 {"label":"Unit","column":"Unit","labelCol":"Unit","tidCol":"UnitTid"}]

// upstream
[{"label":"Region","column":"Region","labelCol":"‹Region internal›","tidCol":"‹Region Tid internal›"},
 {"label":"Estate/Mill","column":"EstateMill","labelCol":"‹EstateMill internal›","tidCol":"‹EstateMill Tid internal›"}]

// sdgi
[{"label":"Refinery","column":"Refinery","labelCol":"‹Refinery internal›","tidCol":"‹Refinery Tid internal›"},
 {"label":"Department","column":"Department","labelCol":"Department","tidCol":"DepartmentTid"}]

// it
[{"label":"I&T Operating Unit","column":"ITOperatingUnit","labelCol":"‹ITOU internal›","tidCol":"‹ITOU Tid internal›"}]

// projects
[{"label":"Project Name","column":"ProjectName","labelCol":"‹ProjectName internal›","tidCol":"‹ProjectName Tid internal›"},
 {"label":"Department","column":"Department","labelCol":"Department","tidCol":"DepartmentTid"},
 {"label":"Unit","column":"Unit","labelCol":"Unit","tidCol":"UnitTid"}]
```

### 3b. `setting` rows (`ConfigType = setting`) — optional overrides

`Title` / `SettingValue` pairs. Defaults are baked into the code, so only add rows to override:
`termSet_documentType`, `termSet_yearPeriod`, `termSet_confidentiality`, `termSet_vendor`,
`stagingLibrary` (default `Staging`), `allowedExtensions` (comma list).

---

## 4. DMS Group Map list — who sees which unit

List: **DMS Group Map**. Columns (all Single line of text unless noted): `GroupId`,
`GroupName`, `Segment` (term-set GUID), `UnitTermGuid`, `Role` (Choice `UPL`|`APR`).

One row per (security group × unit). **Matching is by `GroupId` = Entra Object ID** — names are
cosmetic, so existing groups can be reused without renaming. Only **unit-level** rows are needed;
the form walks the term ancestry to reconstruct Segment → Dept → Unit.

Example (sandbox GHO / Group Legal, Risk & Compliance):

| GroupId (Object ID) | GroupName | Segment | UnitTermGuid | Role |
|---|---|---|---|---|
| ‹object-id› | GCO Uploaders | ‹gho-guid› | ‹GCO unit term guid› | UPL |
| ‹object-id› | GCO Approvers | ‹gho-guid› | ‹GCO unit term guid› | APR |
| … | … | … | … | … |

---

## 5. Staging folders + permissions

1. Create the folder tree down to each **Unit**:
   `Staging/‹StagingFolder›/‹Dept›/‹Unit›/` (e.g. `Staging/Group Head Office/Group Legal, Risk & Compliance/Group Compliance (GCO)/`).
   `Year` and `Document Type` subfolders are created automatically on first upload.
2. On each **Unit** folder: break inheritance, grant the unit's `UPL` group **Contribute** and
   its `APR` group **Design**. (Approvers can also be granted **Read** at the segment/dept
   folders for browse.)

---

## 6. Reconciliation — map units to folders

Open the **Reconciliation** web part (admin). It walks each mode's term set to the leaf and
writes `TermGuid → FolderUniqueId` into the **DMS Folder Map** list. Run it after every folder
add/rename. The form resolves the upload target by `FolderUniqueId`, so this mapping is required
before uploads succeed.

> Reconciliation's offline `DEFAULT_MODES` fallback lists Group Head Office only; the live term
> sets come from the term store, so add the other four sets there / in DMS Config as they exist.

---

## Deploy checklist (client site)

- [ ] Package deployed + Graph permissions approved (§0)
- [ ] 5 term sets created; GUIDs captured (§1)
- [ ] Level columns created; internal names read back (§2)
- [ ] DMS Config mode rows with `Levels` (real `labelCol`/`tidCol`) (§3a)
- [ ] DMS Group Map rows (Object IDs) (§4)
- [ ] Unit folders created + permissioned (§5)
- [ ] Reconciliation run → DMS Folder Map populated (§6)
- [ ] Test upload as an uploader → lands in `…/Unit/Year/Document Type/`, columns tagged

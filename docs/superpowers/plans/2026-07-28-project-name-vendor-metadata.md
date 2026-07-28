# Project Name + Vendor Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Capture free-text Project Name and Vendor/Customer Name on single-file uploads, store them on the file in SharePoint, display them on the Approval Document page, and relay out the upload form to match the approved mockup.

**Architecture:** All work is in three existing SPFx React web parts. `Form.tsx` gains one new text field, starts writing two columns it previously discarded, makes Document Date optional, and reorganises its JSX from two cards (File / Document Information) into two differently-scoped cards (Upload a Document / Document Folder Information). `ApprovalDocument.tsx` gains two rows in its metadata table. `BulkUpload.tsx` gets a one-line correction to a now-stale column name. Column internal names resolve config-first from the `DMS Config` list with hardcoded fallbacks, following the existing `settings.columns` pattern.

**Tech Stack:** SPFx 1.23.0, React, TypeScript, Heft toolchain (NOT gulp), Jest via `heft test`, SharePoint REST (`validateUpdateListItem`, `FieldValuesAsText`).

**Spec:** `docs/superpowers/specs/2026-07-28-project-name-vendor-metadata-design.md`

---

## Context you need before starting

**Verified column facts.** These were confirmed against the live `/fields` API on 2026-07-28. Do not guess internal names from display names — this project has been bitten by that before.

| Display name | Internal name | Type |
|---|---|---|
| `ProjectName` | `ProjectName` | Text |
| `Vendor/CustomerName` | `Vendor_x002f_CustomerName` | Text |

The old `Vendor` column **was deleted**. Two `FIELDS` constants and one read call still name it. Task 1 fixes the constants; Task 5 fixes the read.

**Why `_x002f_`:** SharePoint encodes special characters in internal names. The `/` in "Vendor/CustomerName" becomes `_x002f_`.

**Why the Approval page read key differs again:** `FieldValuesAsText` double-encodes the underscores in its *response keys*, so `Vendor_x002f_CustomerName` comes back as `Vendor_x005f_x002f_x005f_CustomerName`. `ProjectName` has no encoded characters, so it comes back unchanged. This is why Task 5 uses two different lookup styles.

**Testing reality.** This repo only unit-tests pure modules in `src/shared/`. There are no React component tests and no harness for them. Do not add one — it is a large unrelated lift. Only Task 2 has a unit test. Everything else is verified by `npx heft test` (type-check + lint + existing tests) plus the manual UAT in Task 7. This is the established pattern, not a shortcut.

**Baseline:** `npx heft test` currently reports **57 passing, 0 failures** and **16 lint warnings**. Those 16 are pre-existing — do not fix them, but do not add new ones either.

**Commands:**
- Tests + typecheck + lint: `npx heft test` (~26s)
- Full production build: `npm run build`
- Dev server: `npm run start` (requires `nvm use 22`)

---

## File Structure

| File | Responsibility | Tasks |
|---|---|---|
| `src/webparts/form/components/Form.tsx` | The single-file upload form — all its logic | 1, 2, 3, 4, 6 |
| `src/webparts/bulkUpload/components/BulkUpload.tsx` | Bulk upload web part | 1, 2 (stale-name fixes only) |
| `src/webparts/approvalDocument/components/ApprovalDocument.tsx` | Approver's document review page | 5 |
| `src/shared/formModel.test.ts` | Unit tests for the pure model helpers | 2 |

`Form.tsx` is ~1550 lines. The lint rule `max-lines` caps files at 2000, so these additions are safe. Do not restructure the file beyond what Task 6 specifies.

---

### Task 1: Correct the stale `Vendor` internal name

The `Vendor` column was deleted and replaced by `Vendor/CustomerName`. Two `FIELDS` constants still name the dead column. Nothing fails today because both write paths are disabled — but Task 3 re-enables the Form's, so this must land first.

**Files:**
- Modify: `src/webparts/form/components/Form.tsx:33`
- Modify: `src/webparts/bulkUpload/components/BulkUpload.tsx:55`

- [ ] **Step 1: Update the Form's constant**

In `src/webparts/form/components/Form.tsx`, find line 33 inside the `FIELDS` object:

```ts
  vendor: "Vendor",
```

Replace with:

```ts
  // "Vendor/CustomerName" — the "/" encodes to _x002f_ in the internal name.
  // Verified against /fields 2026-07-28. The old "Vendor" column was deleted.
  vendor: "Vendor_x002f_CustomerName",
```

- [ ] **Step 2: Update the Bulk Upload constant**

In `src/webparts/bulkUpload/components/BulkUpload.tsx`, find line 55 inside its own `FIELDS` object:

```ts
  vendor: "Vendor",
```

Replace with:

```ts
  // "Vendor/CustomerName" — the "/" encodes to _x002f_ in the internal name.
  // Verified against /fields 2026-07-28. The old "Vendor" column was deleted.
  vendor: "Vendor_x002f_CustomerName",
```

- [ ] **Step 3: Verify the build is clean**

Run: `npx heft test`

Expected: `Successes: 57`, `Failures: 0`, and exactly 16 warnings (the pre-existing baseline).

- [ ] **Step 4: Commit**

```bash
git add src/webparts/form/components/Form.tsx src/webparts/bulkUpload/components/BulkUpload.tsx
git commit -m "fix: point vendor column at Vendor_x002f_CustomerName

The old Vendor column was deleted and replaced by Vendor/CustomerName.
Both FIELDS constants named the dead column."
```

---

### Task 2: Rename the Group-led Projects level to "Group Project Name"

The Group-led Projects mode has a folder level called "Project Name" writing to column key `ProjectName`. That key now belongs to the new free-text column. Left alone, a Group-led Projects upload would overwrite the user's typed value with the term label. Rename the level to `GroupProjectName`.

This is a rename in two `DEFAULT_MODES` fallbacks. The columns `GroupProjectName` / `GroupProjectNameTid` do not exist yet and are not needed for the 2026 pilot (Head Office segments only) — this is forward-wiring.

**Files:**
- Test: `src/shared/formModel.test.ts`
- Modify: `src/webparts/form/components/Form.tsx:219`
- Modify: `src/webparts/bulkUpload/components/BulkUpload.tsx:272`

- [ ] **Step 1: Write the test**

In `src/shared/formModel.test.ts`, find the `describe("buildLevelFormValues", ...)` block and add this test inside it:

```ts
  it("resolves a Group Project Name level to its label and Tid columns", () => {
    const columnMap: Record<string, ColumnPair> = {
      GroupProjectName: {
        label: "GroupProjectName",
        tid: "GroupProjectNameTid",
      },
    };
    const result = buildLevelFormValues(columnMap, [
      {
        column: "GroupProjectName",
        label: "Blue Sky",
        id: "3f2a1c8e-0000-0000-0000-000000000001",
      },
    ]);
    expect(result).toEqual([
      { FieldName: "GroupProjectName", FieldValue: "Blue Sky" },
      {
        FieldName: "GroupProjectNameTid",
        FieldValue: "3f2a1c8e-0000-0000-0000-000000000001",
      },
    ]);
  });
```

- [ ] **Step 2: Run the test**

Run: `npx heft test`

Expected: `Successes: 58`, `Failures: 0`.

This test passes immediately — `buildLevelFormValues` is generic over the column map, so it needs no change. That is intentional and is not a reason to skip it: it locks in the new naming and proves the contract the `DEFAULT_MODES` rename depends on. If it *fails*, stop — something about `buildLevelFormValues` is not what this plan assumes.

- [ ] **Step 3: Rename the level in the Form's DEFAULT_MODES**

In `src/webparts/form/components/Form.tsx`, find line 219:

```ts
      { label: "Project Name", column: "ProjectName" },
```

Replace with:

```ts
      // Renamed from "Project Name"/"ProjectName" 2026-07-28: the bare
      // ProjectName column is now the free-text field on every upload.
      // GroupProjectName/GroupProjectNameTid are not created yet — not needed
      // for the 2026 Head Office pilot.
      { label: "Group Project Name", column: "GroupProjectName" },
```

- [ ] **Step 4: Rename the level in Bulk Upload's DEFAULT_MODES**

In `src/webparts/bulkUpload/components/BulkUpload.tsx`, find line 272:

```ts
      { label: "Project Name", column: "ProjectName" },
```

Replace with:

```ts
      // Renamed from "Project Name"/"ProjectName" 2026-07-28: the bare
      // ProjectName column is now the free-text field on every upload.
      { label: "Group Project Name", column: "GroupProjectName" },
```

- [ ] **Step 5: Update the stale example in the parser test**

In `src/shared/formModel.test.ts`, the `parseLevels` test at roughly line 36 uses "Project Name"/`ProjectName` as its example of optional column overrides. That example now describes a mapping this codebase no longer uses. Find:

```ts
  it("carries optional labelCol/tidCol internal-name overrides", () => {
    const json =
      '[{"label":"Project Name","column":"ProjectName","labelCol":"Project_x0020_Name","tidCol":"ProjectNameTid"}]';
    expect(parseLevels(json)).toEqual([
      {
        label: "Project Name",
        column: "ProjectName",
        labelCol: "Project_x0020_Name",
        tidCol: "ProjectNameTid",
      },
```

Replace those lines with:

```ts
  it("carries optional labelCol/tidCol internal-name overrides", () => {
    const json =
      '[{"label":"Group Project Name","column":"GroupProjectName","labelCol":"GroupProjectName","tidCol":"GroupProjectNameTid"}]';
    expect(parseLevels(json)).toEqual([
      {
        label: "Group Project Name",
        column: "GroupProjectName",
        labelCol: "GroupProjectName",
        tidCol: "GroupProjectNameTid",
      },
```

Leave the remaining lines of that test exactly as they are.

- [ ] **Step 6: Run the tests**

Run: `npx heft test`

Expected: `Successes: 58`, `Failures: 0`, 16 warnings.

- [ ] **Step 7: Commit**

```bash
git add src/shared/formModel.test.ts src/webparts/form/components/Form.tsx src/webparts/bulkUpload/components/BulkUpload.tsx
git commit -m "refactor: rename Group-led Projects level to Group Project Name

The bare ProjectName column key now belongs to the new free-text
field. Without this rename a Group-led Projects upload would
overwrite the typed value with the term label."
```

---

### Task 3: Add Project Name state and write both columns

Adds the free-text Project Name field's state and config plumbing, and starts writing both it and Vendor to SharePoint. The visible input is added in Task 6 — this task wires everything behind it.

**Files:**
- Modify: `src/webparts/form/components/Form.tsx` (`FIELDS`, `DmsSettings`, `DEFAULT_SETTINGS`, settings loader, state, `resetForm`, `handleUpload`)

- [ ] **Step 1: Add the column constant**

In `src/webparts/form/components/Form.tsx`, in the `FIELDS` object (around lines 26-35), add after the `vendor` entry:

```ts
  // Free-text project name, written on every upload regardless of segment.
  // Distinct from the Group-led Projects "Group Project Name" folder level.
  projectName: "ProjectName",
```

- [ ] **Step 2: Add it to the settings type**

Find the `DmsSettings` type's `columns` block (around lines 238-246). It currently reads:

```ts
  columns: {
    documentType: string;
    yearPeriod: string;
    documentDate: string;
    confidentiality: string;
    vendor: string;
    businessSegmentLabel: string;
    businessSegmentTid: string;
  };
```

Add `projectName` after `vendor`:

```ts
  columns: {
    documentType: string;
    yearPeriod: string;
    documentDate: string;
    confidentiality: string;
    vendor: string;
    projectName: string;
    businessSegmentLabel: string;
    businessSegmentTid: string;
  };
```

- [ ] **Step 3: Add the default value**

Find `DEFAULT_SETTINGS.columns` (around lines 258-266) and add after the `vendor` line:

```ts
    projectName: FIELDS.projectName,
```

- [ ] **Step 4: Read it from DMS Config**

Find the settings loader at around line 485:

```ts
        vendor: get("col_vendor") ?? DEFAULT_SETTINGS.columns.vendor,
```

Add immediately after it:

```ts
        projectName: get("col_projectName") ?? DEFAULT_SETTINGS.columns.projectName,
```

- [ ] **Step 5: Add the state**

Find the vendor state declaration at around line 299:

```ts
  const [vendor, setVendor] = useState<string>("");
```

Add immediately after it:

```ts
  const [projectName, setProjectName] = useState<string>("");
```

- [ ] **Step 6: Clear it on reset**

Find `resetForm` (around lines 783-794) and add after its `setVendor("");` line:

```ts
    setProjectName("");
```

- [ ] **Step 7: Write both columns on upload**

In `handleUpload`, find this comment block at around lines 1054-1056 and **delete it entirely**:

```ts
      // Vendor is captured as free text only to build the document name — it is NOT
      // written as a metadata column on the file (client request), so no Vendor
      // FieldValue is pushed here.
```

In its place, insert:

```ts
      // Project Name and Vendor are free text. Both are optional, so only push a
      // FieldValue when there is something to write — an empty string would
      // overwrite a stored value rather than leave the column untouched.
      if (projectName.trim()) {
        formValues.push({
          FieldName: settings.columns.projectName,
          FieldValue: projectName.trim(),
        });
      }
      if (vendor.trim()) {
        formValues.push({
          FieldName: settings.columns.vendor,
          FieldValue: vendor.trim(),
        });
      }
```

Do **not** run these values through `ILLEGAL_NAME_CHARS`. That filter exists for filenames; these are metadata values and may legitimately contain `&`, `#`, and so on.

- [ ] **Step 8: Verify the build**

Run: `npx heft test`

Expected: `Successes: 58`, `Failures: 0`, 16 warnings.

TypeScript catches a missed step here — omitting Step 2 or 3 produces a type error on `settings.columns.projectName`.

- [ ] **Step 9: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: write Project Name and Vendor to Staging columns

Both were previously captured on the form and discarded. Vendor's
write was explicitly disabled; Project Name had no field at all.
Values are written only when non-blank."
```

---

### Task 4: Make Document Date optional

The mockup shows no asterisk on Document Date. Drop it from validation and skip its `FieldValue` when blank.

**Critical:** the value must be *omitted*, not sent empty. `toSpDate("")` splits an empty string and produces the malformed string `NaN/NaN/`, which SharePoint rejects with a `HasException`.

**Files:**
- Modify: `src/webparts/form/components/Form.tsx:807`, `:1050`, `:1423-1438`

- [ ] **Step 1: Remove it from validation**

In `handleUpload`, find line 807:

```ts
    if (!documentDate) missing.push("Document Date");
```

Delete that line entirely.

- [ ] **Step 2: Make the FieldValue conditional**

Find line 1050 inside the `formValues` array literal:

```ts
        { FieldName: settings.columns.documentDate, FieldValue: toSpDate(documentDate) },
```

Delete that line from the array. Then, immediately after the array literal's closing `];`, add:

```ts
      // Document Date is optional. Omit rather than send an empty value —
      // toSpDate("") yields the malformed "NaN/NaN/" and SharePoint rejects it.
      if (documentDate) {
        formValues.push({
          FieldName: settings.columns.documentDate,
          FieldValue: toSpDate(documentDate),
        });
      }
```

Place this **before** the Project Name / Vendor pushes added in Task 3, so the write order stays close to the original field order.

- [ ] **Step 3: Drop the asterisk from the label**

Find the Document Date field at around lines 1423-1426:

```tsx
          <label className="dms-field">
            <span>
              Document Date <em className="req">*</em>
            </span>
```

Replace with:

```tsx
          <label className="dms-field">
            <span>Document Date</span>
```

- [ ] **Step 4: Verify the build**

Run: `npx heft test`

Expected: `Successes: 58`, `Failures: 0`, 16 warnings.

- [ ] **Step 5: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: make Document Date optional on single-file upload

Matches the approved mockup. The FieldValue is omitted when blank
rather than sent empty, since toSpDate('') yields NaN/NaN/."
```

---

### Task 5: Show both values on the Approval Document page

**Files:**
- Modify: `src/webparts/approvalDocument/components/ApprovalDocument.tsx:376-382`

- [ ] **Step 1: Add the two rows**

Find the `metadata` array at lines 376-382:

```ts
  const metadata: [string, string][] = [
    ["Location",           orgLocation],
    ["Document Type",      pick("Document_x005f_x0020_x005f_Type", "Document_x0020_Type")],
    ["Confidential Level", pick("Confidentiality_x005f_x0020_x005f_Level", "Confidentiality_x0020_Level")],
    ["Year",               pick("Year", "Year_x005f_x002f_x005f_Period", "Year_x002f_Period")],
    ["Vendor",             pick("Vendor")],
  ];
```

Replace the whole array with:

```ts
  const metadata: [string, string][] = [
    ["Location",           orgLocation],
    ["Document Type",      pick("Document_x005f_x0020_x005f_Type", "Document_x0020_Type")],
    ["Confidential Level", pick("Confidentiality_x005f_x0020_x005f_Level", "Confidentiality_x0020_Level")],
    ["Year",               pick("Year", "Year_x005f_x002f_x005f_Period", "Year_x002f_Period")],
    // ProjectName has no encoded characters, so its response key is unencoded.
    ["Project Name",       pick("ProjectName")],
    // Vendor_x002f_CustomerName does, and FieldValuesAsText double-encodes the
    // underscores in response keys — hence the _x005f_ form first.
    ["Vendor/Customer Name", pick("Vendor_x005f_x002f_x005f_CustomerName", "Vendor_x002f_CustomerName")],
  ];
```

- [ ] **Step 2: Verify the build**

Run: `npx heft test`

Expected: `Successes: 58`, `Failures: 0`, 16 warnings.

- [ ] **Step 3: Commit**

```bash
git add src/webparts/approvalDocument/components/ApprovalDocument.tsx
git commit -m "feat: show Project Name and Vendor on the approval page

The Vendor row existed but always rendered a dash because nothing
wrote the column, and it named the since-deleted Vendor column.
Both rows now resolve."
```

---

### Task 6: Relay out the form to match the mockup

The largest change. Today the form is two cards — **File** (picker + document name) and **Document Information** (folder cascade *above* metadata). The mockup inverts this: document metadata first, folder selection second.

Work carefully. This moves large JSX blocks; a mismatched brace fails the build with a confusing error. Move one block at a time and build between moves if unsure.

**Target:**

```
Card 1 "Upload a Document"
  Document Name              (full width, Max. 30 character)
  Upload Document            (full width file card)
  Project Name | Vendor/Customer Name | Document Date
  Year*        | Document Type*       | Confidential Level*

Card 2 "Document Folder Information"
  Upload to  (radio)
  Segment* / Department* / Unit*   (the level cascade)
```

**Files:**
- Modify: `src/webparts/form/components/Form.tsx:1203-1459`

- [ ] **Step 1: Retitle the first card and put Document Name first**

The first card starts at line 1209 with `<p className="dms-section-title">File</p>`, then renders the file picker `<div className="dms-filecard">`, then Document Name in a `<div style={{ marginTop: 16 }}>`.

Change the title:

```tsx
        <p className="dms-section-title">Upload a Document</p>
```

Move the entire `<div style={{ marginTop: 16 }}>…</div>` wrapper (currently lines 1257-1272) so it sits immediately after that title line, and change its wrapper to `<div style={{ marginBottom: 16 }}>`.

Inside it, the Document Name label currently reads:

```tsx
          <label className="dms-field">
            <span>Document name</span>
            <input
              type="text"
              value={docName}
              disabled={!file}
              placeholder={
                file
                  ? `Leave blank to keep "${file.name}"`
                  : "Select a file first"
              }
              onChange={(e) => onDocNameChange(e.target.value)}
            />
          </label>
```

Replace with:

```tsx
          <label className="dms-field">
            <span>Document Name</span>
            <input
              type="text"
              value={docName}
              maxLength={30}
              disabled={!file}
              placeholder={
                file
                  ? `Leave blank to keep "${file.name}"`
                  : "Select a file first"
              }
              onChange={(e) => onDocNameChange(e.target.value)}
            />
            <small>Max. 30 character</small>
          </label>
```

- [ ] **Step 2: Add the metadata grid to the first card**

Immediately after the file picker's `</div>` closing tag, still inside the first `<div className="dms-section">`, add:

```tsx
        <div className="dms-grid" style={{ marginTop: 16 }}>
          {/* Free-text Project Name — distinct from the Group-led Projects
              "Group Project Name" folder level in the card below. */}
          <label className="dms-field">
            <span>Project Name</span>
            <input
              type="text"
              value={projectName}
              maxLength={30}
              placeholder="Type the project name"
              onChange={(e) => setProjectName(e.target.value)}
            />
            <small>Max. 30 character</small>
          </label>

          {/* Vendor is free text. It also feeds the auto-composed document name. */}
          <label className="dms-field">
            <span>Vendor/Customer Name</span>
            <input
              type="text"
              value={vendor}
              maxLength={30}
              placeholder="Type the vendor or customer name"
              onChange={(e) => onVendorChange(e.target.value)}
            />
            <small>Max. 30 character</small>
          </label>

          <label className="dms-field">
            <span>Document Date</span>
            <input
              type="date"
              value={documentDate}
              max={(() => {
                const d = new Date();
                const mm = d.getMonth() + 1;
                const day = d.getDate();
                return `${d.getFullYear()}-${mm < 10 ? "0" + mm : mm}-${day < 10 ? "0" + day : day}`;
              })()}
              onChange={(e) => onDocumentDateChange(e.target.value)}
            />
          </label>

          {renderSelect("Year", true, yearPeriod, setYearPeriod, options.yearPeriod)}

          {renderSelect(
            "Document Type",
            true,
            documentType,
            setDocumentType,
            options.documentType,
          )}

          {renderSelect(
            "Confidential Level",
            true,
            confidentiality,
            setConfidentiality,
            options.confidentiality,
          )}
        </div>
```

- [ ] **Step 3: Delete the moved fields from the second card**

The second card's `<div className="dms-grid">` (starting line 1363) still holds the originals of the six fields you just recreated. Delete these from it:

- the `renderSelect("Year / Period", …)` call
- the `renderSelect("Document Type", …)` call
- the whole `<label className="dms-field">` for Document Date
- the `renderSelect("Confidentiality Level", …)` call
- the whole `<label className="dms-field">` for Vendor, including its `{/* Vendor is free text… */}` comment

After deletion that grid contains exactly two things: the `{(activeMode()?.levels ?? []).map(…)}` block and the `{(() => { … })()}` empty-level warning.

- [ ] **Step 4: Retitle the second card**

Find line 1277:

```tsx
        <p className="dms-section-title">Document Information</p>
```

Replace with:

```tsx
        <p className="dms-section-title">Document Folder Information</p>
```

- [ ] **Step 5: Relabel the radio group**

Find the radio group's heading at around line 1292:

```tsx
          <p>Upload into:</p>
```

Replace with:

```tsx
          <p>Upload to</p>
```

- [ ] **Step 6: Give the level dropdowns "Select …" placeholders**

In the levels `.map(...)`, the 7th argument to `renderSelect` is the placeholder, currently the literal `"--"`. Find:

```tsx
              deptLoading ||
                isLevelLocked(i) ||
                (i > 0 && !levelValues[i - 1]),
              "--",
              true,
```

Replace with:

```tsx
              deptLoading ||
                isLevelLocked(i) ||
                (i > 0 && !levelValues[i - 1]),
              `Select ${lvl.label}`,
              true,
```

This yields "Select Segment", "Select Department", "Select Unit" automatically from each level's label.

- [ ] **Step 7: Verify the build**

Run: `npx heft test`

Expected: `Successes: 58`, `Failures: 0`, 16 warnings.

If you see a `max-lines` warning for `Form.tsx`, you duplicated a block instead of moving it — the file should stay near 1550 lines.

- [ ] **Step 8: Visually verify against the mockup**

Run: `nvm use 22 && npm run start`

Open the workbench:
`https://dcidigitalcom.sharepoint.com/_layouts/workbench.aspx?debugManifestsFile=https://localhost:4321/temp/build/manifests.js&debug=true&noredir=true`

Confirm: two cards in the right order, Document Name above the upload area, two three-column rows, asterisks on exactly Year / Document Type / Confidential Level and the folder levels, and "Max. 30 character" under the three text fields.

- [ ] **Step 9: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: relay out single-file upload form to match mockup

Document metadata now precedes folder selection, and the two cards
are retitled Upload a Document / Document Folder Information.
Adds 30-char limits and hints. Buttons deliberately unchanged."
```

---

### Task 7: End-to-end verification

No code. This is the only place the round-trip is proven, and — because both columns are hidden from the Staging views — the Approval page is the *only* surface where a wrong value is visible. Do not skip.

**Prerequisites** (outside the codebase; confirm before starting):
- `ProjectName` and `Vendor/CustomerName` exist in Staging — true as of 2026-07-28.
- `DMS Config` has `col_projectName` = `ProjectName`.
- `DMS Config` row `col_vendor` = `Vendor_x002f_CustomerName`, **not** the deleted `Vendor`. A stale row here overrides the code fallback and breaks every upload.

- [ ] **Step 1: Full production build**

Run: `npm run build`

Expected: completes with no errors; warnings still number 16.

- [ ] **Step 2: Happy path**

Upload a file to a Business Segment with Project Name = `Blue Sky Refinery`, Vendor = `Acme Trading`, and a Document Date set.

Expected: success toast. The Approval Document page shows both new rows with those values, and the filename is `Acme Trading-<DD-MM-YY>.<ext>`.

- [ ] **Step 3: Both fields blank**

Upload with Project Name and Vendor empty.

Expected: succeeds with no field-error toast; the Approval page shows "—" for both rows.

- [ ] **Step 4: Blank Document Date**

Upload with Document Date empty and Vendor = `Acme Trading`.

Expected: succeeds — no "Uploaded, but a field failed" toast, and specifically no `NaN/NaN/` error. Filename is `Acme Trading.<ext>`.

- [ ] **Step 5: Character limits**

Try typing 31+ characters into Document Name, Project Name, and Vendor.

Expected: each stops accepting input at 30.

- [ ] **Step 6: Group-led Projects**

Switch to the Project side, if a mode is offered to your account.

Expected: both "Project Name" (text, card 1) and "Group Project Name" (dropdown, card 2) are visible and independent.

- [ ] **Step 7: Restricted user**

Sign in as a non-privileged uploader.

Expected: the cascade shows only their authorised paths, and the "Uploading to:" breadcrumb still resolves — both now inside the second card.

- [ ] **Step 8: Regression sweep of the moved card**

Confirm these still render inside the folder card: the "Loading your access…" state, the not-provisioned error, the segment radio group, and the amber empty-level warning.

- [ ] **Step 9: Bulk upload untouched**

Open the Bulk Upload web part and run one upload.

Expected: works exactly as before. Its Vendor field stays hidden; only its internal-name constant changed.

---

## Follow-ups (not in this plan)

- Correct the `2026-07-28-dms-uat-test-plan.md` §2.1 fixture: `Document_x0020_Type` and `Confidentiality_x0020_Level` are `TaxonomyFieldType` not Text, and the Year column's internal name is `Year` not `Year_x002f_Period`.
- Amend `2026-07-21-segment-onboarding-plan.md` and `2026-07-16-tenant-seed-data-runbook.md` for the `GroupProjectName` rename.
- Bulk Upload writes Vendor as managed metadata (`"Label|GUID"`) but the column is Text. Dormant while the field is hidden; a latent bug if it is ever re-enabled.
- Documents library parity: if it has its own Vendor column, the auto-route flow's mapping needs the same rename or approved files lose the value on the way across.

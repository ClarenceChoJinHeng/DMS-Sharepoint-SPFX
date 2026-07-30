# Allowed File Types Dropdown Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the `AllowedFileTypes` multi-select Choice column on `DMS Config` the single source of truth for which file types may be uploaded, so a mistyped extension can no longer silently disable uploads.

**Architecture:** A new `src/shared/allowedFileTypes.ts` normalizes extensions and resolves them into one of three explicit states — `configured`, `none`, `unknown`. `Form.tsx` and `BulkUpload.tsx` read the Choice column (retrying without the field for sites that lack it), then render each state distinctly: normal picker, hard block with a client-actionable message, or code defaults with an admin warning. `SettingValue` is no longer read for this setting.

**Tech Stack:** SPFx 1.23.0, React 17, TypeScript 5.8, Heft + Jest (`@rushstack/heft-jest-plugin`), SharePoint REST via `SPHttpClient`.

**Spec:** `docs/superpowers/specs/2026-07-30-allowed-file-types-dropdown-design.md`

---

## Background the engineer needs

**What went wrong originally.** `allowedExtensions` was a hand-typed comma string. Someone entered `png` without a leading dot. That string is passed straight into `<input type="file" accept="...">`. Browsers only treat an `accept` token as a file extension when it starts with `.`; a bare `png` is parsed as a MIME type, has no `/`, is invalid, and is dropped. PNG files were greyed out in the file dialog. Neither web part has a drag-and-drop handler, so the file input is the only way in — PNG became unuploadable, with no error anywhere.

**Two project-specific constraints:**

1. **Do not use `Array.prototype.includes`, `Promise.allSettled`, or other ES2017+ library methods.** The SPFx tsconfig does not target them (CLAUDE.md gotcha #3). Use `indexOf` and per-item `try/catch`.
2. **Multi-value Choice fields have two possible JSON shapes.** With `odata=nometadata` or `minimalmetadata` SharePoint returns a plain array `[".pdf", ".doc"]`. With `odata=verbose` it returns `{ "results": [".pdf", ".doc"] }`. The current `Accept` header is a bare `application/json`, which leaves this ambiguous. Task 2 handles both shapes and Task 4 pins the header explicitly — belt and braces, because getting this wrong produces `undefined`, which resolves to `unknown`, which silently falls back to defaults. Exactly the failure mode this whole change exists to remove.

**Running tests:** there is no `npm test` script. Use `npx heft test`. Filter with `--test-path-pattern`. Heft type-checks the whole project during `test`, so a test referencing a not-yet-written module fails as a **compile error** — that is a legitimate red for TDD purposes.

---

## File Structure

| File | Responsibility |
|---|---|
| `src/shared/allowedFileTypes.ts` | **new.** Normalization, tolerant Choice-array reading, and resolution into the three states. Pure functions, no SharePoint dependency, unit-testable. |
| `src/shared/allowedFileTypes.test.ts` | **new.** Unit tests for the above. |
| `src/webparts/form/components/Form.tsx` | Reads the column, renders the three states in the single-file upload card. |
| `src/webparts/bulkUpload/components/BulkUpload.tsx` | Same, for the multi-file batch picker. |

`FolderManager.tsx` reads `SettingValue` but never `allowedExtensions`, so it is not touched.

---

## Task 1: Normalizer and state type

**Files:**
- Create: `src/shared/allowedFileTypes.ts`
- Test: `src/shared/allowedFileTypes.test.ts`

- [ ] **Step 1: Write the failing test**

Create `src/shared/allowedFileTypes.test.ts`:

```ts
import { normalizeFileTypes } from "./allowedFileTypes";

describe("normalizeFileTypes", () => {
  // Regression: a bare "png" typed into DMS Config produced an invalid `accept`
  // token, greying PNG out of the file picker with no error. 2026-07-30.
  it("prepends a missing leading dot", () => {
    expect(normalizeFileTypes(["png"])).toEqual([".png"]);
  });

  it("lowercases", () => {
    expect(normalizeFileTypes([".PNG", ".Docx"])).toEqual([".png", ".docx"]);
  });

  it("trims surrounding whitespace", () => {
    expect(normalizeFileTypes(["  .pdf  "])).toEqual([".pdf"]);
  });

  it("drops empty and whitespace-only entries", () => {
    expect(normalizeFileTypes([".pdf", "", "   ", ".doc"])).toEqual([".pdf", ".doc"]);
  });

  it("de-duplicates, including entries that only differ by dot or case", () => {
    expect(normalizeFileTypes([".pdf", "pdf", ".PDF"])).toEqual([".pdf"]);
  });

  it("preserves the order the values arrived in", () => {
    expect(normalizeFileTypes([".xls", ".doc", ".pdf"])).toEqual([".xls", ".doc", ".pdf"]);
  });

  it("returns an empty array for an empty input", () => {
    expect(normalizeFileTypes([])).toEqual([]);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx heft test --test-path-pattern allowedFileTypes`

Expected: FAIL — TypeScript cannot resolve the module `./allowedFileTypes`.

- [ ] **Step 3: Write the minimal implementation**

Create `src/shared/allowedFileTypes.ts`:

```ts
/**
 * Allowed upload file types, resolved from the `AllowedFileTypes` multi-select
 * Choice column on the `DMS Config` list.
 *
 * This is a discriminated union rather than a bare `string[]` on purpose. Three
 * outcomes have to stay distinguishable, because different people fix them:
 *
 *   configured — the client ticked types. Normal operation.
 *   none       — the client unticked everything. Hard block; the CLIENT fixes it.
 *   unknown    — the config could not be read, or the column does not exist on
 *                this site. Fall back to `FALLBACK_FILE_TYPES`; an ADMIN fixes it.
 *
 * Collapsing `none` and `unknown` into an empty array is the specific mistake
 * this type exists to prevent. See
 * docs/superpowers/specs/2026-07-30-allowed-file-types-dropdown-design.md §3, §6.
 */
export type AllowedFileTypes =
  | { kind: "configured"; types: string[] }
  | { kind: "none" }
  | { kind: "unknown"; types: string[] };

/**
 * Last-resort list used only when `DMS Config` cannot be read at all. Kept equal
 * to the column's configured choices so that even the fallback agrees with the
 * dropdown (spec §7).
 */
export const FALLBACK_FILE_TYPES: string[] = [
  ".pdf",
  ".doc",
  ".docx",
  ".xls",
  ".xlsx",
];

/**
 * trim -> lowercase -> prepend "." if missing -> drop empties -> de-duplicate.
 *
 * Defence in depth. The dropdown removes typos from the common path, but an
 * owner editing the column's *Choices* still types the extension by hand, so a
 * fumbled dot is repaired here instead of silently breaking the picker.
 */
export function normalizeFileTypes(raw: readonly string[]): string[] {
  const out: string[] = [];
  raw.forEach((entry) => {
    const trimmed = (entry ?? "").trim().toLowerCase();
    if (trimmed.length === 0) return;
    const dotted = trimmed.charAt(0) === "." ? trimmed : `.${trimmed}`;
    // indexOf, not includes: the SPFx tsconfig does not target ES2017 libs
    // (CLAUDE.md gotcha #3).
    if (out.indexOf(dotted) === -1) out.push(dotted);
  });
  return out;
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx heft test --test-path-pattern allowedFileTypes`

Expected: PASS, 7 tests.

- [ ] **Step 5: Commit**

```bash
git add src/shared/allowedFileTypes.ts src/shared/allowedFileTypes.test.ts
git commit -m "feat: add file-type normalizer with explicit resolved states"
```

---

## Task 2: Tolerant Choice-array reader

SharePoint returns a multi-value Choice field as a plain array under `nometadata`/`minimalmetadata`, but as `{ results: [...] }` under `verbose`. Read both, so a future header change cannot silently turn every site into `unknown`.

**Files:**
- Modify: `src/shared/allowedFileTypes.ts`
- Test: `src/shared/allowedFileTypes.test.ts`

- [ ] **Step 1: Write the failing test**

Append to `src/shared/allowedFileTypes.test.ts`, and add `readChoiceArray` to the existing import at the top of the file:

```ts
describe("readChoiceArray", () => {
  it("reads the plain array shape (odata=nometadata / minimalmetadata)", () => {
    expect(readChoiceArray([".pdf", ".doc"])).toEqual([".pdf", ".doc"]);
  });

  it("reads the wrapped shape (odata=verbose)", () => {
    expect(readChoiceArray({ results: [".pdf", ".doc"] })).toEqual([".pdf", ".doc"]);
  });

  it("reads an empty selection in both shapes as an empty array, NOT as absent", () => {
    expect(readChoiceArray([])).toEqual([]);
    expect(readChoiceArray({ results: [] })).toEqual([]);
  });

  it("returns undefined when the field is absent", () => {
    expect(readChoiceArray(undefined)).toBeUndefined();
    expect(readChoiceArray(null)).toBeUndefined();
  });

  it("returns undefined for shapes that are not a choice array", () => {
    expect(readChoiceArray(".pdf,.doc")).toBeUndefined();
    expect(readChoiceArray({ notResults: [".pdf"] })).toBeUndefined();
  });
});
```

The empty-versus-absent distinction is the important one: `[]` must mean "the client unticked everything" (→ `none`) while `undefined` must mean "no such column here" (→ `unknown`).

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx heft test --test-path-pattern allowedFileTypes`

Expected: FAIL — `readChoiceArray` is not exported from `./allowedFileTypes`.

- [ ] **Step 3: Write the minimal implementation**

Append to `src/shared/allowedFileTypes.ts`:

```ts
/**
 * Reads a SharePoint multi-value Choice field, tolerating both JSON shapes:
 * a plain array (odata=nometadata / minimalmetadata) and `{ results: [...] }`
 * (odata=verbose).
 *
 * Returns `undefined` only when the field is genuinely absent. An empty
 * selection returns `[]`, which callers must treat as "nothing ticked" rather
 * than "not configured".
 */
export function readChoiceArray(value: unknown): string[] | undefined {
  if (Array.isArray(value)) return value as string[];
  if (value !== null && typeof value === "object") {
    const wrapped = (value as { results?: unknown }).results;
    if (Array.isArray(wrapped)) return wrapped as string[];
  }
  return undefined;
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx heft test --test-path-pattern allowedFileTypes`

Expected: PASS, 12 tests.

- [ ] **Step 5: Commit**

```bash
git add src/shared/allowedFileTypes.ts src/shared/allowedFileTypes.test.ts
git commit -m "feat: read multi-choice fields in both odata shapes"
```

---

## Task 3: Resolver

**Files:**
- Modify: `src/shared/allowedFileTypes.ts`
- Test: `src/shared/allowedFileTypes.test.ts`

- [ ] **Step 1: Write the failing test**

Append to `src/shared/allowedFileTypes.test.ts`, adding `resolveAllowedFileTypes`, `FALLBACK_FILE_TYPES`, `NO_TYPES_MESSAGE`, and `CONFIG_UNREADABLE_MESSAGE` to the import at the top:

```ts
describe("resolveAllowedFileTypes", () => {
  it("resolves a populated selection to configured, normalized", () => {
    expect(resolveAllowedFileTypes(["png", ".PDF"])).toEqual({
      kind: "configured",
      types: [".png", ".pdf"],
    });
  });

  // Spec §3: empty is a hard block, never a silent fallback and never an empty
  // allowlist that some caller might read as permissive.
  it("resolves an empty selection to none", () => {
    expect(resolveAllowedFileTypes([])).toEqual({ kind: "none" });
  });

  it("resolves a selection of only blanks to none", () => {
    expect(resolveAllowedFileTypes(["", "  "])).toEqual({ kind: "none" });
  });

  it("resolves an absent field to unknown, carrying the fallback types", () => {
    expect(resolveAllowedFileTypes(undefined)).toEqual({
      kind: "unknown",
      types: FALLBACK_FILE_TYPES,
    });
  });

  // The two failure states must never be confusable — different people fix them.
  it("distinguishes none from unknown", () => {
    expect(resolveAllowedFileTypes([]).kind).not.toBe(
      resolveAllowedFileTypes(undefined).kind,
    );
  });

  it("never reports configured with an empty type list", () => {
    const resolved = resolveAllowedFileTypes([""]);
    expect(resolved.kind).toBe("none");
    expect(resolved).not.toHaveProperty("types");
  });

  it("ships fallback types that are already normalized", () => {
    expect(normalizeFileTypes(FALLBACK_FILE_TYPES)).toEqual(FALLBACK_FILE_TYPES);
  });
});

describe("messages", () => {
  it("tells the client which column and list to fix", () => {
    expect(NO_TYPES_MESSAGE).toContain("Allowed File Types");
    expect(NO_TYPES_MESSAGE).toContain("DMS Config");
  });

  it("uses a different message for an unreadable config", () => {
    expect(CONFIG_UNREADABLE_MESSAGE).not.toBe(NO_TYPES_MESSAGE);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx heft test --test-path-pattern allowedFileTypes`

Expected: FAIL — `resolveAllowedFileTypes`, `NO_TYPES_MESSAGE`, and `CONFIG_UNREADABLE_MESSAGE` are not exported.

- [ ] **Step 3: Write the minimal implementation**

Append to `src/shared/allowedFileTypes.ts`:

```ts
/**
 * Shown when `AllowedFileTypes` is present but nothing is ticked. Names the
 * column and the list because the client is the person who fixes it, in one
 * click, and a vague message would send them to us instead.
 */
export const NO_TYPES_MESSAGE =
  "No file types are configured for upload. Ask your DMS administrator to set Allowed File Types in DMS Config.";

/**
 * Shown when `DMS Config` could not be read at all. Deliberately different from
 * NO_TYPES_MESSAGE: this one is an admin or permissions problem, and uploads
 * continue on built-in defaults rather than stopping.
 */
export const CONFIG_UNREADABLE_MESSAGE =
  "Could not load the upload configuration — using built-in file types. Ask your DMS administrator to check DMS Config.";

/**
 * Turns the raw `AllowedFileTypes` value into one of the three states.
 *
 * `undefined` (field absent — unprovisioned site, or a read that fell back)
 * becomes `unknown`. An empty or all-blank selection becomes `none`.
 */
export function resolveAllowedFileTypes(
  raw: readonly string[] | undefined,
): AllowedFileTypes {
  if (raw === undefined) {
    return { kind: "unknown", types: normalizeFileTypes(FALLBACK_FILE_TYPES) };
  }
  const types = normalizeFileTypes(raw);
  if (types.length === 0) return { kind: "none" };
  return { kind: "configured", types };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx heft test --test-path-pattern allowedFileTypes`

Expected: PASS, 21 tests.

- [ ] **Step 5: Commit**

```bash
git add src/shared/allowedFileTypes.ts src/shared/allowedFileTypes.test.ts
git commit -m "feat: resolve allowed file types into configured/none/unknown"
```

---

## Task 4: Wire `Form.tsx` settings read

**Files:**
- Modify: `src/webparts/form/components/Form.tsx` — imports, `DmsSettings` (line 261), `DEFAULT_SETTINGS` (line 281), `loadSettings` (lines 467-510), `init` (lines 657-664, 678)

- [ ] **Step 1: Add the import**

Add to the import block at the top of `Form.tsx`:

```ts
import {
  AllowedFileTypes,
  CONFIG_UNREADABLE_MESSAGE,
  FALLBACK_FILE_TYPES,
  readChoiceArray,
  resolveAllowedFileTypes,
} from "../../../shared/allowedFileTypes";
```

Verify the relative depth against an existing shared import in the same file (for example the one for `pathEncoding`) and match it rather than trusting the path above.

- [ ] **Step 2: Change the settings type**

In the `DmsSettings` type, replace:

```ts
  allowedExtensions: string[];
```

with:

```ts
  allowedFileTypes: AllowedFileTypes;
```

- [ ] **Step 3: Change the default**

In `DEFAULT_SETTINGS`, replace:

```ts
  allowedExtensions: [".pdf", ".xls", ".xlsx"],
```

with:

```ts
  // "unknown", not "configured": reaching this constant means DMS Config could
  // not be read, and the UI must say so rather than present these as configured
  // values. The old list here was [".pdf", ".xls", ".xlsx"] — missing .doc/.docx,
  // which is part of why the silent fallback was so confusing to diagnose.
  allowedFileTypes: { kind: "unknown", types: FALLBACK_FILE_TYPES },
```

- [ ] **Step 4: Replace `loadSettings`**

Replace the whole body of `loadSettings` (lines 467-510) with:

```ts
  // Two field lists: a site whose DMS Config predates the AllowedFileTypes column
  // answers HTTP 400 to the entire request, which would drop EVERY setting on it —
  // term-set GUIDs, stagingLibrary, column names — not just the file types. So we
  // retry without the new field. Spec 2026-07-30 §5.
  const SETTINGS_FIELDS = "Title,SettingValue,AllowedFileTypes";
  const SETTINGS_FIELDS_LEGACY = "Title,SettingValue";

  const fetchSettingRows = async (
    select: string,
  ): Promise<SPHttpClientResponse> =>
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=${select}&$filter=ConfigType eq 'setting'`,
      SPHttpClient.configurations.v1,
      // Pinned to nometadata so multi-choice fields arrive as a plain array.
      { headers: { Accept: "application/json;odata=nometadata" } },
    );

  const loadSettings = async (): Promise<DmsSettings> => {
    let res = await fetchSettingRows(SETTINGS_FIELDS);
    if (!res.ok) {
      const body = await res.text();
      console.warn(
        `DMS Config read including AllowedFileTypes failed (HTTP ${res.status}). ` +
          `Retrying without that field. Response: ${body}`,
      );
      res = await fetchSettingRows(SETTINGS_FIELDS_LEGACY);
    }
    if (!res.ok) {
      const body = await res.text();
      throw new Error(
        `DMS Config read failed: HTTP ${res.status}. Response: ${body}`,
      );
    }
    const data = await res.json();
    const map: Record<string, string> = {};
    // Left undefined when the column is absent, which resolves to "unknown".
    // An empty tick list arrives as [] and resolves to "none" — a different state.
    let rawFileTypes: string[] | undefined;
    (data.value ?? []).forEach(
      (item: {
        Title: string;
        SettingValue: string;
        AllowedFileTypes?: unknown;
      }) => {
        map[item.Title] = item.SettingValue;
        if (item.Title === "allowedExtensions") {
          rawFileTypes = readChoiceArray(item.AllowedFileTypes);
        }
      },
    );
    const get = (key: string): string | undefined => map[key];
    return {
      termSets: {
        documentType:
          get("termSet_documentType") ?? DEFAULT_SETTINGS.termSets.documentType,
        yearPeriod:
          get("termSet_yearPeriod") ?? DEFAULT_SETTINGS.termSets.yearPeriod,
        confidentiality:
          get("termSet_confidentiality") ??
          DEFAULT_SETTINGS.termSets.confidentiality,
      },
      columns: {
        documentType: get("col_documentType") ?? DEFAULT_SETTINGS.columns.documentType,
        yearPeriod: get("col_yearPeriod") ?? DEFAULT_SETTINGS.columns.yearPeriod,
        documentDate: get("col_documentDate") ?? DEFAULT_SETTINGS.columns.documentDate,
        confidentiality:
          get("col_confidentiality") ?? DEFAULT_SETTINGS.columns.confidentiality,
        vendor: get("col_vendor") ?? DEFAULT_SETTINGS.columns.vendor,
        projectName: get("col_projectName") ?? DEFAULT_SETTINGS.columns.projectName,
        businessSegmentLabel:
          get("col_businessSegment") ?? DEFAULT_SETTINGS.columns.businessSegmentLabel,
        businessSegmentTid:
          get("col_businessSegmentTid") ?? DEFAULT_SETTINGS.columns.businessSegmentTid,
      },
      stagingLibrary: get("stagingLibrary") ?? DEFAULT_SETTINGS.stagingLibrary,
      // SettingValue is deliberately NOT consulted for file types any more —
      // AllowedFileTypes is the single source of truth. Spec 2026-07-30 §3.
      allowedFileTypes: resolveAllowedFileTypes(rawFileTypes),
    };
  };
```

Note `fetchSettingRows` and the two field constants go **inside** the component, alongside `loadSettings`, because they close over `context` and `siteUrl`.

- [ ] **Step 5: Make the init fallbacks loud**

In `init`, replace lines 659-661:

```ts
          loadSettings().catch(() => DEFAULT_SETTINGS),
          loadModes().catch(() => DEFAULT_MODES),
          loadGroupMap().catch(() => [] as GroupMapRow[]),
```

with:

```ts
          // Log the real failure. A silent fallback here is indistinguishable
          // from success and cost four rounds of diagnosis on 2026-07-30.
          // CLAUDE.md gotcha #9.
          loadSettings().catch((err) => {
            console.error("DMS Config settings read failed — using built-in defaults.", err);
            return DEFAULT_SETTINGS;
          }),
          loadModes().catch((err) => {
            console.error("DMS Config mode rows read failed — using built-in modes.", err);
            return DEFAULT_MODES;
          }),
          loadGroupMap().catch((err) => {
            console.error("DMS Group Map read failed — no authorised upload paths.", err);
            return [] as GroupMapRow[];
          }),
```

- [ ] **Step 6: Warn admins when the config could not be read**

In `init`, immediately after `setPrivileged(isPrivileged);` (line 678), insert:

```ts
      // "unknown" means the config could not be read, or this site has no
      // AllowedFileTypes column — an admin problem, not something an uploader can
      // fix, so the toast is admin-only while the console line is always written.
      // Distinct from the "none" state, which the file card handles. Spec §6.
      if (loadedSettings.allowedFileTypes.kind === "unknown") {
        console.warn(
          "AllowedFileTypes not supplied by DMS Config — running on built-in types:",
          loadedSettings.allowedFileTypes.types.join(", "),
        );
        if (isPrivileged) showToast(CONFIG_UNREADABLE_MESSAGE, "error");
      }
```

- [ ] **Step 7: Verify it compiles**

Run: `npx heft test`

Expected: FAIL, with type errors at the three remaining `settings.allowedExtensions` usages in `Form.tsx` (around lines 828, 1299, 1305). Task 5 fixes those. Confirm the errors name only those lines — anything else means a mistake above.

- [ ] **Step 8: Do not commit yet**

The build is intentionally red until Task 5. Commit at the end of Task 5.

---

## Task 5: Wire `Form.tsx` UI

**Files:**
- Modify: `src/webparts/form/components/Form.tsx` — submit guard (lines 826-837), file card (lines 1277-1321)

- [ ] **Step 1: Add the import**

Add `NO_TYPES_MESSAGE` to the `allowedFileTypes` import added in Task 4.

- [ ] **Step 2: Replace the submit-time guard**

Replace lines 826-837:

```ts
    const finalName = buildUploadName(file.name, docName);
    if (
      !settings.allowedExtensions.some((ext) =>
        finalName.toLowerCase().endsWith(ext),
      )
    ) {
      showToast(
        `File type not allowed. Allowed: ${settings.allowedExtensions.join(", ")}`,
        "error",
      );
      return;
    }
```

with:

```ts
    const finalName = buildUploadName(file.name, docName);
    const allowed = settings.allowedFileTypes;
    if (allowed.kind === "none") {
      showToast(NO_TYPES_MESSAGE, "error");
      return;
    }
    if (!allowed.types.some((ext) => finalName.toLowerCase().endsWith(ext))) {
      showToast(
        `File type not allowed. Allowed: ${allowed.types.join(", ")}`,
        "error",
      );
      return;
    }
```

- [ ] **Step 3: Replace the file card**

Replace lines 1277-1321 (from the `{/* The whole card is clickable... */}` comment through the `</div>` that closes `dms-filecard`) with:

```tsx
        {/* An empty AllowedFileTypes selection is a hard block, not a silent
            fallback — spec 2026-07-30 §3. The message names the column and list
            because the client is the one who fixes it. */}
        {settings.allowedFileTypes.kind === "none" ? (
          <div className="dms-filecard" style={{ opacity: 0.6 }}>
            <span>{NO_TYPES_MESSAGE}</span>
          </div>
        ) : (
          /* The whole card is clickable to open the file picker. */
          <div
            className="dms-filecard"
            role="button"
            tabIndex={0}
            style={{ cursor: "pointer" }}
            onClick={() => fileRef.current?.click()}
            onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") fileRef.current?.click(); }}
          >
            {/* Action label sits first so it reads left-aligned in the card. */}
            <span className="dms-link">
              {file ? "Change file" : "Upload Document"}
            </span>
            {file && (
              <div>
                <div className="name">{file.name}</div>
                <div className="size">{(file.size / 1024).toFixed(1)} KB</div>
              </div>
            )}
            <input
              ref={fileRef}
              type="file"
              accept={settings.allowedFileTypes.types.join(",")}
              style={{ display: "none" }}
              onChange={(e) => {
                const picked = e.target.files?.[0];
                const types =
                  settings.allowedFileTypes.kind === "none"
                    ? []
                    : settings.allowedFileTypes.types;
                if (
                  picked &&
                  !types.some((ext) => picked.name.toLowerCase().endsWith(ext))
                ) {
                  showToast(
                    `File type not allowed. Allowed: ${types.join(", ")}`,
                    "error",
                  );
                  setFile(undefined);
                  if (fileRef.current) fileRef.current.value = "";
                  return;
                }
                setStatus("");
                setFile(picked);
              }}
            />
          </div>
        )}
```

TypeScript narrows `settings.allowedFileTypes` to `configured | unknown` inside the false branch, so `.types` is valid on the `accept` line. The `onChange` handler re-checks `kind` because narrowing does not reach inside the callback.

- [ ] **Step 4: Verify the build is green**

Run: `npx heft test`

Expected: PASS. No `allowedExtensions` references remain in `Form.tsx` — confirm with:

`git grep -n "allowedExtensions" -- src/webparts/form`

Expected: no output.

- [ ] **Step 5: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: read allowed file types from the AllowedFileTypes column"
```

---

## Task 6: Wire `BulkUpload.tsx`

`BulkUpload` has no drop handler either — every file arrives through `addFiles`, so that function is the real gate. The picker is not restructured here; the block lives in `addFiles`, which every path crosses.

**Files:**
- Modify: `src/webparts/bulkUpload/components/BulkUpload.tsx` — imports, `DmsSettings` (line 304), `DEFAULT_SETTINGS` (line 324), `loadSettings` (lines 642-685), `addFiles` (lines 1119-1137), file input (line 2155)

- [ ] **Step 1: Add the import**

```ts
import {
  AllowedFileTypes,
  FALLBACK_FILE_TYPES,
  NO_TYPES_MESSAGE,
  readChoiceArray,
  resolveAllowedFileTypes,
} from "../../../shared/allowedFileTypes";
```

Match the relative depth to an existing shared import in this file.

- [ ] **Step 2: Change the type and default**

In `DmsSettings` (line 304), replace `allowedExtensions: string[];` with `allowedFileTypes: AllowedFileTypes;`.

In `DEFAULT_SETTINGS` (line 324), replace `allowedExtensions: [".pdf", ".xls", ".xlsx"],` with:

```ts
  allowedFileTypes: { kind: "unknown", types: FALLBACK_FILE_TYPES },
```

- [ ] **Step 3: Update `loadSettings`**

Replace lines 642-655 — the request, the `!res.ok` throw, and the `forEach` that builds `map`:

```ts
  const loadSettings = async (): Promise<DmsSettings> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error("DMS Config list not found");
    const data = await res.json();
    const map: Record<string, string> = {};
    (data.value ?? []).forEach(
      (item: { Title: string; SettingValue: string }) => {
        map[item.Title] = item.SettingValue;
      },
    );
```

with:

```ts
  // Two field lists: a site whose DMS Config predates the AllowedFileTypes column
  // answers HTTP 400 to the entire request, which would drop EVERY setting on it,
  // not just the file types. So we retry without the new field. Spec §5.
  const SETTINGS_FIELDS = "Title,SettingValue,AllowedFileTypes";
  const SETTINGS_FIELDS_LEGACY = "Title,SettingValue";

  const fetchSettingRows = async (
    select: string,
  ): Promise<SPHttpClientResponse> =>
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=${select}&$filter=ConfigType eq 'setting'`,
      SPHttpClient.configurations.v1,
      // Pinned to nometadata so multi-choice fields arrive as a plain array.
      { headers: { Accept: "application/json;odata=nometadata" } },
    );

  const loadSettings = async (): Promise<DmsSettings> => {
    let res = await fetchSettingRows(SETTINGS_FIELDS);
    if (!res.ok) {
      const body = await res.text();
      console.warn(
        `DMS Config read including AllowedFileTypes failed (HTTP ${res.status}). ` +
          `Retrying without that field. Response: ${body}`,
      );
      res = await fetchSettingRows(SETTINGS_FIELDS_LEGACY);
    }
    if (!res.ok) {
      const body = await res.text();
      throw new Error(
        `DMS Config read failed: HTTP ${res.status}. Response: ${body}`,
      );
    }
    const data = await res.json();
    const map: Record<string, string> = {};
    // Left undefined when the column is absent, which resolves to "unknown".
    // An empty tick list arrives as [] and resolves to "none" — a different state.
    let rawFileTypes: string[] | undefined;
    (data.value ?? []).forEach(
      (item: {
        Title: string;
        SettingValue: string;
        AllowedFileTypes?: unknown;
      }) => {
        map[item.Title] = item.SettingValue;
        if (item.Title === "allowedExtensions") {
          rawFileTypes = readChoiceArray(item.AllowedFileTypes);
        }
      },
    );
```

`fetchSettingRows` and the two constants go **inside** the component, immediately above `loadSettings`, because they close over `context` and `siteUrl`.

Then replace the returned `allowedExtensions` property (lines 681-683):

```ts
      allowedExtensions: get("allowedExtensions")
        ? (get("allowedExtensions") as string).split(",").map((e) => e.trim())
        : DEFAULT_SETTINGS.allowedExtensions,
```

with:

```ts
      // SettingValue is deliberately NOT consulted for file types any more —
      // AllowedFileTypes is the single source of truth. Spec 2026-07-30 §3.
      allowedFileTypes: resolveAllowedFileTypes(rawFileTypes),
```

Leave every other property in this return untouched — `BulkUpload`'s `DmsSettings` carries a
`termSets.vendor` that `Form.tsx` does not, and it is out of scope here.

- [ ] **Step 4: Update `addFiles`**

Replace lines 1119-1137:

```ts
  const addFiles = (list: FileList | null): void => {
    if (!list || list.length === 0) return;
    const incoming = Array.from(list);
    const allowed: File[] = [];
    let rejectedCount = 0;
    incoming.forEach((f) => {
      const ok = settings.allowedExtensions.some((ext) =>
        f.name.toLowerCase().endsWith(ext),
      );
      if (ok) allowed.push(f);
      else rejectedCount++;
    });

    if (rejectedCount > 0) {
      showToast(
        `Skipped ${rejectedCount} file(s) with a disallowed type. Allowed: ${settings.allowedExtensions.join(", ")}`,
        "error",
      );
    }
```

with:

```ts
  const addFiles = (list: FileList | null): void => {
    if (!list || list.length === 0) return;
    // Every file in this web part arrives through addFiles, so the hard block for
    // an empty AllowedFileTypes selection lives here. Spec 2026-07-30 §3.
    const allowedTypes = settings.allowedFileTypes;
    if (allowedTypes.kind === "none") {
      showToast(NO_TYPES_MESSAGE, "error");
      return;
    }
    const incoming = Array.from(list);
    const allowed: File[] = [];
    let rejectedCount = 0;
    incoming.forEach((f) => {
      const ok = allowedTypes.types.some((ext) =>
        f.name.toLowerCase().endsWith(ext),
      );
      if (ok) allowed.push(f);
      else rejectedCount++;
    });

    if (rejectedCount > 0) {
      showToast(
        `Skipped ${rejectedCount} file(s) with a disallowed type. Allowed: ${allowedTypes.types.join(", ")}`,
        "error",
      );
    }
```

- [ ] **Step 5: Update the file input**

Replace line 2155:

```tsx
                accept={settings.allowedExtensions.join(",")}
```

with:

```tsx
                accept={
                  settings.allowedFileTypes.kind === "none"
                    ? undefined
                    : settings.allowedFileTypes.types.join(",")
                }
```

`undefined` rather than `""`, because an empty `accept` attribute means "no filter" and would show every file in the dialog. `addFiles` blocks them either way, but not offering them is clearer.

- [ ] **Step 6: Verify the build is green**

Run: `npx heft test`

Expected: PASS.

Then confirm nothing is left behind:

`git grep -n "allowedExtensions" -- src/`

Expected: no output. (`Form-Copy.txt` and `Form.reference.tsx` still contain the old name but are not compiled — if either appears, confirm the path is one of those two and leave it alone.)

- [ ] **Step 7: Commit**

```bash
git add src/webparts/bulkUpload/components/BulkUpload.tsx
git commit -m "feat: read allowed file types from the column in bulk upload"
```

---

## Task 7: Documentation

**Files:**
- Modify: `CLAUDE.md` — "Allowed File Types" section, "Critical Rules / Gotchas"
- Modify: `docs/superpowers/specs/2026-07-28-client-site-migration-runbook.md` — line 146 area

- [ ] **Step 1: Correct the allowed types in CLAUDE.md**

Replace the "Allowed File Types" section body with:

```markdown
Driven by the **`AllowedFileTypes`** multi-select Choice column on `DMS Config`
(row `allowedExtensions`) — the **single source of truth**. `SettingValue` is no
longer read for this setting; clear it once a site is verified.

Current choices: `.pdf .doc .docx .xls .xlsx` (client policy as of 2026-07-30).
Everything else is rejected before upload.

- **Nothing ticked = hard block**, with a message naming the column and list. The
  client fixes it themselves in one click.
- **Column absent / config unreadable** = code fallback `FALLBACK_FILE_TYPES` in
  `src/shared/allowedFileTypes.ts`, plus an admin-only warning. Adding the column
  is a required step when provisioning any new site.
- Adding a type the client cannot already tick means editing the column's
  **Choices**, which needs Manage Lists (site Owner/admin).

See `docs/superpowers/specs/2026-07-30-allowed-file-types-dropdown-design.md`.
```

- [ ] **Step 2: Add a gotcha**

Append to "Critical Rules / Gotchas":

```markdown
10. **`accept` tokens need a leading dot.** `<input type="file" accept="png">` is
    not a filter for `.png` — a token without a leading `.` is parsed as a MIME
    type, is invalid without a `/`, and is silently dropped, greying the type out
    of the file dialog. Neither upload web part has a drag-and-drop handler, so the
    picker is the only way in and this is a hard block. Worse, the JS validator
    uses `endsWith(ext)`, which *accepts* a dotless `png` — so one typo produced two
    different behaviours. Normalized in `src/shared/allowedFileTypes.ts`; cost half
    a day on 2026-07-30.
```

- [ ] **Step 3: Update the migration runbook**

In the seed-rows code block, replace line 146:

```
allowedExtensions         .pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.csv,.jpg,.jpeg,.png
```

with:

```
allowedExtensions         (leave SettingValue EMPTY — see AllowedFileTypes below)
```

Then insert immediately after the code block that closes on line 147:

```markdown
> ⚠ **REQUIRED: the `AllowedFileTypes` column.** File types are read from a
> multi-select Choice column on `DMS Config`, **not** from `SettingValue`. Create it
> before the first upload test:
>
> | Property | Value |
> |---|---|
> | Internal name | `AllowedFileTypes` |
> | Type | Choice, **allow multiple selections** |
> | Fill-in choices | **Disabled** (enabling it restores the typo risk it exists to remove) |
> | Choices | `.pdf` `.doc` `.docx` `.xls` `.xlsx` |
>
> Then tick all five on the `allowedExtensions` row. There is **no fallback to
> `SettingValue`** — a site without this column runs on the hardcoded
> `FALLBACK_FILE_TYPES` in `src/shared/allowedFileTypes.ts` and shows admins a
> warning. Verify with:
>
> ```
> /_api/web/lists/getbytitle('DMS%20Config')/fields?$select=InternalName,FillInChoice,Choices&$filter=InternalName%20eq%20'AllowedFileTypes'
> ```
>
> Expect `"FillInChoice": false` and exactly those five choices, each with a leading dot.
```

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md docs/superpowers/specs/2026-07-28-client-site-migration-runbook.md
git commit -m "docs: record the AllowedFileTypes column and the accept-token gotcha"
```

---

## Task 8: Manual verification on the live site

Not automatable — needs a real SharePoint tenant and a real file dialog.

**Every check needs a hard refresh (Ctrl+F5).** Settings are read once in a mount-time `useEffect`, so a stale page shows stale config. That was half the confusion on 2026-07-30.

- [ ] **Step 1: Deploy**

```bash
npm run build
```

Upload the `.sppkg` from `sharepoint/solution/` to the tenant App Catalog, then hard-refresh the page hosting the Upload Form.

- [ ] **Step 2: Confirm the happy path**

With all five ticked on the `allowedExtensions` row: open the picker, confirm `.pdf .doc .docx .xls .xlsx` are selectable and a `.png` is not offered. Upload a PDF end to end.

- [ ] **Step 3: Confirm the single source of truth**

Set `SettingValue` on that row to `.png` — contradicting the ticks. Hard-refresh. Confirm **nothing changes**: PNG is still refused, the five are still offered. This is the regression check for spec §3.

- [ ] **Step 4: Confirm the hard block**

Untick everything in `AllowedFileTypes`. Hard-refresh. Confirm the card is replaced by the message naming Allowed File Types and DMS Config, no file can be selected, and no upload is possible. Re-tick the five and confirm recovery after a refresh.

- [ ] **Step 5: Confirm the unreadable path looks different**

In DevTools, block the `DMS Config` request (Network → right-click → Block request URL) and reload. Confirm: the console names the HTTP status, an admin sees the *unreadable* message rather than the *empty* one, and uploads still work on the five fallback types.

- [ ] **Step 6: Repeat steps 2 and 4 on the Bulk Upload web part**

Confirm the batch picker offers the same types, and that with nothing ticked `addFiles` refuses every file with the same message.

- [ ] **Step 7: Clean up**

Restore `SettingValue` and then clear it (rollout step 5) — leaving a stale value in a cell nothing reads is the `termSet_vendor` problem. Then confirm one more hard-refresh still behaves correctly with the cell empty.

---

## Out of scope

Carried from the spec, deliberately not in this plan:

- Showing the accepted types next to the file picker (spec item (b), never confirmed).
- Offering `.ppt .pptx .txt .csv .jpg .jpeg` as choices.
- Removing the retired `termSet_vendor` row, and the `Levels` JSON fill-dragged into the
  `col_businessSegment` / `col_businessSegmentTid` rows. Both inert, both misleading.

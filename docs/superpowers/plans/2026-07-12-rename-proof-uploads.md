# Rename-Proof Uploads Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Staging uploads resolve their destination folder by permanent `UniqueId` instead of a name-built path, so renaming a folder never breaks uploads.

**Architecture:** A new `DMS Folder Map` SharePoint list stores `TermGuid → FolderUniqueId` rows. A shared TypeScript module reads/writes those rows and resolves folders. `Form.tsx` looks up the selected term's `FolderUniqueId` and uploads via `GetFolderById(guid'…')`. A new read-only `reconciliation` admin web part populates the list by auto-matching terms to existing folders (or letting an admin paste a path), capturing each folder's `UniqueId`.

**Tech Stack:** SPFx 1.23 / React / TypeScript, `context.spHttpClient` direct REST, Heft build (no unit-test harness — verification is `npm run build` + manual workbench testing).

**Spec:** [2026-07-12-rename-proof-upload-design.md](../specs/2026-07-12-rename-proof-upload-design.md)

---

## Conventions for this plan

- **No unit-test harness exists** (`heft test` only lint/builds; Jest absent). "Verify" steps therefore mean: run `npm run build` and confirm it exits green, and/or run the app in the workbench and observe the stated outcome. Where a step defines a pure function, the verification is a compile check plus a described manual exercise — do **not** scaffold Jest.
- **Build command:** `nvm use 22 && npm run build` (must be Node 22). Dev server: `npm run start`.
- **Workbench URL:** `https://dcidigitalcom.sharepoint.com/_layouts/workbench.aspx?debugManifestsFile=https://localhost:4321/temp/build/manifests.js&debug=true&noredir=true`
- **Commit after every task.** We are on branch `feat/rename-proof-uploads`.

---

## File Structure

**Create:**
- `src/shared/dmsFolderMap.ts` — the only shared module. List/field names, types, and all `DMS Folder Map` read/write + folder-resolve helpers. Shared by `Form` and `reconciliation` so the schema lives in exactly one place.
- `src/webparts/reconciliation/ReconciliationWebPart.ts` — web part entry (mirrors `OnboardingWebPart.ts`).
- `src/webparts/reconciliation/ReconciliationWebPart.manifest.json` — manifest (new GUID).
- `src/webparts/reconciliation/components/IReconciliationProps.ts` — props (context only).
- `src/webparts/reconciliation/components/Reconciliation.tsx` — the admin UI + term/folder logic.
- `src/webparts/reconciliation/loc/en-us.js` and `mystrings.d.ts` — loc stubs (mirror onboarding).

**Modify:**
- `src/webparts/form/components/Form.tsx` — replace name-path resolution with `UniqueId` resolution.
- `config/config.json` — register the `reconciliation` bundle + localized resource.

**Manual (SharePoint, no code file):**
- Create the `DMS Folder Map` list via a browser-console REST script (Task 1).

---

## Task 1: Create the `DMS Folder Map` SharePoint list

This is a one-time tenant setup step, run in the browser console on the site (same approach used for the flat-view PATCH). It is not compiled code.

**Files:** none (SharePoint infra).

- [ ] **Step 1: Run the list-creation script in the browser console**

Open the site `https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground`, press F12 → Console, paste and run:

```js
(async () => {
  const web = "https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground";
  const digest = await fetch(`${web}/_api/contextinfo`, {
    method: "POST", headers: { Accept: "application/json;odata=nometadata" }
  }).then(r => r.json()).then(d => d.FormDigestValue);

  const post = (url, body) => fetch(`${web}${url}`, {
    method: "POST",
    headers: {
      "Accept": "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest
    },
    body: JSON.stringify(body)
  }).then(async r => ({ status: r.status, text: await r.text() }));

  // 1. Create the list (BaseTemplate 100 = generic list)
  console.log("create list:", await post("/_api/web/lists",
    { Title: "DMS Folder Map", BaseTemplate: 100, Description: "Term GUID -> folder UniqueId map for rename-proof uploads" }));

  const fieldsUrl = "/_api/web/lists/getbytitle('DMS Folder Map')/fields";
  // 2. Add text fields (internal name == title because no spaces)
  for (const name of ["TermGuid", "FolderUniqueId", "FolderUrl"]) {
    console.log(`field ${name}:`, await post(fieldsUrl,
      { Title: name, FieldTypeKind: 2, MaxLength: 255 })); // 2 = Single line of text
  }
  // 3. Add the Section choice field
  console.log("field Section:", await post(fieldsUrl, {
    Title: "Section", FieldTypeKind: 6, // 6 = Choice
    Choices: { results: ["Departments", "Projects"] }
  }));
  console.log("DONE — verify in List Settings");
})();
```

- [ ] **Step 2: Index the `TermGuid` column**

In the browser console:

```js
(async () => {
  const web = "https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground";
  const digest = await fetch(`${web}/_api/contextinfo`, {
    method: "POST", headers: { Accept: "application/json;odata=nometadata" }
  }).then(r => r.json()).then(d => d.FormDigestValue);

  const res = await fetch(`${web}/_api/web/lists/getbytitle('DMS Folder Map')/fields/getbytitle('TermGuid')`, {
    method: "POST",
    headers: {
      "Accept": "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest,
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*"
    },
    body: JSON.stringify({ Indexed: true })
  });
  console.log("index TermGuid:", res.status); // 204 = success
})();
```

- [ ] **Step 3: Verify**

Go to `Site Contents` → confirm the `DMS Folder Map` list exists. Open List Settings → confirm columns `Title`, `TermGuid`, `FolderUniqueId`, `FolderUrl`, `Section` and that `TermGuid` shows as indexed. Expected: all five columns present; TermGuid indexed.

- [ ] **Step 4: Add ONE test row by hand (needed to test Form in Task 3)**

In the list, add an item. Set `Title` = a real leaf you can pick in the upload form (e.g. `Account > AI Engineer`), `Section` = `Departments`. For `TermGuid` and `FolderUniqueId`, get real values in the console:

```js
// Replace the path with a real, existing leaf folder in Staging you will upload into:
(async () => {
  const web = "https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground";
  const path = "/sites/SPFX-Sandbox-Testing-Ground/Staging/Departments/Account D/AI Engineer";
  const r = await fetch(`${web}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeURIComponent(path)}'&$select=UniqueId,ServerRelativeUrl`,
    { headers: { Accept: "application/json;odata=nometadata" } }).then(x => x.json());
  console.log("FolderUniqueId =", r.UniqueId, " FolderUrl =", r.ServerRelativeUrl);
})();
```

Put the printed `UniqueId` into `FolderUniqueId` and the path into `FolderUrl`. For `TermGuid`, use the term GUID of that leaf (find it in the Term Store, or temporarily `console.log(subTeamValue)` when selecting it in the form). Save the row.

There is no code commit for this task.

---

## Task 2: Shared `dmsFolderMap` module

**Files:**
- Create: `src/shared/dmsFolderMap.ts`

- [ ] **Step 1: Write the module**

```ts
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

/** The rename-proof lookup list. Field internal names have no spaces. */
export const FOLDER_MAP_LIST = "DMS Folder Map";

export interface FolderMapping {
  termGuid: string;
  folderUniqueId: string;
  title: string;
  folderUrl: string;
  section: string;
}

/** Read the mapping row for a single term. Returns null if the term is not mapped. */
export async function lookupFolderMapping(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  termGuid: string,
): Promise<FolderMapping | null> {
  const url =
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(FOLDER_MAP_LIST)}')/items` +
    `?$select=Title,TermGuid,FolderUniqueId,FolderUrl,Section&$filter=TermGuid eq '${termGuid}'&$top=1`;
  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) throw new Error(`Folder map lookup failed: HTTP ${res.status}`);
  const data = await res.json();
  const row = (data.value ?? [])[0];
  if (!row) return null;
  return {
    termGuid: row.TermGuid,
    folderUniqueId: row.FolderUniqueId,
    title: row.Title,
    folderUrl: row.FolderUrl,
    section: row.Section,
  };
}

/** Read every mapping row's TermGuid — used to skip already-mapped terms. */
export async function loadMappedTermGuids(
  spHttpClient: SPHttpClient,
  siteUrl: string,
): Promise<Set<string>> {
  const guids = new Set<string>();
  let url: string | null =
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(FOLDER_MAP_LIST)}')/items` +
    `?$select=TermGuid&$top=5000`;
  while (url) {
    const res: SPHttpClientResponse = await spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`Folder map read failed: HTTP ${res.status}`);
    const data = await res.json();
    (data.value ?? []).forEach((r: { TermGuid?: string }) => {
      if (r.TermGuid) guids.add(r.TermGuid);
    });
    url = data["@odata.nextLink"] ?? null;
  }
  return guids;
}

/** Resolve a server-relative folder path to its stable UniqueId. Returns null if the folder does not exist. */
export async function resolveFolderByPath(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  serverRelativePath: string,
): Promise<{ uniqueId: string; serverRelativeUrl: string } | null> {
  const url =
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeURIComponent(serverRelativePath)}'` +
    `&$select=UniqueId,ServerRelativeUrl`;
  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) return null;
  const d = await res.json();
  return { uniqueId: d.UniqueId, serverRelativeUrl: d.ServerRelativeUrl };
}

/** Create one mapping row. */
export async function writeFolderMapping(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  m: FolderMapping,
): Promise<void> {
  const res: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(FOLDER_MAP_LIST)}')/items`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
      },
      body: JSON.stringify({
        Title: m.title,
        TermGuid: m.termGuid,
        FolderUniqueId: m.folderUniqueId,
        FolderUrl: m.folderUrl,
        Section: m.section,
      }),
    },
  );
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Write mapping failed: HTTP ${res.status} ${body.slice(0, 200)}`);
  }
}
```

- [ ] **Step 2: Verify it compiles**

Run: `nvm use 22 && npm run build`
Expected: build succeeds (no TS errors). The module is unused so far — that is fine; it must still typecheck.

- [ ] **Step 3: Commit**

```bash
git add src/shared/dmsFolderMap.ts
git commit -m "feat: add shared DMS Folder Map read/write module"
```

---

## Task 3: Switch `Form.tsx` upload to resolve by UniqueId

**Files:**
- Modify: `src/webparts/form/components/Form.tsx` (import; replace lines ~613-745; adjust item-fetch after upload)

- [ ] **Step 1: Add the import**

At the top of `Form.tsx`, after the existing `IFormProps` import (line 4), add:

```ts
import { lookupFolderMapping } from "../../../shared/dmsFolderMap";
```

- [ ] **Step 2: Replace the path-build + folder-check + upload block**

Replace everything from the `const currentMode = modes.find(...)` line (currently [Form.tsx:613](../../../src/webparts/form/components/Form.tsx#L613)) through the end of the `uploadRes` block and its error handling (currently ending at [Form.tsx:771](../../../src/webparts/form/components/Form.tsx#L771)) — i.e. the label→path construction, the `existsRes` duplicate check, the `folderRes` existence probe/diagnostic, and the `uploadRes` POST — with this:

```ts
    // Resolve the destination folder by its stable UniqueId (rename-proof),
    // NOT by a name-built path. The selected leaf term is the lookup key.
    setBusy(true);
    setStatus("Locating destination folder…");

    const mapping = await lookupFolderMapping(
      context.spHttpClient,
      siteUrl,
      selectedSubTeam.id,
    ).catch((e) => {
      console.error("Folder map lookup error:", e);
      return null;
    });

    if (!mapping || !mapping.folderUniqueId) {
      showToast(
        `This folder hasn't been mapped yet. Ask an administrator to register it in the reconciliation tool. (term ${selectedSubTeam.label})`,
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }
    const folderId = mapping.folderUniqueId;

    let uploadedServerRelativeUrl = "";
    try {
      // Duplicate check — GetFolderById targets the folder by UniqueId, so a
      // rename of that folder does not affect this lookup.
      setStatus("Checking for duplicates…");
      const dupRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')?$select=Exists`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (dupRes.ok) {
        showToast(
          `A file named "${finalName}" already exists in this location. Rename your document or choose a different file.`,
          "error",
        );
        setStatus("");
        setBusy(false);
        return;
      }
      // Any non-200 (typically 404 "not found") means no duplicate — proceed.

      setStatus("Uploading…");
      const uploadRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=false)?$select=ServerRelativeUrl`,
        SPHttpClient.configurations.v1,
        { body: file },
      );
      if (!uploadRes.ok) {
        // Surface SharePoint's real status/body, not a generic message.
        let detail = `HTTP ${uploadRes.status}`;
        try {
          const bodyText = await uploadRes.text();
          try {
            const errJson = JSON.parse(bodyText);
            const spMsg = errJson?.error?.message?.value ?? errJson?.error?.message;
            detail += spMsg ? ` — ${spMsg}` : bodyText ? ` — ${bodyText.slice(0, 300)}` : "";
          } catch {
            if (bodyText) detail += ` — ${bodyText.slice(0, 300)}`;
          }
        } catch {
          /* body already consumed */
        }
        console.error("Upload failed:", folderId, detail);
        // A 404 here means the mapped folder no longer exists (deleted after mapping).
        const hint =
          uploadRes.status === 404
            ? " The mapped folder may have been deleted — ask an administrator to re-run the reconciliation tool."
            : "";
        showToast(`Upload failed (${detail}).${hint}`, "error");
        setStatus("");
        setBusy(false);
        return;
      }
      const uploadJson = await uploadRes.json();
      uploadedServerRelativeUrl = uploadJson.ServerRelativeUrl;
```

- [ ] **Step 3: Point the item-fetch at the uploaded file's real URL**

The metadata step needs the new item's `Id`. Replace the current item-fetch (currently [Form.tsx:773-774](../../../src/webparts/form/components/Form.tsx#L773-L774)) that uses the old `fileUrl` variable:

```ts
      const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeURIComponent(fileUrl)}'`,
        SPHttpClient.configurations.v1,
      );
```

with (uses the URL returned by the upload, which reflects the folder's current location):

```ts
      const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeURIComponent(uploadedServerRelativeUrl)}'`,
        SPHttpClient.configurations.v1,
      );
```

- [ ] **Step 4: Remove now-dead code**

Delete the now-unused `normalizePathSegment` helper ([Form.tsx:53-54](../../../src/webparts/form/components/Form.tsx#L53-L54)) and its comment block ([49-54](../../../src/webparts/form/components/Form.tsx#L49-L54)), and the now-unused local variables in `handleUpload` that fed the old path build — `currentMode`, `subfolder`, `serverRelative`, `folderDeptLabel`, `subTeamSegments`, `libraryPath`, `fileUrl` — none are referenced after this change. Keep `selectedChoice` and `selectedSubTeam` (still used for metadata). Build will fail on any leftover reference, which tells you if you missed one.

- [ ] **Step 5: Verify build**

Run: `nvm use 22 && npm run build`
Expected: green. If it errors with "`libraryPath` is not defined" or "`normalizePathSegment` declared but never read", finish removing the dead references from Step 4.

- [ ] **Step 6: Verify in the workbench (manual, end-to-end)**

Run `npm run start`, open the workbench, add the Form web part. Using the same leaf you mapped in Task 1 Step 4:
1. Fill the form, pick that department + sub-team, attach a small allowed file, submit.
2. Expected: "Document uploaded successfully and is pending review." and the file appears in the mapped folder in Staging.
3. **Rename that folder in SharePoint** (e.g. `AI Engineer` → `AI Eng Team`). Upload again with a differently-named file.
4. Expected: upload **still succeeds** into the renamed folder — proving resolution is by `UniqueId`, not name.
5. Pick a leaf that has **no** row in `DMS Folder Map`. Expected: the "This folder hasn't been mapped yet…" toast, no upload.

- [ ] **Step 7: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: resolve upload folder by UniqueId (rename-proof)"
```

---

## Task 4: Scaffold the `reconciliation` web part

**Files:**
- Create: `src/webparts/reconciliation/ReconciliationWebPart.ts`
- Create: `src/webparts/reconciliation/ReconciliationWebPart.manifest.json`
- Create: `src/webparts/reconciliation/components/IReconciliationProps.ts`
- Create: `src/webparts/reconciliation/components/Reconciliation.tsx` (placeholder body this task)
- Create: `src/webparts/reconciliation/loc/en-us.js`
- Create: `src/webparts/reconciliation/loc/mystrings.d.ts`
- Modify: `config/config.json`

- [ ] **Step 1: Create `IReconciliationProps.ts`**

```ts
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReconciliationProps {
  context: WebPartContext;
}
```

- [ ] **Step 2: Create `ReconciliationWebPart.ts`**

```ts
import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import Reconciliation from "./components/Reconciliation";
import { IReconciliationProps } from "./components/IReconciliationProps";

export default class ReconciliationWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<IReconciliationProps> = React.createElement(
      Reconciliation,
      { context: this.context },
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }
}
```

- [ ] **Step 3: Create `ReconciliationWebPart.manifest.json`** (fresh, unique GUID — do not reuse onboarding's)

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
  "id": "a1c47d92-6b3e-4f08-9d21-7e5a2c9b64f1",
  "alias": "ReconciliationWebPart",
  "componentType": "WebPart",

  "version": "*",
  "manifestVersion": 2,

  "requiresCustomScript": false,
  "supportedHosts": ["SharePointWebPart", "SharePointFullPage"],
  "supportsThemeVariants": true,

  "preconfiguredEntries": [{
    "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
    "group": { "default": "Advanced" },
    "title": { "default": "Folder Reconciliation" },
    "description": { "default": "Admin tool — map term-store values to existing folder UniqueIds so uploads survive folder renames" },
    "officeFabricIconFontName": "Link",
    "properties": {}
  }]
}
```

- [ ] **Step 4: Create `loc/mystrings.d.ts`**

```ts
declare interface IReconciliationWebPartStrings {}

declare module "ReconciliationWebPartStrings" {
  const strings: IReconciliationWebPartStrings;
  export = strings;
}
```

- [ ] **Step 5: Create `loc/en-us.js`**

```js
define([], function () {
  return {};
});
```

- [ ] **Step 6: Create a placeholder `Reconciliation.tsx`** (real UI comes in Task 5 — this just renders so the bundle builds)

```tsx
import * as React from "react";
import { IReconciliationProps } from "./IReconciliationProps";

const Reconciliation: React.FC<IReconciliationProps> = () => {
  return <div>Folder Reconciliation — loading…</div>;
};

export default Reconciliation;
```

- [ ] **Step 7: Register the bundle in `config/config.json`**

In the `bundles` object add:

```json
    "reconciliation-web-part": {
      "components": [
        {
          "entrypoint": "./lib/webparts/reconciliation/ReconciliationWebPart.js",
          "manifest": "./src/webparts/reconciliation/ReconciliationWebPart.manifest.json"
        }
      ]
    }
```

In the `localizedResources` object add:

```json
    "ReconciliationWebPartStrings": "lib/webparts/reconciliation/loc/{locale}.js"
```

- [ ] **Step 8: Verify build**

Run: `nvm use 22 && npm run build`
Expected: green, and the build output lists a `reconciliation-web-part` bundle.

- [ ] **Step 9: Commit**

```bash
git add src/webparts/reconciliation config/config.json
git commit -m "feat: scaffold reconciliation web part"
```

---

## Task 5: Implement the reconciliation UI

**Files:**
- Modify: `src/webparts/reconciliation/components/Reconciliation.tsx` (replace placeholder with full implementation)

This component reads every mode's term tree, filters out already-mapped terms, computes each term's expected folder path, and lets the admin resolve+save each (auto-match) or edit the path and save (manual). It only ever writes to `DMS Folder Map` — it never touches folders/files.

- [ ] **Step 1: Replace the file with the full implementation**

```tsx
import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IReconciliationProps } from "./IReconciliationProps";
import {
  loadMappedTermGuids,
  resolveFolderByPath,
  writeFolderMapping,
} from "../../../shared/dmsFolderMap";

/** One selectable folder-bearing term, with everything needed to map it. */
interface Candidate {
  termGuid: string;
  title: string;          // "Account > AI Engineer"
  expectedPath: string;   // computed server-relative path
  section: string;        // "Departments" | "Projects"
  status: "" | "saving" | "saved" | "notfound" | "error";
  message: string;
}

interface Mode {
  key: string;
  termSetGuid: string;
  stagingFolder: string; // "Departments" | "Projects"
}

// Fallback modes — mirror Form.tsx DEFAULT_MODES. Reconciliation only needs
// the term set + staging folder per mode.
const DEFAULT_MODES: Mode[] = [
  { key: "department", termSetGuid: "eaba82e5-3e5f-4719-9a76-091f034ad407", stagingFolder: "Departments" },
  { key: "project", termSetGuid: "94ce322b-4515-4fda-8f50-35709f1f521d", stagingFolder: "Projects" },
];

const Reconciliation: React.FC<IReconciliationProps> = ({ context }) => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const serverRelative = context.pageContext.web.serverRelativeUrl;

  const [rows, setRows] = useState<Candidate[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [stagingLibrary, setStagingLibrary] = useState("Staging");

  type TermLite = { id: string; label: string };

  const loadChildren = async (termSetId: string, parentId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${parentId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({
      id: t.id,
      label: t.labels[0].name,
    }));
  };

  const loadTops = async (termSetId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error(`Term set ${termSetId} returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({
      id: t.id,
      label: t.labels[0].name,
    }));
  };

  // Try to read stagingLibrary from DMS Config; fall back to "Staging".
  const loadStagingLibrary = async (): Promise<string> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,SettingValue&$filter=Title eq 'stagingLibrary'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return "Staging";
      const data = await res.json();
      return (data.value ?? [])[0]?.SettingValue || "Staging";
    } catch {
      return "Staging";
    }
  };

  useEffect(() => {
    (async () => {
      try {
        const lib = await loadStagingLibrary();
        setStagingLibrary(lib);
        const mapped = await loadMappedTermGuids(context.spHttpClient, siteUrl);
        const candidates: Candidate[] = [];

        for (const mode of DEFAULT_MODES) {
          const tops = await loadTops(mode.termSetGuid).catch(() => []);
          for (const top of tops) {
            // Walk every descendant; each is a folder a file could land in.
            const walk = async (parentId: string, ancestors: string[]): Promise<void> => {
              const children = await loadChildren(mode.termSetGuid, parentId);
              for (const child of children) {
                const chain = [...ancestors, child.label];
                if (!mapped.has(child.id)) {
                  candidates.push({
                    termGuid: child.id,
                    title: `${top.label} > ${chain.join(" > ")}`,
                    expectedPath: `${serverRelative}/${lib}/${mode.stagingFolder}/${top.label}/${chain.join("/")}`,
                    section: mode.stagingFolder,
                    status: "",
                    message: "",
                  });
                }
                await walk(child.id, chain);
              }
            };
            await walk(top.id, []);
          }
        }
        setRows(candidates);
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        setLoading(false);
      }
    })().catch((e) => {
      setError(String(e));
      setLoading(false);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const updateRow = (i: number, patch: Partial<Candidate>): void => {
    setRows((prev) => prev.map((r, idx) => (idx === i ? { ...r, ...patch } : r)));
  };

  const saveRow = async (i: number): Promise<void> => {
    const row = rows[i];
    updateRow(i, { status: "saving", message: "" });
    const resolved = await resolveFolderByPath(context.spHttpClient, siteUrl, row.expectedPath);
    if (!resolved) {
      updateRow(i, { status: "notfound", message: "Folder not found — check the path and try again." });
      return;
    }
    try {
      await writeFolderMapping(context.spHttpClient, siteUrl, {
        termGuid: row.termGuid,
        folderUniqueId: resolved.uniqueId,
        title: row.title,
        folderUrl: resolved.serverRelativeUrl,
        section: row.section,
      });
      updateRow(i, { status: "saved", message: `Mapped → ${resolved.uniqueId}` });
    } catch (e) {
      updateRow(i, { status: "error", message: e instanceof Error ? e.message : String(e) });
    }
  };

  const autoMatchAll = async (): Promise<void> => {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].status !== "saved") {
        // eslint-disable-next-line no-await-in-loop
        await saveRow(i);
      }
    }
  };

  if (loading) return <div style={{ padding: 16 }}>Loading terms and existing mappings…</div>;
  if (error) return <div style={{ padding: 16, color: "#a00" }}>Error: {error}</div>;

  const remaining = rows.filter((r) => r.status !== "saved").length;

  return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>Folder Reconciliation</h2>
      <p>
        Library: <strong>{stagingLibrary}</strong> &nbsp;|&nbsp; Unmapped terms:{" "}
        <strong>{remaining}</strong> of {rows.length}
      </p>
      <button onClick={() => { autoMatchAll().catch(console.error); }} disabled={remaining === 0}>
        Auto-match all (resolve every expected path &amp; save)
      </button>
      <table style={{ width: "100%", marginTop: 12, borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ textAlign: "left", borderBottom: "2px solid #ccc" }}>
            <th style={{ padding: 6 }}>Term</th>
            <th style={{ padding: 6 }}>Folder path (edit if it doesn&apos;t match)</th>
            <th style={{ padding: 6 }}>Action</th>
            <th style={{ padding: 6 }}>Status</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={r.termGuid} style={{ borderBottom: "1px solid #eee", opacity: r.status === "saved" ? 0.5 : 1 }}>
              <td style={{ padding: 6 }}>{r.title}</td>
              <td style={{ padding: 6 }}>
                <input
                  style={{ width: "100%" }}
                  value={r.expectedPath}
                  disabled={r.status === "saved"}
                  onChange={(e) => updateRow(i, { expectedPath: e.target.value })}
                />
              </td>
              <td style={{ padding: 6 }}>
                <button onClick={() => { saveRow(i).catch(console.error); }} disabled={r.status === "saving" || r.status === "saved"}>
                  {r.status === "saving" ? "Saving…" : "Resolve & Save"}
                </button>
              </td>
              <td style={{ padding: 6, color: r.status === "saved" ? "#0a0" : r.status === "notfound" || r.status === "error" ? "#a00" : "#555" }}>
                {r.status === "saved" ? "✔ " : ""}{r.message}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default Reconciliation;
```

- [ ] **Step 2: Verify build**

Run: `nvm use 22 && npm run build`
Expected: green.

- [ ] **Step 3: Verify in the workbench (manual)**

Run `npm run start`, open the workbench, add the **Folder Reconciliation** web part.
1. Expected: it lists unmapped terms with a pre-filled expected path each, and a count.
2. Click **Resolve & Save** on a row whose folder exists. Expected: status turns green "✔ Mapped → <guid>", row dims, count decreases. Confirm a new row appeared in the `DMS Folder Map` list with correct `TermGuid`, `FolderUniqueId`, `FolderUrl`, `Section`, and that Created By/Created are populated automatically (this is the "who mapped it/when" requirement).
3. Edit a row's path to something wrong, Save. Expected: red "Folder not found — check the path and try again." No row written.
4. Click **Auto-match all**. Expected: every row whose expected path resolves gets saved; the rest stay with a red status for manual fixing.
5. Reload the web part. Expected: already-saved terms no longer appear (they're filtered by `loadMappedTermGuids`).

- [ ] **Step 4: Commit**

```bash
git add src/webparts/reconciliation/components/Reconciliation.tsx
git commit -m "feat: implement reconciliation UI (term -> folder UniqueId mapping)"
```

---

## Task 6: End-to-end verification & cleanup

**Files:**
- Modify: `src/webparts/form/components/Form.tsx` (remove temporary diagnostics if desired — optional)

- [ ] **Step 1: Full flow test**

With real data in the sandbox:
1. In **Folder Reconciliation**, Auto-match all, then manually resolve any leftovers until the count reads `0 of N`.
2. In **Form**, upload into several different departments/sub-teams. Expected: all succeed into the correct folders.
3. Rename two of those folders in SharePoint. Upload again into them. Expected: both still succeed (UniqueId resolution).
4. Create a brand-new folder natively, add its term as a new leaf in the Term Store (or reuse an unmapped one), reload Reconciliation, map it, then upload. Expected: succeeds.

- [ ] **Step 2: Confirm the "unmapped" guard**

Delete one row from `DMS Folder Map`, then upload into that term in the Form. Expected: "This folder hasn't been mapped yet…" toast, no upload.

- [ ] **Step 3: (Optional) remove leftover TEMP DIAGNOSTIC logging in Form.tsx**

The `[DMS DEBUG]` `console.log`/`console.table` lines around the metadata step ([Form.tsx:818-839](../../../src/webparts/form/components/Form.tsx#L818-L839)) are pre-existing debug aids. Remove them only if you've confirmed metadata tagging still works after this change.

- [ ] **Step 4: Final build**

Run: `nvm use 22 && npm run build`
Expected: green.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "chore: end-to-end verification for rename-proof uploads"
```

---

## Self-review notes

- **Spec coverage:** Config list (Task 1) · shared read/write module (Task 2) · Form resolve-by-UniqueId incl. duplicate check, upload, no-mapping guard, deleted-folder hint (Task 3) · reconciliation web part scaffold (Task 4) + UI with name auto-match and manual path pairing (Task 5) · "who mapped it" via built-in Author/Created columns surfaced in the list (Task 5 Step 3.2) · end-to-end rename test (Task 6). Spreadsheet import is explicitly out of scope for v1 per the spec — not planned.
- **Only leaf/descendant terms mapped, never department roots** — matches the spec's resolved risk (sub-team is required).
- **`FolderUrl` is written for display only**; resolution always uses `FolderUniqueId` via `GetFolderById`.
- **Naming consistency:** `lookupFolderMapping`, `loadMappedTermGuids`, `resolveFolderByPath`, `writeFolderMapping`, `FOLDER_MAP_LIST`, `FolderMapping` are used identically across Tasks 2, 3, 5.

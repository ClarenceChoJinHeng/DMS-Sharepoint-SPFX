# Share Guard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Members can request to share a DMS file/folder with an internal person; only a central admin approves, and on approval the recipient is granted item-scoped Read/Edit access.

**Architecture:** SPFx ListView Command Set button (mirrors the existing `uploadCommand`) navigates to a Request-Share web part page pre-filled with the selected item; the page writes a Pending row to the `DMS Share Requests` list; an admin-only Approval Queue web part grants the recipient access on Approve via `addroleassignment`. Enforcement is the native "owners-only sharing" lockdown (manual admin config). Pure logic lives in `src/shared/shareGuard.ts` with jest tests; web parts/extension are integration code following existing repo patterns.

**Tech Stack:** SPFx 1.23.0, React 17, TypeScript, Heft (NOT gulp), `context.spHttpClient` REST, `@pnp/spfx-controls-react` PeoplePicker, Power Automate (notify only).

**Design reference:** `docs/superpowers/specs/2026-07-20-share-guard-design.md`

**Approach refinement vs spec:** the spec said the command set "opens a dialog." This plan instead uses the **page-navigation** pattern already established by `src/extensions/uploadCommand/UploadCommandCommandSet.ts` (button → navigate to a page with query params). Lower risk, consistent with the codebase, and lets the request form reuse the PnP PeoplePicker in a normal web part.

**In-repo templates to copy from (not other tasks — real existing files):**
- Command set: `src/extensions/uploadCommand/UploadCommandCommandSet.ts` (+ its `.manifest.json`)
- Web part boilerplate: `src/webparts/approvalDocument/` (WebPart.ts, manifest, components/, IProps)
- Pure module + jest: `src/shared/siteMap.ts` + `src/shared/siteMap.test.ts`
- REST helpers (ensureuser, breakroleinheritance, addroleassignment, roledefinitions, GetFolderById): `src/webparts/folderManager/components/FolderManager.tsx`
- Registration examples: `config/config.json`, `config/package-solution.json`

---

## Task 1: Provision the list + lockdown (manual admin runbook)

**Files:**
- Create: `docs/superpowers/specs/2026-07-20-share-guard-runbook.md` (steps the admin follows)

No app code. This task produces a runbook and is a hard prerequisite for testing later tasks.

- [ ] **Step 1: Create the `DMS Share Requests` list** (Site contents → New → List). Add columns exactly:
  - `Title` (Single line) — default.
  - `ItemUrl` (Single line of text)
  - `ItemUniqueId` (Single line of text)
  - `ItemType` (Choice: File, Folder)
  - `Library` (Choice: Staging, Documents)
  - `Requester` (Person)
  - `Recipient` (Person)
  - `AccessLevel` (Choice: Read, Edit)
  - `Reason` (Multiple lines, plain text)
  - `Status` (Choice: Pending, Approved, Rejected; default Pending)
  - `DecidedBy` (Person)
  - `DecidedOn` (Date and Time)

- [ ] **Step 2: Read back the internal field names** via `https://.../_api/web/lists/getbytitle('DMS%20Share%20Requests')/fields?$select=Title,InternalName&$filter=Hidden eq false`. Record them — SharePoint may freeze names unpredictably (e.g. a column typed "Access Level" becomes `Access_x0020_Level`). **Later tasks must use the read-back internal names, never guesses.** Write them into this runbook.

- [ ] **Step 3: Item-level security** — List Settings → Advanced settings → Item-level Permissions: Read = "Read items that were created by the user"; Create and Edit = "Create items and edit items that were created by the user". This makes members see only their own requests.

- [ ] **Step 4: Grant members Contribute on the list** — List Settings → Permissions → stop inheriting → grant the **global site-access group** Contribute on this list only (so members can submit; item-level security limits visibility).

- [ ] **Step 5: NATIVE SHARING LOCKDOWN (the enforcement)** — Site Permissions → *Sharing settings* (gear or `/_layouts/15/mngsiteadmin.aspx` / Site permissions panel) → set **"Only site owners can share files, folders, and the site."** Verify: sign in as a member, confirm the Share/Copy-link either is hidden or errors. **If this step is skipped the guard is bypassable.**

- [ ] **Step 6: Record the two site role definition IDs** used later — GET `https://.../_api/web/roledefinitions?$select=Id,Name` and note the Id for **"Read"** and **"Edit"**. (The web part looks these up by name at runtime; recording them is just for verification.)

---

## Task 2: Pure logic — `src/shared/shareGuard.ts` (TDD)

**Files:**
- Create: `src/shared/shareGuard.ts`
- Test: `src/shared/shareGuard.test.ts`

- [ ] **Step 1: Write the failing test** `src/shared/shareGuard.test.ts`:

```ts
import {
  AccessLevel,
  roleNameForAccess,
  parseShareTarget,
  buildRequestPayload,
  itemAbsoluteUrl,
  ShareTarget,
} from "./shareGuard";

describe("roleNameForAccess", () => {
  it("maps Read and Edit to SharePoint role names", () => {
    expect(roleNameForAccess("Read")).toBe("Read");
    expect(roleNameForAccess("Edit")).toBe("Edit");
  });
});

describe("parseShareTarget", () => {
  it("reads item context from URLSearchParams", () => {
    const q = new URLSearchParams(
      "itemUrl=%2Fsites%2FX%2FStaging%2FGHO%2Ff.pdf&itemId=abc-123&itemType=File&library=Staging",
    );
    const t = parseShareTarget(q);
    expect(t).toEqual<ShareTarget>({
      itemUrl: "/sites/X/Staging/GHO/f.pdf",
      itemUniqueId: "abc-123",
      itemType: "File",
      library: "Staging",
    });
  });

  it("returns null when required params are missing", () => {
    expect(parseShareTarget(new URLSearchParams("itemUrl=x"))).toBeNull();
  });

  it("defaults itemType to File and library to Documents when unrecognised", () => {
    const t = parseShareTarget(
      new URLSearchParams("itemUrl=%2Fa&itemId=z&itemType=Bogus&library=Bogus"),
    );
    expect(t?.itemType).toBe("File");
    expect(t?.library).toBe("Documents");
  });
});

describe("buildRequestPayload", () => {
  it("builds the list item body with Pending status", () => {
    const target: ShareTarget = {
      itemUrl: "/sites/X/Staging/GHO/f.pdf",
      itemUniqueId: "abc-123",
      itemType: "File",
      library: "Staging",
    };
    const body = buildRequestPayload(target, 42, "Audit review", "Read");
    expect(body).toEqual({
      Title: "f.pdf",
      ItemUrl: "/sites/X/Staging/GHO/f.pdf",
      ItemUniqueId: "abc-123",
      ItemType: "File",
      Library: "Staging",
      RecipientId: 42,
      AccessLevel: "Read",
      Reason: "Audit review",
      Status: "Pending",
    });
  });

  it("derives Title from a folder path without a trailing slash", () => {
    const target: ShareTarget = {
      itemUrl: "/sites/X/Documents/GHO/GCO",
      itemUniqueId: "u",
      itemType: "Folder",
      library: "Documents",
    };
    expect(buildRequestPayload(target, 1, "", "Edit").Title).toBe("GCO");
  });
});

describe("itemAbsoluteUrl", () => {
  it("joins origin and encoded server-relative url", () => {
    expect(itemAbsoluteUrl("https://t.sharepoint.com", "/sites/X/Docs/a b.pdf"))
      .toBe("https://t.sharepoint.com/sites/X/Docs/a%20b.pdf");
  });
});
```

- [ ] **Step 2: Run it, verify it fails** — `npx heft test` → FAIL (module not found).

- [ ] **Step 3: Implement `src/shared/shareGuard.ts`:**

```ts
// Pure logic for Share Guard: parse the item context passed from the command-set
// button, map access level to a SharePoint role name, and build the request list
// item body. No SPHttpClient here — kept pure for jest (mirrors siteMap.ts).

export type AccessLevel = "Read" | "Edit";
export type ItemType = "File" | "Folder";
export type LibraryName = "Staging" | "Documents";

export interface ShareTarget {
  itemUrl: string;        // server-relative path of the file/folder
  itemUniqueId: string;   // rename-proof handle used for the grant
  itemType: ItemType;
  library: LibraryName;
}

/** Access level -> SharePoint role definition name (looked up by name at runtime). */
export function roleNameForAccess(level: AccessLevel): string {
  return level === "Edit" ? "Edit" : "Read";
}

/** Parse the item context the command-set button puts on the query string. */
export function parseShareTarget(q: URLSearchParams): ShareTarget | null {
  const itemUrl = q.get("itemUrl") ?? "";
  const itemUniqueId = q.get("itemId") ?? "";
  if (!itemUrl || !itemUniqueId) return null;
  const itemType: ItemType = q.get("itemType") === "Folder" ? "Folder" : "File";
  const library: LibraryName = q.get("library") === "Staging" ? "Staging" : "Documents";
  return { itemUrl, itemUniqueId, itemType, library };
}

/** Last path segment (no trailing slash) — the display name / Title. */
export function itemLeafName(serverRelativeUrl: string): string {
  const trimmed = serverRelativeUrl.replace(/\/+$/, "");
  const idx = trimmed.lastIndexOf("/");
  return idx >= 0 ? trimmed.slice(idx + 1) : trimmed;
}

/** Build the DMS Share Requests list item body. RecipientId = ensured SP user id.
 *  NOTE: field internal names below are the DEFAULT-cased names; if Task 1 step 2
 *  reveals different frozen internal names, update this object AND the tests. */
export function buildRequestPayload(
  target: ShareTarget,
  recipientId: number,
  reason: string,
  accessLevel: AccessLevel,
): Record<string, unknown> {
  return {
    Title: itemLeafName(target.itemUrl),
    ItemUrl: target.itemUrl,
    ItemUniqueId: target.itemUniqueId,
    ItemType: target.itemType,
    Library: target.library,
    RecipientId: recipientId,
    AccessLevel: accessLevel,
    Reason: reason,
    Status: "Pending",
  };
}

/** Absolute click-through URL for an item, origin + encoded server-relative path. */
export function itemAbsoluteUrl(origin: string, serverRelativeUrl: string): string {
  return `${origin}${encodeURI(serverRelativeUrl)}`;
}
```

- [ ] **Step 4: Run tests, verify pass** — `npx heft test` → all green (existing 32 + new).

- [ ] **Step 5: Commit** — `git add src/shared/shareGuard.ts src/shared/shareGuard.test.ts && git commit` (message: "feat: shareGuard pure helpers + tests").

> After Task 1 step 2 gives real internal names, if any differ from the defaults (`RecipientId`, `AccessLevel`, `ItemUniqueId`, etc.), update `buildRequestPayload` and the test in lock-step. `RecipientId` is the Person field's internal name + `Id` suffix (SharePoint sets Person fields by `<InternalName>Id`).

---

## Task 3: Approval Queue web part (admin: Approve → grant)

**Files:**
- Create: `src/webparts/shareApproval/ShareApprovalWebPart.ts`
- Create: `src/webparts/shareApproval/ShareApprovalWebPart.manifest.json`
- Create: `src/webparts/shareApproval/components/ShareApproval.tsx`
- Create: `src/webparts/shareApproval/components/IShareApprovalProps.ts`
- Create: `src/webparts/shareApproval/loc/mystrings.d.ts`, `src/webparts/shareApproval/loc/en-us.js`
- Modify: `config/config.json` (bundle + localizedResource)
- Modify: `config/package-solution.json` (componentId in feature + version bump)

- [ ] **Step 1: Scaffold WebPart + manifest + IProps + loc** by copying `src/webparts/approvalDocument/` files and renaming `ApprovalDocument`→`ShareApproval`. Generate a NEW GUID for the manifest `id` (e.g. run `[guid]::NewGuid()` in PowerShell). `IShareApprovalProps` = `{ context: WebPartContext }` (like IApprovalDocumentProps).

- [ ] **Step 2: Implement `ShareApproval.tsx`.** Reuse the exact REST idioms from `FolderManager.tsx`. Core logic:

```tsx
// Admin-only queue for DMS Share Requests. Lists Pending rows; Approve grants the
// recipient the chosen role on the item (resolved by UniqueId) and sets Status.
import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IShareApprovalProps } from "./IShareApprovalProps";
import { itemAbsoluteUrl } from "../../../shared/shareGuard";

const LIST = "DMS Share Requests";

interface Row {
  id: number; title: string; itemUrl: string; itemUniqueId: string;
  itemType: "File" | "Folder"; recipientId: number; recipientTitle: string;
  accessLevel: "Read" | "Edit"; reason: string; status: string;
}

const ShareApproval: React.FC<IShareApprovalProps> = ({ context }) => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const origin = new URL(siteUrl).origin;
  const [rows, setRows] = useState<Row[]>([]);
  const [busy, setBusy] = useState<number | null>(null);
  const [msg, setMsg] = useState("");

  const load = useCallback(async (): Promise<void> => {
    const url = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items`
      + `?$select=Id,Title,ItemUrl,ItemUniqueId,ItemType,AccessLevel,Reason,Status,RecipientId,Recipient/Title`
      + `&$expand=Recipient&$filter=Status eq 'Pending'&$orderby=Id desc&$top=200`;
    const res: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) { setMsg(`Could not load requests (HTTP ${res.status}).`); return; }
    const data = await res.json();
    setRows((data.value ?? []).map((r: Record<string, unknown>) => ({
      id: r.Id, title: r.Title, itemUrl: r.ItemUrl, itemUniqueId: r.ItemUniqueId,
      itemType: r.ItemType, recipientId: r.RecipientId, recipientTitle: (r.Recipient as {Title?: string})?.Title ?? "",
      accessLevel: r.AccessLevel, reason: r.Reason, status: r.Status,
    })) as Row[]);
  }, [context, siteUrl]);

  useEffect(() => { load().catch(() => undefined); }, [load]);

  const roleDefId = async (name: string): Promise<number | undefined> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/roledefinitions?$select=Id,Name&$filter=Name eq '${name}'`,
      SPHttpClient.configurations.v1, { headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) return undefined;
    return ((await res.json()).value ?? [])[0]?.Id as number | undefined;
  };

  // item path segment for the REST verb: files use GetFileById, folders GetFolderById
  const itemListItemPath = (r: Row): string =>
    r.itemType === "Folder"
      ? `GetFolderById(guid'${r.itemUniqueId}')/ListItemAllFields`
      : `GetFileById(guid'${r.itemUniqueId}')/ListItemAllFields`;

  const approve = async (r: Row): Promise<void> => {
    setBusy(r.id); setMsg("");
    try {
      const rdid = await roleDefId(r.accessLevel === "Edit" ? "Edit" : "Read");
      if (rdid === undefined) throw new Error(`No "${r.accessLevel}" role definition on site.`);
      const base = `${siteUrl}/_api/web/${itemListItemPath(r)}`;
      // Preserve existing access, add the recipient: break with copyRoleAssignments=true.
      await context.spHttpClient.post(
        `${base}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`,
        SPHttpClient.configurations.v1, { headers: { Accept: "application/json;odata=nometadata" } });
      await context.spHttpClient.post(
        `${base}/roleassignments/addroleassignment(principalid=${r.recipientId},roledefid=${rdid})`,
        SPHttpClient.configurations.v1, { headers: { Accept: "application/json;odata=nometadata" } });
      await setStatus(r.id, "Approved");
      setMsg(`Approved — ${r.recipientTitle} now has ${r.accessLevel} on ${r.title}.`);
      await load();
    } catch (e) { setMsg(`Approve failed: ${(e as Error).message}`); }
    finally { setBusy(null); }
  };

  const reject = async (r: Row): Promise<void> => {
    setBusy(r.id); try { await setStatus(r.id, "Rejected"); await load(); }
    finally { setBusy(null); }
  };

  const setStatus = async (id: number, status: string): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items(${id})`,
      SPHttpClient.configurations.v1, {
        headers: { Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "IF-MATCH": "*", "X-HTTP-Method": "MERGE" },
        body: JSON.stringify({ Status: status, DecidedOn: new Date().toISOString() }),
      });
    if (!res.ok) throw new Error(`status update HTTP ${res.status}`);
  };

  return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>Share requests — pending approval</h2>
      {msg && <p style={{ color: "#0f6c3f" }}>{msg}</p>}
      {rows.length === 0 && <p style={{ color: "#999" }}>No pending requests.</p>}
      {rows.map(r => (
        <div key={r.id} style={{ border: "1px solid #e0e0e0", borderRadius: 6, padding: 12, marginBottom: 8 }}>
          <strong>{r.title}</strong> → {r.recipientTitle} ({r.accessLevel})
          <div style={{ fontSize: 12, color: "#666" }}>{r.reason}</div>
          <a href={itemAbsoluteUrl(origin, r.itemUrl)} target="_blank" rel="noreferrer" style={{ fontSize: 12 }}>open item</a>
          <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
            <button disabled={busy === r.id} onClick={() => { approve(r).catch(() => undefined); }}>Approve</button>
            <button disabled={busy === r.id} onClick={() => { reject(r).catch(() => undefined); }}>Reject</button>
          </div>
        </div>
      ))}
    </div>
  );
};
export default ShareApproval;
```

- [ ] **Step 3: Register the bundle** in `config.json` under `bundles` (mirror `cross-site-browser-web-part`):

```json
"share-approval-web-part": {
  "components": [
    { "entrypoint": "./lib/webparts/shareApproval/ShareApprovalWebPart.js",
      "manifest": "./src/webparts/shareApproval/ShareApprovalWebPart.manifest.json" }
  ]
},
```
and add to `localizedResources`: `"ShareApprovalWebPartStrings": "lib/webparts/shareApproval/loc/{locale}.js"`.

- [ ] **Step 4: Register the componentId** — add the manifest's new GUID to the `componentIds` array in `config/package-solution.json` and bump `version` (e.g. `1.0.16.0`).

- [ ] **Step 5: Build** — `npm run build`. Expected: build clean, `sd-gatrie.sppkg` created, 32+ tests pass.

- [ ] **Step 6: Commit** — "feat: Share Approval queue web part (grant on approve)".

> **Verification (manual, after deploy):** hand-create a Pending row in `DMS Share Requests` pointing at a real item's UniqueId, add the web part to a page, click Approve, confirm the recipient gains access and Status flips to Approved.

---

## Task 4: Request-Share web part (member submits)

**Files:**
- Create: `src/webparts/requestShare/RequestShareWebPart.ts` + manifest + `components/RequestShare.tsx` + `components/IRequestShareProps.ts` + loc files
- Modify: `config/config.json`, `config/package-solution.json`

- [ ] **Step 1: Scaffold** by copying the `approvalDocument` web part shape; new GUID for the manifest.

- [ ] **Step 2: Implement `RequestShare.tsx`:** read the query string via `parseShareTarget(new URLSearchParams(window.location.search))`; render item name (from target), a **PnP `PeoplePicker`** (`import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";` — `principalTypes={[PrincipalType.User]}`, `personSelectionLimit={1}`), an AccessLevel `<select>` (Read/Edit), and a Reason `<textarea>`. On submit:
  1. `ensureuser` the picked login name → principal Id (same POST as `FolderManager.ensureGroupPrincipal`, but `logonName` = the PeoplePicker result's `loginName` for a user, e.g. `i:0#.f|membership|user@domain`).
  2. `buildRequestPayload(target, recipientId, reason, accessLevel)` → POST to `DMS Share Requests/items`.
  3. Show a success message; the row is Pending.

```tsx
// key submit logic (REST idioms mirror FolderManager.tsx)
const ensureUser = async (loginName: string): Promise<number> => {
  const res = await context.spHttpClient.post(`${siteUrl}/_api/web/ensureuser`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" },
      body: JSON.stringify({ logonName: loginName }) });
  if (!res.ok) throw new Error(`ensureuser HTTP ${res.status}`);
  return (await res.json()).Id as number;
};

const submit = async (): Promise<void> => {
  if (!target) { setErr("No item context — open this via the Request to Share button."); return; }
  if (!recipientLogin) { setErr("Pick a recipient."); return; }
  const rid = await ensureUser(recipientLogin);
  const body = buildRequestPayload(target, rid, reason, accessLevel);
  const res = await context.spHttpClient.post(
    `${siteUrl}/_api/web/lists/getbytitle('DMS%20Share%20Requests')/items`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata" },
      body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`submit HTTP ${res.status} ${(await res.text()).slice(0,200)}`);
  setDone(true);
};
```

- [ ] **Step 3: Register** bundle + localizedResource in `config.json`; componentId in `package-solution.json`.

- [ ] **Step 4: Create the page** — after deploy, add a Site Page `Request-Share.aspx` hosting only this web part. Record its server-relative URL for Task 5.

- [ ] **Step 5: Build** — `npm run build` (clean, tests pass).

- [ ] **Step 6: Commit** — "feat: Request-Share web part (member submits share request)".

---

## Task 5: Guarded Share button (ListView Command Set)

**Files:**
- Create: `src/extensions/shareRequest/ShareRequestCommandSet.ts`
- Create: `src/extensions/shareRequest/ShareRequestCommandSet.manifest.json`
- Modify: `config/config.json`, `config/package-solution.json`

- [ ] **Step 1: Manifest** — copy `uploadCommand`'s manifest, new GUID, alias `ShareRequestCommandSet`, one item:
```json
"items": { "REQUEST_SHARE": { "title": { "default": "Request to Share" }, "type": "command" } }
```

- [ ] **Step 2: Command set** `ShareRequestCommandSet.ts` (mirrors `UploadCommandCommandSet.ts`):

```ts
import {
  BaseListViewCommandSet,
  type IListViewCommandSetExecuteEventParameters,
  type IListViewCommandSetListViewUpdatedParameters,
} from "@microsoft/sp-listview-extensibility";

const REQUEST_SHARE_PAGE = "/SitePages/Request-Share.aspx";

export default class ShareRequestCommandSet extends BaseListViewCommandSet<Record<string, never>> {
  public onInit(): Promise<void> { return Promise.resolve(); }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const cmd = this.tryGetCommand("REQUEST_SHARE");
    if (!cmd) return;
    const lib = this.context.pageContext.list?.title;
    // Visible on Staging/Documents, exactly one row selected.
    cmd.visible = (lib === "Staging" || lib === "Documents") && event.selectedRows.length === 1;
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== "REQUEST_SHARE" || event.selectedRows.length !== 1) return;
    const row = event.selectedRows[0];
    const itemUrl = row.getValueByName("FileRef") as string;          // server-relative path
    const uniqueId = String(row.getValueByName("UniqueId") ?? "").replace(/[{}]/g, "");
    const fsObjType = row.getValueByName("FSObjType");                 // "1" = folder
    const itemType = String(fsObjType) === "1" ? "Folder" : "File";
    const library = this.context.pageContext.list?.title === "Staging" ? "Staging" : "Documents";
    const base = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, "");
    const q = `itemUrl=${encodeURIComponent(itemUrl)}&itemId=${encodeURIComponent(uniqueId)}`
      + `&itemType=${itemType}&library=${library}`;
    window.location.href = `${base}${REQUEST_SHARE_PAGE}?${q}`;
  }
}
```

- [ ] **Step 3: Register** the extension bundle in `config.json` (mirror `hideAppBar`/`uploadCommand` entries) and add the manifest GUID to `package-solution.json` `componentIds`; bump version.

- [ ] **Step 4: ClientSideComponentProperties / placement** — after deploy, register the command set on the Staging and Documents libraries (tenant-wide extension via `ClientSideComponentId`, or add through the list's extension registration). Document the exact IDs in the runbook.

- [ ] **Step 5: Build** — `npm run build`.

- [ ] **Step 6: Commit** — "feat: Request-to-Share command set button (Staging + Documents)".

---

## Task 6: Notification flow (runbook)

**Files:**
- Modify: `docs/superpowers/specs/2026-07-20-share-guard-runbook.md`

- [ ] **Step 1:** Power Automate → automated cloud flow, trigger "When an item is created or modified" on `DMS Share Requests`.
- [ ] **Step 2:** Condition `Status == Approved` (and modified) → **Send an email (V2)** as the **DMS No Reply** shared mailbox to `Recipient Email`, body includes `ItemUrl` (absolute). Optional second branch: on create + Status Pending, notify site admins.
- [ ] **Step 3:** Test by approving a request in the queue and confirming the email arrives with a working link.

---

## Task 7: End-to-end verification

- [ ] **Step 1:** As a **member** (non-admin, e.g. Kaisya): confirm the native Share button is blocked (lockdown) and the **Request to Share** button appears on a selected Staging/Documents item.
- [ ] **Step 2:** Submit a request to an internal recipient with Read → confirm a Pending row is created and the member sees only their own request.
- [ ] **Step 3:** As **admin**: open the Approval Queue, Approve → confirm the recipient gets access to exactly that item (open the emailed link as the recipient) and Status = Approved.
- [ ] **Step 4:** Reject a second request → confirm no access granted, Status = Rejected.
- [ ] **Step 5:** Confirm the recipient can open ONLY the shared item, not other unit content.

---

## Notes / decisions carried from the spec
- Enforcement is the manual **owners-only sharing** lockdown (Task 1 step 5) — the single most important step.
- Grant preserves existing access (`copyRoleAssignments=true`) and **adds** the recipient — item-scoped, never the unit.
- All field internal names must come from Task 1 step 2's read-back, not guesses (SharePoint freezes names).
- Fallback if the command set proves troublesome: the Request-Share web part page still works standalone (paste an item's UniqueId), so Tasks 3–4 deliver value even without Task 5.

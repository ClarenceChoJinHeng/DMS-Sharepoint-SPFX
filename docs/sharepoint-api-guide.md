# Calling the SharePoint REST API from SPFx

This guide explains how to send API requests to SharePoint from an SPFx web part using `SPHttpClient`, with examples taken from [Form.tsx](../src/webparts/form/components/Form.tsx).

---

## 1. The basics

SharePoint exposes a REST API at:

```
{siteUrl}/_api/...
```

From inside an SPFx web part you don't call it with `fetch`. You use `context.spHttpClient`, which automatically:

- Adds the bearer token / digest
- Sets the `OData-Version`, `Accept`, and `X-RequestDigest` headers
- Handles the SharePoint authentication context

You get `context` from your web part props (`IFormProps`).

### Imports

```ts
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
```

### Anatomy of a call

```ts
const res: SPHttpClientResponse = await context.spHttpClient.post(
  url,                                  // 1. The REST endpoint
  SPHttpClient.configurations.v1,       // 2. SPFx config (always v1)
  {                                     // 3. Request init
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ /* ... */ }),
  },
);

if (!res.ok) {
  // handle error
}
const data = await res.json();
```

Use `.get(...)` for reads, `.post(...)` for create/update/delete (SharePoint REST uses POST for most write operations).

---

## 2. Building the base URL

Always derive URLs from `context.pageContext`. Never hardcode them.

```ts
const siteUrl = context.pageContext.web.absoluteUrl;
// e.g. https://tenant.sharepoint.com/sites/MySite

const libraryServerRelativeUrl =
  `${context.pageContext.web.serverRelativeUrl}/Shared Documents`;
// e.g. /sites/MySite/Shared Documents
```

| Variable | Use it for |
|----------|------------|
| `absoluteUrl` | Building `_api` endpoints |
| `serverRelativeUrl` | Referencing folders/files inside REST calls |

Always `encodeURIComponent()` any path segment you pass inside an OData string like `'...'`.

---

## 3. Example: Create a folder

From [Form.tsx:66-72](../src/webparts/form/components/Form.tsx#L66-L72):

```ts
const folderRes: SPHttpClientResponse = await context.spHttpClient.post(
  `${siteUrl}/_api/web/folders/addUsingPath(decodedurl='${encodeURIComponent(folderServerRelativeUrl)}')`,
  SPHttpClient.configurations.v1,
  {
    headers: { "Content-Type": "application/json" },
  },
);
if (!folderRes.ok) {
  setStatus(`Could not create folder for File Set ${i + 1}.`);
  return;
}
```

Key points:

- Endpoint: `/_api/web/folders/addUsingPath(decodedurl='...')` is the modern folder-creation API (handles special characters better than the legacy `add('...')`).
- The path inside `decodedurl='...'` is a **server-relative URL**, URL-encoded.
- No body is needed — the path itself defines the folder.
- Always check `res.ok` before continuing.

---

## 4. Example: Upload a file

From [Form.tsx:78-87](../src/webparts/form/components/Form.tsx#L78-L87):

```ts
const uploadRes: SPHttpClientResponse = await context.spHttpClient.post(
  `${siteUrl}/_api/web/getfolderbyserverrelativeurl('${encodeURIComponent(folderServerRelativeUrl)}')/Files/Add(url='${encodeURIComponent(file.name)}',overwrite=true)`,
  SPHttpClient.configurations.v1,
  { body: file },
);
```

Key points:

- Endpoint pattern: `getfolderbyserverrelativeurl('<path>')/Files/Add(url='<name>',overwrite=true)`.
- The `body` is the raw `File` object — **do not** `JSON.stringify` it and do not set `Content-Type`. `SPHttpClient` streams the binary directly.
- `overwrite=true` replaces existing files with the same name.
- This API is fine for files under ~250 MB. For larger files use the chunked upload (`StartUpload`/`ContinueUpload`/`FinishUpload`).

---

## 5. Example: Read a folder's list item (GET)

From [Form.tsx:90-98](../src/webparts/form/components/Form.tsx#L90-L98):

```ts
const folderItemRes = await context.spHttpClient.get(
  `${siteUrl}/_api/web/getfolderbyserverrelativeurl('${encodeURIComponent(folderServerRelativeUrl)}')/ListItemAllFields?$select=Id`,
  SPHttpClient.configurations.v1,
);
if (!folderItemRes.ok) {
  setStatus(`Could not find folder list item for File Set ${i + 1}.`);
  return;
}
const folderItem = await folderItemRes.json();
// folderItem.Id is the list item ID for the folder
```

Key points:

- Folders in a document library are also list items, accessible via `ListItemAllFields`.
- Use `$select=Id` (and `$expand`, `$filter`) to keep responses small — these are standard OData query options.
- Use `.get(...)` instead of `.post(...)` for reads.

---

## 6. Example: Update metadata (managed-metadata field)

From [Form.tsx:104-122](../src/webparts/form/components/Form.tsx#L104-L122):

```ts
const departmentFieldValue = entry.department
  .map((term) => `${term.labels[0].name}|${term.id}`)
  .join(";");

const metaRes = await context.spHttpClient.post(
  `${siteUrl}/_api/web/lists/getbytitle('${libraryTitle}')/items(${folderItem.Id})/validateUpdateListItem`,
  SPHttpClient.configurations.v1,
  {
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      formValues: [
        { FieldName: "Department", FieldValue: departmentFieldValue },
      ],
    }),
  },
);
```

Key points:

- `validateUpdateListItem` is the right endpoint for **managed metadata** (taxonomy) columns. The simpler `items(id)` PATCH endpoint can't set taxonomy fields correctly.
- The `FieldValue` for a taxonomy column is `Label|TermGUID`, joined with `;` for multi-select.
- The response includes a `value` array. Each entry may have `HasException: true` with an `ErrorMessage` — you must check it because the HTTP response can still be 200 even when a field update failed:

```ts
const metaJson = await metaRes.json();
const fieldError = (metaJson.value ?? []).find(
  (v: { HasException?: boolean }) => v.HasException,
);
if (fieldError) {
  console.error("Field update error:", fieldError);
  return;
}
```

---

## 7. Patterns and gotchas

### Always check `res.ok`

`SPHttpClient` does **not** throw on HTTP errors. A 403 or 500 returns a response with `ok === false`. Handle it explicitly.

### `getbytitle` vs `getbyid`

- `lists/getbytitle('Documents')` — uses the library's **display title**, not its URL segment. For the default library this is `"Documents"`, even though the URL says `Shared Documents`.
- `lists/getbyid('<guid>')` — most stable; immune to renames.

### URL-encode path segments

Anything that goes inside `'...'` in an OData endpoint must be `encodeURIComponent`'d, especially folder paths with spaces or special characters.

### Status reporting

Use a `status` state hook to show progress and errors to the user, as in [Form.tsx:24](../src/webparts/form/components/Form.tsx#L24). Set it before each long call (`"Uploading…"`) and replace it with the failure reason on error.

### Sequential vs parallel

The form runs uploads sequentially (`for...of`) so failures can short-circuit cleanly. For independent uploads across multiple sets you could `Promise.all` them, but you lose ordered error reporting.

---

## 8. Quick reference of endpoints used

| Operation | Endpoint |
|-----------|----------|
| Create folder | `POST /_api/web/folders/addUsingPath(decodedurl='<path>')` |
| Upload file | `POST /_api/web/getfolderbyserverrelativeurl('<path>')/Files/Add(url='<name>',overwrite=true)` |
| Get folder's list item | `GET  /_api/web/getfolderbyserverrelativeurl('<path>')/ListItemAllFields?$select=Id` |
| Update list item (taxonomy-safe) | `POST /_api/web/lists/getbytitle('<title>')/items(<id>)/validateUpdateListItem` |
| Get list items | `GET  /_api/web/lists/getbytitle('<title>')/items?$select=...&$filter=...` |
| Delete item | `POST /_api/web/lists/getbytitle('<title>')/items(<id>)` with header `IF-MATCH: *` and `X-HTTP-Method: DELETE` |

---

## 9. Minimal template

```ts
async function callSharePoint(): Promise<void> {
  const siteUrl = context.pageContext.web.absoluteUrl;

  const res = await context.spHttpClient.post(
    `${siteUrl}/_api/web/...`,
    SPHttpClient.configurations.v1,
    {
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ /* payload */ }),
    },
  );

  if (!res.ok) {
    console.error(await res.text());
    return;
  }

  const data = await res.json();
  // use data
}
```

That's the full pattern — every call in [Form.tsx](../src/webparts/form/components/Form.tsx) is just a variation of this shape.

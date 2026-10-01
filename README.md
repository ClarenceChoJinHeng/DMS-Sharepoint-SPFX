# sd-gatrie (CRS / Guthrie Document Centre)

SPFx 1.23 web parts and extensions for SD Guthrie's document system on SharePoint Online.
Live site: `https://sdguthrie.sharepoint.com/sites/CRS`

## 1. Install

You need:

- **Node.js 22** (22.14 or later, below 23). Use nvm to switch versions.
- **Access** (ask SD Guthrie IT): Owner on the CRS site and its App Catalog, and the `gdc@sdguthrie.com` service account for Power Automate work.
- **A test site set up like CRS.** The web parts read real SharePoint lists and libraries (CRS Config, CRS Group Map, CRS Submissions, the approval and document libraries). On an empty site most screens will not load. How to build one: section 9 of the CRS Developer Guide (link under "Where things are").

Then, in this folder:

```
nvm use 22
npm ci
npx heft trust-dev-cert
```

`trust-dev-cert` is needed once per machine (it trusts the local HTTPS certificate).

## 2. Run locally

1. Open `config/serve.json` and change `initialPage` and `serveConfigurations.default.pageUrl` to **your test site's** workbench:
   `https://<tenant>.sharepoint.com/sites/<test-site>/_layouts/15/workbench.aspx`
   It points at the live CRS site by default. The web parts read and write real lists, so do not test on live.
2. Start the dev server:
   ```
   npm run start
   ```
3. The workbench opens. Add a web part from the toolbox.

To test on a real page instead, open the page with:
`?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/build/manifests.js`

Other commands:

| Command | What it does |
| --- | --- |
| `npx tsc --noEmit` | Type check |
| `npx heft test --clean` | Run the tests |

This project uses **Heft, not Gulp**. Ignore `gulp` in `package.json`.

## 3. Build

1. In `config/package-solution.json`, add 1 to the last number of `version` (e.g. `1.0.578.0` → `1.0.579.0`).
   Uploading the same version again makes it look like nothing changed.
2. Build:
   ```
   npm run build
   ```
   This runs the tests, then creates `sharepoint/solution/sd-guthrie.sppkg`.

**Always use `npm run build`.** Never run `heft package-solution` on its own: it packages the old JavaScript from the last build.

## 4. Deploy

1. Go to the site's App Catalog: `https://sdguthrie.sharepoint.com/sites/CRS/AppCatalog`
   (a site collection App Catalog, not the tenant one).
2. Upload `sd-guthrie.sppkg`, replacing the old file. Click **Deploy**.
3. Go to **Site contents**, find `sd-guthrie-client-side-solution` and click **Update** if it is offered.
4. Hard refresh (Ctrl+F5) or open a new tab. Check the version in Site contents matches the one you built.

Deploy to your test site first, then to CRS.

**First upload after the rename (1.0.580.0):** the package used to be called `sd-gatrie.sppkg`. In the App Catalog, rename the existing `sd-gatrie.sppkg` to `sd-guthrie.sppkg` first (**…** › **Rename**), then upload, so the upload replaces it. Never delete the old entry: the app's code is served from it.

## Where things are

| Folder | What is in it |
| --- | --- |
| `src/webparts/` | One folder per web part (e.g. `form` = Upload Form, `approvalDocument` = approval page, `mySubmissions`, `requests`, `documentSearch`, admin pages in `folderManager` and `userAccess`). |
| `src/extensions/` | Library page add-ons: the "Approve or reject" panel (`bulkApprove`) and the + New Folder button that also hides SharePoint's own Upload/Approve/Share (`hideAppBar`). `uploadCommand` is legacy: ignore it. |
| `src/shared/` | The app's logic (permissions, folder paths, upload checks, list names), each with a `.test.ts`. To change how something behaves, change it here. |
| `config/` | `package-solution.json` (version, web part list) and `serve.json` (local run page). |
| `scripts/` | Browser console scripts for checking live data (paste into F12 on the site). `check-*` / `diagnose-*` only read. |
| *(not in the repo)* | Power Automate flow exports. Kept outside the repo by the project owner; ask for the latest `.zip` exports. |

For the full design (architecture, where the data is stored, every list and column, security, components, flows, test site, troubleshooting), see the **[CRS Software Design Specification](docs/developer-guide/CRS-Software-Design-Specification.docx)** (Word; a [PDF copy](docs/developer-guide/CRS-Software-Design-Specification.pdf) sits next to it).

## Good to know

- **List and library names:** the client renames them. The code finds them from the candidate lists in `src/shared/naming.ts`. If a library stops loading after a rename, add its new title there, then build and deploy.
- **Power Automate:** about 25 flows handle routing, tagging, emails and deletes. They are not in this code, and their exports are not in this repo (the project owner keeps them). They all run as the service account `gdc@sdguthrie.com`. Never remove that account from the site's Owners group.
- **New web part:** add its ID to `componentIds` in `config/package-solution.json`, or it won't show in the toolbox (a test checks this).

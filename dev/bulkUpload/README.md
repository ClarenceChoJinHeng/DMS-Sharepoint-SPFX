# Bulk Upload — local mock harness

Runs the **real** `BulkUpload` component in a plain browser with a **mock
SharePoint context** — no site, no tenant, no login. For local UX testing when
you can't reach the workbench (SPFx 1.23 has no local workbench).

## Use it

1. Build the bundle (from the project root):

   ```
   npx esbuild dev/bulkUpload/harness.tsx --bundle \
     --outfile=dev/bulkUpload/harness.bundle.js \
     --alias:@microsoft/sp-http=./dev/bulkUpload/spHttpStub.ts \
     --jsx=automatic --define:process.env.NODE_ENV='"development"'
   ```

2. Open `dev/bulkUpload/index.html` in a browser (double-click it, or drag into
   a tab). Re-run step 1 and refresh after editing the component.

## What's mocked

- Signed in as a **site admin** → the privileged (full manual cascade) path.
- Two segments: **Group Head Office** and **Minamas Head Office**, each with
  Department → Unit trees. Doc Type / Year / Confidentiality / Vendor are canned.
- Folder routing always resolves; **uploads are simulated** (a short delay each,
  then reported "uploaded"). Nothing is written anywhere.

## What you can exercise

- Add batch → pick files → set folder + metadata → Save → folder tile.
- Second batch to a different folder; the 2-batch cap; Edit / Remove.
- Validation (missing fields, duplicate filename within a batch).
- Upload all → per-batch results.

## What it does NOT prove

Real SharePoint behaviour: term store, folder existence in Documents, ACLs,
`validateUpdateListItem` field validation, duplicate-skip against live files.
Those still need a real site.

## Files

- `harness.tsx` — entry; renders the component with the mock context.
- `mockContext.ts` — fake `spHttpClient` + page context (synthetic data).
- `spHttpStub.ts` — esbuild alias target for `@microsoft/sp-http`.
- `index.html` — page shell; loads `harness.bundle.js` (generated, gitignored).

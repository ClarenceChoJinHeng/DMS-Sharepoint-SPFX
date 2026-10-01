---
name: sdg-console-script
description: Use when live SharePoint data must be checked or repaired and the user will run a script in the browser DevTools console (diagnostics, counts, finding bad rows, bulk repairs).
---
# Browser console script

1. Copy `template.js` from this folder to `scripts/<verb>-<thing>.js` (for example `scripts/check-tag-status.js`).
2. Fill in the `CONFIG` block and the `MAIN` part. Keep it **read-only** unless the user asked for a repair.
   - A repair must have `DO_WRITE = false` by default and print what it *would* do first.
3. Rules for the script:
   - Find the site with `resolveWeb()` (never parse it from the page path by hand).
   - Every read goes through `getJson()` / `getAll()` (they cache-bust and check the answer is JSON).
   - Find libraries by title candidates or URL segment, never one hardcoded title.
   - `getAll()` pages with `nextLink`. Never trust a single `$top` read to prove something is absent.
   - A long loop prints progress every 25 items.
   - Big results go on `window.__result`. Print small summary tables only.
4. **Test it before handing it over:** run it in Node with a fake `fetch` returning synthetic data covering the good case, the bad case and a failed read.
5. Tell the user: open the site, press F12 → Console, paste, Enter, and send back the summary.

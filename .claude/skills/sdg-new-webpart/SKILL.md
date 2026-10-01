---
name: sdg-new-webpart
description: Use when adding a new web part, a new page, or a new extension to the SD Guthrie solution.
---
# Add a web part or page

1. Create `src/webparts/<name>/` like an existing small one (`auditLog` is a good model): `<Name>WebPart.ts`, `components/<Name>.tsx`, `I<Name>Props.ts`, manifest.
2. Register it:
   - Bundle entry in `config/config.json`
   - Its id in `config/package-solution.json` → `componentIds` (`solutionPackaging.test.ts` checks this). **If it's missing, the web part never appears in the toolbox.**
3. **Page access:** add or confirm a rule in `src/shared/pageAccessPolicy.ts` for the page file name.
   - Admin page names are matched by `/folder|group.?manag|config|setting|mapping|admin|access|audit/i` and locked to Owners.
   - Otherwise list the roles that may open it. Add a test.
4. **Admin page?** Add a row to CRS Settings (`src/shared/adminPages.ts`, resolved by page name, never a hardcoded URL) and wrap it with `withBackToSettings` in `render()`.
5. Styling: `/sdg-mobile-css`. Hooks above early returns.
6. Build and deploy with `/sdg-deploy`.
7. On the site: create the page, add the web part, **Publish** (a draft denies read-only users), then run Folder Reconciliation (it grants page access).
8. Add `Components/<Name>.md` with `/sdg-obsidian-update`.

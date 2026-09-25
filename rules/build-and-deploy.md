# Build and deploy

- Package with `npm run build` only. `heft package-solution` alone ships stale JavaScript.
- Bump `version` in `config/package-solution.json` on every deploy.
- Verify inside the `.sppkg`: `unzip -p sharepoint/solution/sd-gatrie.sppkg "ClientSideAssets/<bundle>_*.js" | grep -c "<new string>"`. Some components compile into two bundles.
- A web part must be in `package-solution.json` → `componentIds` (a test checks this).
- Checklist skill: `/sdg-deploy`.

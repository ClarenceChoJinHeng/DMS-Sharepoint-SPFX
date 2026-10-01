---
name: sdg-deploy
description: Use when packaging or deploying the SD Guthrie SPFx solution, or when the user says "build", "package", "deploy" or "ready to upload".
---
# SDG deploy checklist

1. Bump `version` in `config/package-solution.json` (last number +1).
2. Run `npm run build`. Never run `heft package-solution` alone, because it ships stale JavaScript.
3. Confirm the version in the package:
   `unzip -p sharepoint/solution/sd-guthrie.sppkg AppManifest.xml | grep -o 'Version="[^"]*"'`
4. Confirm the change is really inside the package. Grep a string the change added:
   `unzip -p sharepoint/solution/sd-guthrie.sppkg "ClientSideAssets/<bundle>_*.js" | grep -c "<string>"`
   (StructureManager/SegmentCreator/GroupManager compile into two bundles, so check both.)
5. Tell the user, in bullets:
   - Upload the `.sppkg` to the App Catalog.
   - Update the app in Site Contents.
   - Close old tabs, or hard refresh.
6. Say clearly what was verified (build, grep) and what is **not** site-tested yet.

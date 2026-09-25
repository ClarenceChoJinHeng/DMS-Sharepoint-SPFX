---
name: sdg-diagnose
description: Use when something is broken or behaves unexpectedly on the SharePoint site (blank web part, wrong data, button missing, error message, "it still doesn't work").
---
# Diagnose before fixing

Never name a cause until a check has shown it. Go in this order and stop at the first check that explains it.

1. **Which site?** `/sites/CRS` (SDG) or `/sites/ClarenceDMSTesting` (test).
2. **Which build is running?**
   - Did they hard refresh or close old tabs after deploying?
   - Is the version in Site Contents the one just built?
   - Is there a string on screen that only the new build has?
3. **Console:**
   - `Minified React error #310` = a hook below an early return.
   - `MIME type ('text/html')` = stale tab.
4. **Network tab:** the failing request's **status** and **response body**.
   - 400 = bad request or unknown field.
   - 403 = permissions.
   - 404 = wrong title/path, or security-trimmed.
   - 304 = cached answer.
   - 500 = server; read the body.
5. **Live data:** read the actual item/list over REST (write a read-only console script with `/sdg-console-script`).
6. **Flow involved?** Run `/sdg-flow-check` on the latest export.
7. Only now: state the cause, **with the evidence**, then propose a fix.

Report to the user in short bullets: what was checked, what it showed, the cause, the fix.

/*
 * diagnose-site-members-leak.js — READ ONLY. Browser console, run as a SITE ADMINISTRATOR.
 *
 * WHY THIS EXISTS
 * ----------------
 * `CRS_SITE_MEMBERS` is DELIBERATELY granted Read at the ROOT of the `Documents` library (never
 * `Staging`) — see `FolderManager.tsx`'s reconciliation, "Documents MUST keep it — the approval
 * guard resolves the destination folder as the approver and depends on that Read." That grant is
 * meant to stop the moment any Department/Unit folder below the root breaks its own unique
 * permissions, which reconciliation is supposed to do for EVERY Department and Unit folder.
 *
 * Reported live 2026-09-22: a genuinely bare member (nothing but `CRS_SITE_MEMBERS`, no unit or
 * persona group) opened Document Search on the Home page and got back two real documents from
 * `NBPOLHO / CDS / UPSUPPORT / 2024 / Agreement`, with full metadata (Created By, Confidentiality
 * Level, Legally Privileged, Business Segment, Department all populated) — which should be
 * impossible if `UPSUPPORT` (or an ancestor) correctly broke inheritance. This walks that exact
 * chain, folder by folder, and reports `HasUniqueRoleAssignments` plus whether a `_SITE_MEMBERS`
 * principal appears at each level, so the ONE folder that failed to break can be pointed at
 * directly rather than guessed at. It also checks whether `CRS_SITE_MEMBERS` itself holds any
 * grant BELOW the library root, which it never should.
 *
 * `NBPOLHO/CDS/UPSUPPORT` has been this project's own most heavily hand-tested unit this month
 * (self-approve tests, the managed-metadata fix, the tag/approve-proxy end-to-end test, several
 * stale-group findings) — a real candidate for having had its ACL hand-edited or left mid-repair at
 * some point, rather than this being a defect in the `CRS_SITE_MEMBERS` grant itself.
 *
 * ⚠ v2 — TWO REAL BUGS FIXED FROM THE FIRST LIVE RUN, KEPT HERE SO THE NEXT EDIT DOES NOT
 * REINTRODUCE EITHER:
 *   1. THE LIBRARY ROOT'S OWN ROLE ASSIGNMENTS 404'D. A library's root FOLDER does not expose role
 *      assignments the same way an ordinary subfolder does — `GetFolderByServerRelativeUrl(@f)/
 *      ListItemAllFields/roleassignments` only works on folders that are genuine list items below
 *      the root. The LIBRARY's own ACL has to be read through the LIST endpoint instead
 *      (`lists(guid'...')/roleassignments`), exactly as `diagnose-blank-libraries.js` already does
 *      — copying the folder-scoped helper for this one call was the mistake.
 *   2. THE FULL PER-LEVEL ROLE-ASSIGNMENT DUMP TRUNCATED IN DEVTOOLS before reaching the levels
 *      that actually mattered (CDS/UPSUPPORT/2024/Agreement never printed). NBPOLHO alone holds
 *      dozens of entries — every persona group under every department/unit needs a `Read, Limited
 *      Access` ancestor-browse corridor grant there, which is normal and expected, not a leak, but
 *      it is not worth console-printing in full at every level. The console output is now a
 *      SUMMARY per level (unique?, how many real principals, whether one matches `_SITE_MEMBERS`)
 *      — full detail is still computed and kept on `window.__leakReport` for follow-up
 *      (`console.table(window.__leakReport.chain)`, or inspect one level's `roleAssignments` by
 *      hand) rather than being force-printed and truncated.
 *
 * HOW TO RUN
 *   1. Open the site in a browser, signed in as a SITE ADMINISTRATOR (not the affected bare
 *      member — an ordinary account usually cannot read role assignments at all, which is exactly
 *      why this needs an admin session to diagnose).
 *   2. Adjust LIBRARY_URL_SEGMENT / PATH_SEGMENTS / FILE_NAMES below if you want to check a
 *      different document than the one reported.
 *   3. F12 -> Console -> paste this whole file -> Enter.
 *   4. Read the printed summary against the checklist at the bottom. Full detail is on
 *      `window.__leakReport` if you need to dig into one level's exact ACL.
 *
 * ⚠ EVERY REQUEST IS A GET. It creates nothing, changes nothing, deletes nothing.
 *
 * ⚠ `null`/`{ __error }` MEANS "COULD NOT READ", NEVER "NOT THERE". An absent grant reports an
 *   empty array; an unreadable one reports an error with the HTTP status. Do not collapse them —
 *   an admin account failing to read something here is itself worth knowing.
 */
(async () => {
  /* Same pattern as dump-site-inventory.js / diagnose-blank-libraries.js: never rely on
     `_spPageContextInfo` alone, it is undefined on several modern pages and in any non-top frame.
     Set SITE_OVERRIDE if this ever misreads. */
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const derived = (() => {
    const m = location.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    return location.origin + (m ? m[1] : "");
  })();
  const web = (SITE_OVERRIDE || derived).replace(/\/+$/, "");
  console.log("reading site:", web || "(root)");

  /* Report the exact path from the client's screenshot by default. Edit these two if the leak you
     are chasing is somewhere else — the LIBRARY segment must be the URL segment (never a title,
     since this client renames libraries routinely), and PATH_SEGMENTS is every folder from the
     library root down to (and including) the leaf folder holding the file. */
  const LIBRARY_URL_SEGMENT = "Shared Documents"; // "Documents" library's real URL segment
  const PATH_SEGMENTS = ["NBPOLHO", "CDS", "UPSUPPORT", "2024", "Agreement"];
  const FILE_NAMES = [
    "Error-Page_has_broken_image.xlsx",
    "Notice-Redirected_page_has_no_incoming_internal_links-links.xlsx",
  ];

  const SITE_MEMBERS_PATTERN = /site.?members/i;

  const get = async (url) => {
    try {
      const r = await fetch(web + url, { headers: { Accept: "application/json;odata=nometadata" } });
      if (!r.ok) return { __error: r.status };
      return await r.json();
    } catch (e) {
      return { __error: String(e && e.message ? e.message : e) };
    }
  };

  const rolesFor = (assignments) =>
    (assignments || []).map((a) => ({
      principal: a.Member ? a.Member.Title : `id ${a.PrincipalId}`,
      loginName: a.Member ? a.Member.LoginName : undefined,
      roles: (a.RoleDefinitionBindings || []).map((b) => b.Name),
    }));

  // ⚠ "Limited Access" ONLY is SharePoint's automatic entry for a principal granted DEEPER in the
  // tree, not a real grant at this level — filtered out so the one thing worth reading is not
  // drowned by it.
  const realRolesOnly = (rows) => rows.filter((r) => r.roles.some((x) => x !== "Limited Access"));

  // A FOLDER's role assignments — every ordinary path segment below the library root. Uses the
  // parameter-ALIAS OData form throughout — gotcha #9 in this project's own CLAUDE.md: an inline
  // literal 400s once the path is long/deep enough, and this path is six levels deep.
  const roleAssignmentsForFolder = async (serverRelUrl) => {
    const j = await get(
      `/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments` +
        `?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/LoginName,RoleDefinitionBindings/Name` +
        `&@f='${encodeURIComponent(serverRelUrl)}'`,
    );
    return j.__error !== undefined ? { __error: j.__error } : rolesFor(j.value);
  };

  // A LIBRARY's own role assignments — its ROOT folder does not expose them the same way an
  // ordinary subfolder does, so this goes through the LIST endpoint by GUID instead, exactly as
  // `diagnose-blank-libraries.js` already does. Do not "simplify" this back to the folder helper —
  // that was the v1 bug (404 on the library root).
  const roleAssignmentsForList = async (listId) => {
    const j = await get(
      `/_api/web/lists(guid'${listId}')/roleassignments` +
        `?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/LoginName,RoleDefinitionBindings/Name`,
    );
    return j.__error !== undefined ? { __error: j.__error } : rolesFor(j.value);
  };

  const summarize = (rolesOrError) => {
    if (rolesOrError.__error !== undefined) return { __error: rolesOrError.__error };
    const real = realRolesOnly(rolesOrError);
    const siteMembers = real.find((r) => SITE_MEMBERS_PATTERN.test(r.principal));
    return {
      realPrincipalCount: real.length,
      siteMembersEntry: siteMembers ? { principal: siteMembers.principal, roles: siteMembers.roles } : null,
    };
  };

  const out = { site: web || "(root)" };

  // ── Resolve the library by URL SEGMENT, never by title. ────────────────────────────────────
  const libs = await get(
    "/_api/web/lists?$select=Id,Title,HasUniqueRoleAssignments&$expand=RootFolder&$filter=BaseTemplate eq 101",
  );
  if (libs.__error !== undefined) {
    out.libraries = { __error: libs.__error };
    console.log(JSON.stringify(out, null, 2));
    return;
  }
  const lib = (libs.value || []).find((l) => {
    const u = (l.RootFolder && l.RootFolder.ServerRelativeUrl) || "";
    return u.toLowerCase().endsWith("/" + LIBRARY_URL_SEGMENT.toLowerCase());
  });
  if (!lib) {
    out.error = `Library with URL segment "${LIBRARY_URL_SEGMENT}" not found on this site.`;
    console.log(JSON.stringify(out, null, 2));
    return;
  }
  const libRootPath = lib.RootFolder.ServerRelativeUrl;
  out.library = {
    liveTitle: lib.Title,
    urlSegment: LIBRARY_URL_SEGMENT,
    serverRelativeUrl: libRootPath,
    hasUniqueRoleAssignments: lib.HasUniqueRoleAssignments,
  };

  // ── The library's OWN role assignments — confirms the deliberate `CRS_SITE_MEMBERS` root grant
  //    this script exists to trace the consequence of. Via the LIST endpoint (v2 fix #1). ──────
  const libraryRoles = await roleAssignmentsForList(lib.Id);
  out.libraryRoleAssignments = libraryRoles;
  out.librarySummary = summarize(libraryRoles);

  // ── Walk the chain, ONE FOLDER AT A TIME. Console gets a SUMMARY per level (v2 fix #2); the
  //    full ACL for every level is still stored on `out.chain[i].roleAssignments` for follow-up,
  //    just not force-printed. The FIRST summary showing `hasUniqueRoleAssignments: false` is the
  //    culprit: everything below it is inheriting from above, all the way back to the library
  //    root's CRS_SITE_MEMBERS grant. ─────────────────────────────────────────────────────────
  out.chain = [];
  let path = libRootPath;
  for (const seg of PATH_SEGMENTS) {
    path = `${path}/${seg}`;
    const meta = await get(
      `/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields` +
        `?$select=HasUniqueRoleAssignments&@f='${encodeURIComponent(path)}'`,
    );
    const entry = {
      segment: seg,
      path,
      hasUniqueRoleAssignments:
        meta.__error !== undefined ? { __error: meta.__error } : meta.HasUniqueRoleAssignments,
    };
    // Only pull the full ACL when it is worth reading: either it broke inheritance (confirm WHO is
    // on it) or it did NOT (confirm it really inherited nothing of its own — the smoking gun).
    if (meta.__error === undefined) {
      const roles = await roleAssignmentsForFolder(path);
      entry.roleAssignments = roles;
      entry.summary = summarize(roles);
    }
    out.chain.push(entry);
  }

  // ── The two reported files themselves — a file's OWN list item can (rarely) carry unique
  //    permissions independent of its parent folder, so checked directly rather than assumed. ───
  out.files = {};
  for (const name of FILE_NAMES) {
    const filePath = `${path}/${name}`;
    const meta = await get(
      `/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields` +
        `?$select=HasUniqueRoleAssignments&@f='${encodeURIComponent(filePath)}'`,
    );
    out.files[name] = {
      hasUniqueRoleAssignments:
        meta.__error !== undefined ? { __error: meta.__error } : meta.HasUniqueRoleAssignments,
    };
  }

  // Full detail is here for follow-up (`window.__leakReport.chain[2].roleAssignments`, etc.) —
  // never force-printed in full, which is what truncated in v1.
  window.__leakReport = out;

  console.log("=== library ===");
  console.log(JSON.stringify(out.library, null, 2));
  console.log("library role summary:", JSON.stringify(out.librarySummary, null, 2));

  console.log("\n=== chain (summary — full detail is on window.__leakReport.chain) ===");
  console.table(
    out.chain.map((c) => ({
      segment: c.segment,
      hasUniqueRoleAssignments:
        typeof c.hasUniqueRoleAssignments === "object" ? JSON.stringify(c.hasUniqueRoleAssignments) : c.hasUniqueRoleAssignments,
      realPrincipalCount: c.summary ? c.summary.realPrincipalCount : undefined,
      siteMembersEntry: c.summary ? JSON.stringify(c.summary.siteMembersEntry) : undefined,
    })),
  );

  console.log("\n=== files ===");
  console.log(JSON.stringify(out.files, null, 2));

  console.log(`
── READ IT AGAINST THIS ──
1. "library role summary" SHOULD show a "siteMembersEntry" holding "Read" — that confirms the
   deliberate root grant this whole report is about. If it is null, something else is going on and
   this whole theory is wrong; look elsewhere.
2. Read the "chain" table in order (NBPOLHO, then CDS, then UPSUPPORT, then 2024, then Agreement).
   Find the FIRST row whose "hasUniqueRoleAssignments" reads FALSE. That is the culprit: it never
   broke its own inheritance, so it and everything below it are still reading straight through from
   the library root — which is exactly how a bare "_SITE_MEMBERS" member sees real documents.
   - If it is "UPSUPPORT" itself (the Unit folder) — Folder Reconciliation, scoped to NBPOLHO,
     should re-break it and re-grant its own groups. Re-run it and diff this script's output
     before/after.
   - If it is "2024" or "Agreement" (below Unit) — those are DESIGNED to inherit from the Unit
     folder, per this project's own rule ("nothing below Unit breaks inheritance"). If THOSE read
     false while "UPSUPPORT" above them reads TRUE, that is actually correct and expected — the
     leak in that case is that UPSUPPORT's own ACL is wrong (check its "siteMembersEntry" for a
     non-null value that should not be there, or open
     "window.__leakReport.chain.find(c => c.segment === 'UPSUPPORT').roleAssignments" for the full
     list if a group looks missing).
3. On whichever row reads unique (true), check its "siteMembersEntry": non-null means a
   "_SITE_MEMBERS" grant sits there DIRECTLY. If that row is NBPOLHO or CDS (segment/department
   level), that is the deliberate ancestor-browse-adjacent root grant showing through and is
   expected. If it appears at UPSUPPORT (or below), that is a stray direct grant, not an
   inheritance failure, and should be REMOVED there specifically — re-running reconciliation will
   not touch a grant placed at a folder scope by hand.
4. "files": if either file's own "hasUniqueRoleAssignments" reads TRUE, that ONE file has its own
   ACL independent of its folder — inspect it the same way via
   "window.__leakReport.files['<name>']", since re-running reconciliation on the folder will not
   touch a file-level override.
`);
})();

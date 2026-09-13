/*
 * diagnose-blank-libraries.js — READ ONLY. Browser console, run as a SITE ADMINISTRATOR.
 *
 * WHY THIS EXISTS
 * ----------------
 * After a Folder Reconciliation run, all four CRS document libraries (Approval for Document,
 * Approval for Highly Confidential Document, Restricted & Confidential Document, Highly
 * Confidential Document) appear BLANK to an uploader — survives a hard refresh and an incognito
 * window, which rules out a stale browser tab/bundle. That leaves a real ACL problem: either
 * inheritance was broken and the uploader's group was never re-granted, or something stripped a
 * grant the group already held.
 *
 * This dumps the LIVE role assignments on the WEB and on all four libraries, using the single
 * authoritative REST query this project's own naming.ts calls out for exactly this class of
 * question: `RootFolder/ServerRelativeUrl` never changes when a library is retitled, so libraries
 * are matched by URL SEGMENT, never by title.
 *
 * HOW TO RUN
 *   1. Open the site in a browser, signed in as a SITE ADMINISTRATOR (not the affected uploader —
 *      an ordinary account often cannot read role assignments at all).
 *   2. F12 -> Console -> paste this whole file -> Enter.
 *   3. Copy the printed JSON and read it against the checklist at the bottom of this file.
 *
 * ⚠ EVERY REQUEST IS A GET. It creates nothing, changes nothing, deletes nothing.
 *
 * ⚠ `null` MEANS "COULD NOT READ", NEVER "NOT THERE". An absent thing reports an empty array or
 *   `found: false`; an unreadable one reports `null` with the HTTP status. Do not collapse them.
 */
(async () => {
  /* Same pattern as dump-site-inventory.js: never rely on `_spPageContextInfo`, it is undefined on
     several modern pages and in any non-top frame. Set SITE_OVERRIDE if this ever misreads. */
  const SITE_OVERRIDE = ""; // e.g. "https://dcidigitalcom.sharepoint.com/sites/CRS"
  const derived = (() => {
    const m = location.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    return location.origin + (m ? m[1] : "");
  })();
  const web = (SITE_OVERRIDE || derived).replace(/\/+$/, "");
  console.log("reading site:", web || "(root)");

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
  // tree — on a site with hundreds of groups (SDG's), the web-level ACL is almost entirely this
  // noise and drowns the one thing worth reading. Real roles (Read, Edit, CRS Upload, CRS Approve,
  // CRS Delete, Full Control, ...) survive this filter; a pure Limited-Access row does not.
  const realRolesOnly = (rows) =>
    rows.filter((r) => r.roles.some((x) => x !== "Limited Access"));

  const out = { site: web || "(root)" };

  // ── The WEB's own role assignments, FILTERED to real roles only. This is the ancestor-browse
  //    corridor: every ordinary uploader/approver group needs Read HERE (or up their own path) or
  //    every library reads empty, whatever that library's own ACL says. ─────────────────────────
  const webRoles = await get(
    "/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/LoginName,RoleDefinitionBindings/Name",
  );
  out.web = {
    realRoleAssignments:
      webRoles.__error !== undefined ? { __error: webRoles.__error } : realRolesOnly(rolesFor(webRoles.value)),
  };

  // ── Every document library on the site, resolved by BaseTemplate 101 + RootFolder — the same
  //    authoritative read naming.ts's own comment names as the only trustworthy source. ────────
  const libs = await get(
    "/_api/web/lists?$select=Id,Title,HasUniqueRoleAssignments&$expand=RootFolder&$filter=BaseTemplate eq 101",
  );
  if (libs.__error !== undefined) {
    out.libraries = { __error: libs.__error };
    console.log(JSON.stringify(out, null, 2));
    return;
  }

  // URL segments this project documents as STABLE across every rename — never match by title.
  const targets = [
    { key: "Staging (approval)", segment: "ApprovalDocument" },
    { key: "StagingHC (HC approval)", segment: "HCApprovalDocument" },
    { key: "Documents", segment: "Shared Documents" },
    { key: "DocumentsHC", segment: "HCDocuments" },
  ];

  out.libraries = {};
  for (const t of targets) {
    const match = (libs.value || []).find((l) => {
      const u = (l.RootFolder && l.RootFolder.ServerRelativeUrl) || "";
      return u.toLowerCase().endsWith("/" + t.segment.toLowerCase());
    });
    if (!match) {
      out.libraries[t.key] = { found: false, expectedUrlSegment: t.segment };
      continue;
    }
    const acl = await get(
      `/_api/web/lists(guid'${match.Id}')/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/LoginName,RoleDefinitionBindings/Name`,
    );
    out.libraries[t.key] = {
      found: true,
      liveTitle: match.Title,
      urlSegment: t.segment,
      hasUniqueRoleAssignments: match.HasUniqueRoleAssignments,
      roleAssignments: acl.__error !== undefined ? { __error: acl.__error } : rolesFor(acl.value),
    };
  }

  console.log(JSON.stringify(out, null, 2));
  console.log(`
── READ IT AGAINST THIS ──
1. Under "web.roleAssignments": is there a group ending "_SITE_MEMBERS" (or similar) holding
   "Read"? If it is MISSING or its role list is empty, the ancestor-browse corridor is gone and
   every library reads empty for every non-admin — regardless of the library's own ACL.
2. Under each library: does "hasUniqueRoleAssignments" read true? All four SHOULD (reconciliation
   deliberately breaks inheritance on them — an INHERITING CRS library is itself a leak, not a fix).
3. Under each library's "roleAssignments": is the uploader's OWN group there (e.g.
   "<SEGMENT>_<DEPT>_<UNIT>_UPLOADER" or "..._UPLOADER_HIGHLY_CONFIDENTIAL"), holding "CRS Upload"
   or "Read"? Look it up by name in Group Management first if you are not sure of it.
   - Present with the right role -> the library ACL is fine; look elsewhere (the uploader's own
     session, or which library they actually mean).
   - MISSING entirely -> reconciliation did not grant it. Re-run Folder Reconciliation and diff.
   - Present but role list is EMPTY or only "Limited Access" -> a grant was stripped, not skipped.
`);
})();

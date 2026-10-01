/*
 * dump-site-inventory.js — READ ONLY. Browser console, run on ONE site at a time.
 *
 * WHY THIS EXISTS
 * ---------------
 * Porting the system to a second site means knowing what that site already has. Reconstructing
 * that from CLAUDE.md is exactly how this project has repeatedly acted on a stale claim — so this
 * asks the site instead. Run it on BOTH sites, paste both outputs, and diff them.
 *
 * HOW TO RUN
 *   1. Open the site in a browser, signed in as an administrator.
 *   2. F12 -> Console -> paste this whole file -> Enter.
 *   3. Copy the JSON it prints.
 *
 * ⚠ EVERY REQUEST IS A GET. It creates nothing, changes nothing, deletes nothing.
 *
 * ⚠ `null` MEANS "COULD NOT READ", NEVER "NOT THERE". A failed read and an absent object are
 *   different facts and this project has conflated them repeatedly — an absent thing reports
 *   `present:false`, an unreadable one reports `null`. Do not collapse them when reading output.
 *
 * ⚠ LIBRARIES ARE ADDRESSED BY URL SEGMENT, NEVER BY TITLE. A rename never touches the URL
 *   (gotcha #12) and this client renames libraries routinely — three times in two days in August.
 *   Titles are REPORTED so the diff can show them, but never used to look anything up.
 */
(async () => {
  /* ⚠ DO NOT REINTRODUCE `_spPageContextInfo`. It is undefined on several modern pages and in any
     non-top frame, which is how the first version of this script died with
     "ReferenceError: _spPageContextInfo is not defined". The address bar is always there.
     Set SITE_OVERRIDE if a site ever sits at a path this pattern does not match. */
  const SITE_OVERRIDE = ""; // e.g. "https://dcidigitalcom.sharepoint.com/sites/CRS"

  const derived = (() => {
    const m = location.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    return location.origin + (m ? m[1] : ""); // no match => root-hosted web
  })();
  const web = (SITE_OVERRIDE || derived).replace(/\/+$/, "");
  const rel = web.slice(location.origin.length);
  console.log("reading site:", web || "(root)");

  const get = async (url) => {
    try {
      const r = await fetch(web + url, {
        headers: { Accept: "application/json;odata=nometadata" },
      });
      if (!r.ok) return { __error: r.status };
      return await r.json();
    } catch (e) {
      return { __error: String(e && e.message ? e.message : e) };
    }
  };

  // Read-and-empty vs not-read. See the header warning.
  const rows = async (url) => {
    const j = await get(url);
    if (j.__error !== undefined) return null;
    return j.value || [];
  };

  /* Libraries by URL SEGMENT. `Shared Documents` is the normal Documents library — its title has
     been renamed at least twice and its segment never has. */
  const LIB_SEGMENTS = [
    "ApprovalDocument",
    "HCApprovalDocument",
    "Shared Documents",
    "HCDocuments",
    "Archive",
    "HCArchive",
  ];

  /* Columns whose ABSENCE has actually caused a silent failure in this project. Not an exhaustive
     field dump — that would be thousands of lines and unreadable. */
  const WANT_COLUMNS = [
    "Document_x0020_Type", "Year", "DocumentDate", "Confidentiality_x0020_Level",
    "LegallyPrivileged", "Remark", "Vendor_x002f_CustomerName", "_ExtendedDescription",
    "Business_x0020_Segment", "BusinessSegmentTid", "Department", "DepartmentTid",
    "Unit", "UnitTid", "SubUnit", "SubUnitTid",
    "SubmissionId", "BatchId", "SubmissionFileId", "BulkImport", "Archived",
    "ApprovedBy", "Keyword", "Full_x0020_Name", "FullName0",
  ];

  const LISTS = [
    "CRS Config", "DMS Config",
    "CRS Group Map", "DMS Group Map",
    "CRS Folder Map", "DMS Folder Map",
    "CRS Term Abbreviation", "DMS Term Abbreviation",
    "CRS Requests", "CRS Audit Log", "CRS Submissions",
  ];

  const out = { site: web, ranAt: new Date().toISOString() };

  /* ---- libraries: settings + which wanted columns exist ------------------------------------ */
  out.libraries = {};
  for (const seg of LIB_SEGMENTS) {
    const u = `${rel}/${seg}`.replace(/\/+/g, "/");
    /* ⚠ `Id` IS NOT DECORATION — it is what makes a flow port possible. Every SharePoint action in
       the 21 flows stores its list as a GUID, and those GUIDs are per-SITE, so porting means
       mapping each one to a list by name on the target. Without the Id here that mapping is
       guesswork. */
    const meta = await get(
      `/_api/web/getlist(@u)?@u='${encodeURIComponent(u)}'` +
        `&$select=Id,Title,ItemCount,EnableModeration,DraftVersionVisibility,` +
        `EnableVersioning,MajorVersionLimit,EnableMinorVersions`,
    );
    if (meta.__error !== undefined) {
      // 404 here means the library genuinely is not at that URL — a real answer, not a failure.
      out.libraries[seg] =
        meta.__error === 404 ? { present: false } : { present: null, error: meta.__error };
      continue;
    }
    const fields = await rows(
      `/_api/web/getlist(@u)/fields?@u='${encodeURIComponent(u)}'` +
        `&$select=InternalName&$top=5000`,
    );
    out.libraries[seg] = {
      present: true,
      id: meta.Id,
      title: meta.Title,
      items: meta.ItemCount,
      contentApproval: meta.EnableModeration,
      draftVisibility: meta.DraftVersionVisibility, // 0 any reader, 1 author+editor, 2 approver+author
      versioning: meta.EnableVersioning,
      majorVersionLimit: meta.MajorVersionLimit,
      minorVersions: meta.EnableMinorVersions,
      // null propagates: an unreadable /fields must not read as "no columns".
      columnsPresent: fields === null ? null
        : WANT_COLUMNS.filter((c) => fields.some((f) => f.InternalName === c)),
      columnsMissing: fields === null ? null
        : WANT_COLUMNS.filter((c) => !fields.some((f) => f.InternalName === c)),
      fieldCount: fields === null ? null : fields.length,
    };
  }

  /* ---- lists: presence + item counts ------------------------------------------------------- */
  out.lists = {};
  for (const t of LISTS) {
    const j = await get(
      `/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=Id,Title,ItemCount,Hidden`,
    );
    out.lists[t] =
      j.__error === 404 ? { present: false }
      : j.__error !== undefined ? { present: null, error: j.__error }
      : { present: true, id: j.Id, items: j.ItemCount, hidden: j.Hidden };
  }

  /* ---- config rows: the actual payload of a port ------------------------------------------- */
  /* ⚠ THREE reads, not one. `Levels`/`PendingLevels` are created ON DEMAND, and ONE unknown field
     name fails the WHOLE $select (gotcha #11) — which is what made a perfectly good site report
     "the segment list could not be read" on 2026-08-18. Settings, modes and pending, separately. */
  const cfgTitle = (out.lists["CRS Config"] || {}).present ? "CRS Config" : "DMS Config";
  out.configListUsed = cfgTitle;
  out.configSettings = await rows(
    `/_api/web/lists/getbytitle('${encodeURIComponent(cfgTitle)}')/items` +
      `?$select=Title,SettingValue&$top=500`,
  );
  out.configModes = await rows(
    `/_api/web/lists/getbytitle('${encodeURIComponent(cfgTitle)}')/items` +
      `?$select=Title,ModeLabel,Category,TermSetGuid,StagingFolder,SortOrder,Levels&$top=500`,
  );
  out.configPending = await rows(
    `/_api/web/lists/getbytitle('${encodeURIComponent(cfgTitle)}')/items` +
      `?$select=Title,PendingLevels&$top=500`,
  );

  /* ---- permission levels ------------------------------------------------------------------- */
  const defs = await rows(`/_api/web/roledefinitions?$select=Name&$top=200`);
  out.permissionLevels = defs === null ? null : defs.map((d) => d.Name);

  /* ---- pages: the client renames these, so the diff needs the real file names --------------- */
  const pages = await rows(
    `/_api/web/lists/getbytitle('Site Pages')/items?$select=FileLeafRef&$top=500`,
  );
  out.pages = pages === null ? null : pages.map((p) => p.FileLeafRef).sort();

  /* ---- groups ------------------------------------------------------------------------------ */
  /* ⚠ KEYSET-PAGED, and $top alone is NOT enough. `/sitegroups` has no continuation link at any
     metadata level, so `$top` simply TRUNCATES — and a truncated read is indistinguishable from
     the groups being absent (1.0.443.0). One provisioned segment is ~320 groups. */
  let groups = [];
  let last = 0;
  let capped = false;
  for (let i = 0; i < 20; i++) {
    const page = await rows(
      `/_api/web/sitegroups?$select=Id,Title&$orderby=Id&$filter=Id gt ${last}&$top=500`,
    );
    if (page === null) { groups = null; break; }
    groups = groups.concat(page);
    if (page.length < 500) break;
    last = page[page.length - 1].Id;
    if (i === 19) capped = true;
  }
  out.groupCount = groups === null ? null : groups.length;
  out.groupCountCapped = capped;
  out.siteEntryGroup = groups === null ? null
    : (groups.find((g) => /_SITE_MEMBERS$/i.test(g.Title)) || {}).Title || false;

  /* ---- site assets / site pages list ACLs --------------------------------------------------- */
  /* Only ever a LOCKOUT here, never a leak — so worth reporting and never worth guessing. Site
     Assets holds the hero banner; Site Pages with unique permissions needs an explicit Read for
     the site-entry group or every INHERITING page denies everyone (2026-08-21). */
  const sa = await get(
    `/_api/web/lists/getbytitle('Site Assets')?$select=Title,HasUniqueRoleAssignments`,
  );
  out.siteAssets = sa.__error !== undefined ? { present: null, error: sa.__error }
    : { present: true, uniquePermissions: sa.HasUniqueRoleAssignments };

  const sp = await get(
    `/_api/web/lists/getbytitle('Site Pages')?$select=Title,HasUniqueRoleAssignments`,
  );
  out.sitePagesList = sp.__error !== undefined ? { present: null, error: sp.__error }
    : { present: true, uniquePermissions: sp.HasUniqueRoleAssignments };

  /* ---- user custom actions ------------------------------------------------------------------ */
  /* ⚠ The two extensions that do NOT provision from elements.xml on this tenant and had to be
     POSTed by hand — the `+ New Folder` customizer and the bulk-approve command set. An empty
     array here on SDG is a real finding, not noise. */
  out.userCustomActions = await rows(
    `/_api/web/usercustomactions?$select=Title,Location,ClientSideComponentId,RegistrationId`,
  );

  /* ⚠ DO NOT `console.log` THE WHOLE OBJECT. Chrome TRUNCATES long console output, and the first
     run of this script lost everything after the libraries section — silently, mid-string, with no
     indication anything was missing. `copy()` is a DevTools function that puts the value on the
     clipboard whole. */
  const json = JSON.stringify(out);
  try {
    copy(json);
    console.log("✅ FULL RESULT COPIED TO CLIPBOARD — just paste it. " + json.length + " chars.");
  } catch (e) {
    console.log("copy() unavailable — falling back to chunks. Paste ALL of them, in order.");
    const size = 15000;
    for (let i = 0; i < json.length; i += size) {
      console.log(`--- chunk ${i / size + 1} of ${Math.ceil(json.length / size)} ---`);
      console.log(json.slice(i, i + size));
    }
  }

  /* A short summary that is readable even when the JSON is on the clipboard rather than on screen. */
  console.log("\n=== summary ===");
  console.log("site:", out.site);
  Object.keys(out.libraries).forEach((k) => {
    const l = out.libraries[k];
    console.log(
      ` lib ${k}: ` +
        (l.present === true
          ? `"${l.title}" id=${l.id} items=${l.items} approval=${l.contentApproval} draft=${l.draftVisibility}` +
            (l.columnsMissing && l.columnsMissing.length
              ? ` MISSING=[${l.columnsMissing.join(",")}]`
              : " (all wanted columns present)")
          : l.present === false
            ? "ABSENT"
            : "COULD NOT READ (" + l.error + ")"),
    );
  });
  Object.keys(out.lists).forEach((k) => {
    const l = out.lists[k];
    if (l.present === true) console.log(` list ${k}: id=${l.id} items=${l.items}`);
    else if (l.present === null) console.log(` list ${k}: COULD NOT READ (${l.error})`);
  });
  console.log(" config list used:", out.configListUsed);
  console.log(" permission levels:", (out.permissionLevels || []).join(", "));
  console.log(" groups:", out.groupCount, "entry group:", out.siteEntryGroup);
  console.log(" site assets unique perms:", (out.siteAssets || {}).uniquePermissions);
  console.log(" site pages unique perms:", (out.sitePagesList || {}).uniquePermissions);
  console.log(" custom actions:", (out.userCustomActions || []).length);
  console.log(" pages:", (out.pages || []).length);
  return out;
})();

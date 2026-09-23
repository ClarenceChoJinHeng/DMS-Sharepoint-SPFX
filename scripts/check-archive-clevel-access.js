/*
 * Does a C-Level (GLOBAL/SEGVIEW) group hold more than "Read" on any Archive/ArchiveHC folder?
 * -----------------------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, renamed, deleted or granted.
 *
 * Run it in the browser console, ON THE CRS SITE, signed in as an ADMINISTRATOR (it reads the
 * whole `CRS Group Map` and every group's title, not just your own).
 *
 * WHY THIS EXISTS: client, 2026-09-23 — "can you check why c level can edit archive folders,
 * shouldn't they only read." Reading the CODE first (`FolderManager.tsx`):
 *
 *   - `LIBRARY_ROLES.Archive` / `.ArchiveHC` = `["GLOBAL", "SEGVIEW"]` ONLY — no other role has
 *     ever been granted anything on these two libraries since they were created (2026-08-22).
 *   - `permissionForRole()`'s `READ_ONLY_LIBS` branch (Archive/ArchiveHC) maps GLOBAL and SEGVIEW
 *     through `ROLE_TO_PERMISSION`, which currently reads `GLOBAL: "Read"` / `SEGVIEW: "Read"` —
 *     reverted from a brief "Restricted View" experiment on 2026-09-22, but NEVER, at any point in
 *     this project's history, anything WIDER than Read. There is no code path that computes
 *     "Edit"/"Contribute"/"CRS Approve" for a C-Level role on these libraries.
 *
 * So as CODE stands today, a fresh Folder Reconciliation run can only ever grant these two roles
 * exactly "Read" here. If a real account is observed EDITING an archive folder, the leading
 * explanations are all LIVE-DATA questions this script exists to answer, not code defects:
 *
 *   1. A STALE GRANT PREDATING THE READ-ONLY RULE (2026-08-22) OR THE FOLLOW-UP SWEEP (2026-09-02/
 *      03, "prune archive access to C-Level only" — built, run once on both sites, then REMOVED).
 *      Folder-scope reconciliation only ever ADDS a missing grant; it never downgrades or removes a
 *      wider one already there (`groupsToRemove`'s full-ACL assertion is deliberately page-scope
 *      only — see CLAUDE.md). If a C-Level group's OWN binding on an archive folder was ever
 *      something other than Read — by hand, or from before this rule existed — nothing since has
 *      corrected it.
 *   2. AN INHERITANCE GAP. CLAUDE.md records that the 2026-09-02/03 sweep found grants "all on the
 *      SEGMENT ROOT — nothing below it needed touching — that whole subtree inherits from the
 *      segment root rather than breaking its own inheritance per department/unit, unlike Documents/
 *      HC Documents." If a folder BELOW a segment root has since had its inheritance broken by hand
 *      (in SharePoint directly, outside this tool), it now has its OWN ACL that nothing here manages
 *      or corrects.
 *   3. THE ACCOUNT TESTED IS ALSO A SITE COLLECTION ADMINISTRATOR / `CRS Owners` MEMBER. Full
 *      Control bypasses every folder-level grant entirely, regardless of what their C-Level role
 *      itself holds — this would look identical to "C-Level can edit" if the same account happens
 *      to be both.
 *
 * WHAT THIS SCRIPT DOES:
 *   1. Reads CRS Config for segment labels (StagingFolder, keyed by TermSetGuid).
 *   2. Reads CRS Group Map, finds every group holding a GLOBAL or SEGVIEW row (the ONLY roles
 *      Archive/ArchiveHC ever admit), and which segment each SEGVIEW row scopes to.
 *   3. Resolves live group titles (never trusts the stored GroupName column, which a rename leaves
 *      stale).
 *   4. Resolves the Archive/ArchiveHC libraries and reads:
 *        a. the LIBRARY ROOT's own role assignments (via the list resource, not the folder resource
 *           — a library root does not expose roleassignments the normal folder way), and
 *        b. every SEGMENT folder's role assignments, and
 *        c. one level DOWN (department folders) under each segment, to catch an inheritance break
 *           nobody expects to be there.
 *   5. Reports, for every REAL binding found (Owners and Limited Access excluded):
 *        - a GLOBAL/SEGVIEW group holding anything OTHER than "Read" — the exact question asked.
 *        - ANY other group (not GLOBAL/SEGVIEW, not Owners) holding anything at all — per
 *          LIBRARY_ROLES this should never happen, and is the shape the 2026-09-02/03 sweep existed
 *          to clear.
 *        - a folder below the segment root that has broken its own inheritance (unexpected).
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"

  /* ── Resolve the site — same three-route fallback every other script here uses ──────────── */
  const resolveWeb = () => {
    if (SITE_OVERRIDE) return SITE_OVERRIDE.replace(/\/+$/, "");
    try {
      if (
        typeof _spPageContextInfo !== "undefined" &&
        _spPageContextInfo.webAbsoluteUrl
      ) {
        return _spPageContextInfo.webAbsoluteUrl.replace(/\/+$/, "");
      }
    } catch (e) {
      /* not defined on this page — fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
    return (location.origin + (m ? m[1] : "")).replace(/\/+$/, "");
  };
  const web = resolveWeb();
  console.log("Site:", web);

  const bust = (u) =>
    (u.indexOf("?") === -1 ? "?" : "&") +
    `_=${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;

  const withRetry = async (fn) => {
    let attempt = 0;
    for (;;) {
      const res = await fn();
      if (res.status !== 429 && res.status !== 503) return res;
      attempt++;
      if (attempt >= 5) return res;
      const ra = Number(res.headers && res.headers.get && res.headers.get("Retry-After"));
      const waitMs = isFinite(ra) && ra > 0 ? ra * 1000 : Math.min(1000 * 2 ** attempt, 15000);
      await new Promise((r) => setTimeout(r, waitMs));
    }
  };

  const get = async (path) => {
    const url = web + path;
    const r = await withRetry(() =>
      fetch(url + bust(url), {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Cache-Control": "no-cache",
          Pragma: "no-cache",
        },
        cache: "no-store",
      }),
    );
    if (r.status === 404) return { __notFound: true };
    const type = r.headers.get("content-type") || "";
    if (type.indexOf("json") === -1) {
      throw new Error(
        `Expected JSON and got "${type}" from ${url}\nThe site resolved to ${web}. If that is ` +
          `wrong, set SITE_OVERRIDE at the top and re-run.`,
      );
    }
    if (!r.ok) throw new Error(`HTTP ${r.status} from ${url}`);
    return r.json();
  };
  const tryGet = async (path) => {
    try {
      return await get(path);
    } catch (e) {
      return { __error: e.message };
    }
  };

  const resolveList = async (candidates) => {
    for (const t of candidates) {
      const probe = await tryGet(
        `/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=Title,Id,ItemCount`,
      );
      if (probe && probe.Title) return probe;
    }
    return undefined;
  };

  /* Also pulls the library's own GUID (for the list-scope roleassignments read, needed for the
     ROOT — a library root does not expose roleassignments via the folder resource) and its root
     server-relative URL (for walking segment/department folders underneath it). */
  const resolveLibrary = async (candidates) => {
    for (const t of candidates) {
      const probe = await tryGet(
        `/_api/web/lists/getbytitle('${encodeURIComponent(t)}')` +
          `?$expand=RootFolder&$select=Title,Id,RootFolder/ServerRelativeUrl`,
      );
      if (probe && probe.Title && probe.RootFolder) {
        return { title: probe.Title, id: probe.Id, rootUrl: probe.RootFolder.ServerRelativeUrl };
      }
    }
    return undefined;
  };

  /* ── Role normalisation — the same rules `normalizeRoleValue` uses in groupMapModel.ts,
     copied verbatim from check-stale-delete-roles.js so the two agree on what a raw Role value
     means. ────────────────────────────────────────────────────────────────────────────────── */
  const SHORT_ROLE_CODES = [
    "MEMBER", "UPL", "APR", "DEL", "DELS", "SEGVIEW", "UPLHC", "APRHC",
    "DELSHC", "DELHC", "SHAREHC", "MEMBERHC", "DEPTVIEW", "SHARE", "GLOBAL", "ENTRY",
  ];
  const ROLE_ALIASES = {
    UPLOADER: "UPL",
    APPROVER: "APR",
    DELETER_DOCUMENTS: "DEL",
    DELETER_STAGING: "DELS",
    MEMBER: "MEMBER",
    VIEWER: "MEMBER",
    SEGMENTVIEW: "SEGVIEW",
    DEPARTMENTVIEW: "DEPTVIEW",
    VIEWER_HC: "MEMBERHC",
  };
  const normalizeRoleValue = (raw) => {
    const v = (raw ?? "").trim().toUpperCase().replace(/[\s-]+/g, "_");
    if (v.length === 0) return "";
    if (SHORT_ROLE_CODES.indexOf(v) !== -1) return v;
    return ROLE_ALIASES[v] ?? v;
  };

  /* The ONLY two roles `LIBRARY_ROLES.Archive`/`.ArchiveHC` admit — see the header comment. */
  const CLEVEL_ROLES = ["GLOBAL", "SEGVIEW"];

  /* ── 1. Segments — for labelling only ──────────────────────────────────────────────────── */
  console.log("\n=== 1. Segments (CRS Config) ===");
  const configList = await resolveList(["CRS Config", "DMS Config"]);
  if (!configList) {
    console.log("  ✗ Could not find CRS Config / DMS Config — cannot label segments, aborting.");
    return;
  }
  const modeRes = await tryGet(
    `/_api/web/lists/getbytitle('${encodeURIComponent(configList.Title)}')/items` +
      `?$select=Id,ModeLabel,StagingFolder,TermSetGuid,SortOrder,Levels,ConfigType` +
      `&$filter=ConfigType eq 'mode'&$orderby=SortOrder&$top=5000`,
  );
  const allModeRows = (modeRes && modeRes.value) || [];
  const codeByGuid = {};
  allModeRows
    .filter((r) => (r.TermSetGuid || "").trim() && (r.Levels || "").trim())
    .forEach((r) => {
      const g = (r.TermSetGuid || "").trim().toLowerCase();
      if (codeByGuid[g] !== undefined) return;
      codeByGuid[g] = (r.StagingFolder || "").trim();
    });
  console.log(`  ${Object.keys(codeByGuid).length} segment code(s) resolved.`);

  /* ── 2. Group Map rows — find every GLOBAL/SEGVIEW-holding group and its segment scope ───── */
  console.log("\n=== 2. Group Map rows (CRS Group Map) ===");
  const gmList = await resolveList(["CRS Group Map", "DMS Group Map"]);
  if (!gmList) {
    console.log("  ✗ Could not find CRS Group Map / DMS Group Map — aborting.");
    return;
  }
  const gmRes = await tryGet(
    `/_api/web/lists/getbytitle('${encodeURIComponent(gmList.Title)}')/items` +
      `?$select=GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
  );
  if (!gmRes || gmRes.__error) {
    console.log("  ✗ Could not read the Group Map:", gmRes && gmRes.__error);
    return;
  }
  const gmRows = gmRes.value || [];
  console.log(
    `  ${gmRows.length} row(s) read` +
      (gmRows.length >= 5000
        ? "  ⚠ AT THE CAP — some rows may be missing, page this manually"
        : "."),
  );

  const clevelByGroup = {}; // GroupId -> { roles: Set, segCodes: Set }
  gmRows.forEach((r) => {
    const gid = (r.GroupId ?? "").trim();
    const role = normalizeRoleValue(r.Role ?? "");
    if (!gid || CLEVEL_ROLES.indexOf(role) === -1) return;
    const entry =
      clevelByGroup[gid] || (clevelByGroup[gid] = { roles: new Set(), segCodes: new Set() });
    entry.roles.add(role);
    // GLOBAL rows carry no Segment (they reach every segment); SEGVIEW rows do.
    const segGuid = (r.Segment ?? "").trim().toLowerCase();
    if (segGuid && codeByGuid[segGuid]) entry.segCodes.add(codeByGuid[segGuid]);
  });
  console.log(`  ${Object.keys(clevelByGroup).length} group(s) hold GLOBAL and/or SEGVIEW.`);

  /* ── 3. Live SharePoint group titles ──────────────────────────────────────────────────────
     Never trust the stored GroupName column — a rename leaves it stale (the 1.0.162.0 lesson);
     only GroupId, joined here, survives a rename. */
  console.log("\n=== 3. Live group titles (sitegroups) ===");
  const allGroupsRes = await tryGet(`/_api/web/sitegroups?$select=Id,Title&$top=5000`);
  const allGroups = (allGroupsRes && allGroupsRes.value) || [];
  console.log(
    `  ${allGroups.length} group(s) on the site` +
      (allGroups.length >= 5000 ? "  ⚠ AT THE CAP — page this manually" : "."),
  );
  const titleById = {};
  allGroups.forEach((g) => {
    titleById[String(g.Id)] = g.Title;
  });
  const clevelGroupIds = new Set(Object.keys(clevelByGroup));

  /* ── 4. Resolve the Archive / ArchiveHC libraries — candidate lists copied verbatim from
     shared/naming.ts (ARCHIVE_CANDIDATES / ARCHIVE_HC_CANDIDATES). This client renames libraries
     routinely; an unlisted title fails SILENTLY (the library is skipped as "not present"). ────── */
  console.log("\n=== 4. Resolving the archive libraries ===");
  const ARCHIVE_CANDIDATES = ["Archive", "Archive Restricted & Confidential Document", "CRSArchive"];
  const ARCHIVE_HC_CANDIDATES = [
    "Archive Highly Confidential Document", "HC Archive", "HC Archives", "HCArchive",
    "Highly Confidential Archive Document", "Highly Confidential Archive",
  ];
  const libraries = [
    { key: "Archive", candidates: ARCHIVE_CANDIDATES },
    { key: "ArchiveHC", candidates: ARCHIVE_HC_CANDIDATES },
  ];
  for (const lib of libraries) {
    lib.resolved = await resolveLibrary(lib.candidates);
    console.log(
      lib.resolved
        ? `  ${lib.key}: "${lib.resolved.title}" (${lib.resolved.rootUrl})`
        : `  ${lib.key}: not present on this site (normal if this site has no HC pair / no archive yet).`,
    );
  }
  const resolvedLibraries = libraries.filter((l) => l.resolved);
  if (resolvedLibraries.length === 0) {
    console.log("\n✗ No archive library resolved on this site — nothing to check.");
    return;
  }

  /* ── 5. Read live folder role assignments and classify every REAL binding ────────────────── */
  console.log("\n=== 5. Walking live archive folder permissions ===");

  const OWNER_BINDING_NAMES = ["Full Control"]; // site Owners; excluded by design everywhere else
  // RoleTypeKind 1 = Guest, 5 = WebAnalyticsAdmin, 7 = "Limited Access" auto-entries — never a
  // real grant anyone chose. Matched on the binding's OWN RoleTypeKind, not the level's name,
  // since "Limited Access" itself can also appear as a named level.
  const isRealBinding = (b) => {
    if (b.RoleTypeKind === 1 || b.RoleTypeKind === 7) return false;
    if (OWNER_BINDING_NAMES.indexOf(b.Name) !== -1) return false;
    return true;
  };

  // Library-ROOT roleassignments must go through the LIST resource, not the folder resource — a
  // library root does not expose `ListItemAllFields/roleassignments` the normal folder way.
  const libraryRootAssignments = async (listId) => {
    const j = await tryGet(
      `/_api/web/lists(guid'${listId}')/roleassignments` +
        `?$expand=Member,RoleDefinitionBindings` +
        `&$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind`,
    );
    if (!j || j.__error || j.__notFound) return undefined;
    return j.value || [];
  };
  const folderAssignments = async (path) => {
    const j = await tryGet(
      `/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments` +
        `?$expand=Member,RoleDefinitionBindings` +
        `&$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind` +
        `&@f='${encodeURIComponent(path)}'`,
    );
    if (!j || j.__error || j.__notFound) return undefined;
    return j.value || [];
  };
  const uniquePerms = async (path) => {
    const j = await tryGet(
      `/_api/web/GetFolderByServerRelativeUrl(@f)?$select=HasUniqueRoleAssignments&@f='${encodeURIComponent(path)}'`,
    );
    if (!j || j.__error || j.__notFound) return undefined;
    return j.HasUniqueRoleAssignments;
  };
  const childFolders = async (path) => {
    const j = await tryGet(
      `/_api/web/GetFolderByServerRelativeUrl(@f)/Folders?$select=Name,ServerRelativeUrl&@f='${encodeURIComponent(path)}'`,
    );
    if (!j || j.__error || j.__notFound) return undefined;
    return j.value || [];
  };

  const wrongLevelFindings = []; // GLOBAL/SEGVIEW group holding something other than Read
  const strayGroupFindings = []; // any OTHER (non-C-Level, non-Owner) group holding anything
  const unexpectedInheritanceBreaks = []; // a folder below the segment root with its own ACL

  /* ⚠⚠ FIXED 2026-09-23, AFTER A LIVE RUN PRODUCED A FALSE POSITIVE — every "wrong level" hit came
     back as `Read, Restricted View`, and the FIRST version of this check treated anything other
     than "exactly one binding named Read" as a finding. That is wrong: `Restricted View` is a
     built-in SharePoint level that is NARROWER than Read (it routes viewing through Office
     Online's own renderer with no download/Open Items right — see `ROLE_TO_PERMISSION`'s own
     comment in `FolderManager.tsx`), not wider. This project's own documented history explains the
     PAIR exactly: MEMBER/GLOBAL/SEGVIEW briefly used "Restricted View" (2026-09-14/15) before being
     reverted back to "Read" (2026-09-22) — and because folder-scope reconciliation only ever ADDS a
     grant and never removes one, the OLD "Restricted View" binding is still sitting on every
     already-provisioned folder alongside the NEW "Read" one. SharePoint UNIONS role assignments, so
     the effective right is Read either way; having both denies nothing and grants nothing extra.
     Flagging that pair as "more than Read" was a bug in THIS SCRIPT, not a finding about the site. */
  const HARMLESS_LEVELS = ["Read", "Restricted View"];

  const classifyAssignments = (libKey, where, assignments) => {
    if (!assignments) return;
    for (const ra of assignments) {
      const bindings = (ra.RoleDefinitionBindings || []).filter(isRealBinding);
      if (bindings.length === 0) continue; // Owners / Limited Access only — nothing to report
      const gid = String(ra.PrincipalId ?? "");
      const title = (ra.Member && ra.Member.Title) || titleById[gid] || `id ${gid}`;
      const levelNames = bindings.map((b) => b.Name).sort().join(", ");
      if (clevelGroupIds.has(gid)) {
        const allHarmless = bindings.every((b) => HARMLESS_LEVELS.indexOf(b.Name) !== -1);
        if (!allHarmless) {
          wrongLevelFindings.push({ libKey, where, gid, title, levelNames });
        }
      } else {
        strayGroupFindings.push({ libKey, where, gid, title, levelNames });
      }
    }
  };

  const MAX_WALK_REQUESTS = 4000;
  let budget = MAX_WALK_REQUESTS;
  const spend = (n) => {
    budget -= n;
    return budget > 0;
  };
  let checked = 0;
  const heartbeat = () => {
    checked++;
    if (checked % 25 === 0) {
      console.log(
        `  … ${checked} folder(s)/checks so far — ${wrongLevelFindings.length + strayGroupFindings.length} finding(s), ${budget} request(s) left`,
      );
    }
  };

  for (const lib of resolvedLibraries) {
    console.log(`\n--- ${lib.key} ("${lib.resolved.title}") ---`);

    // 5a. Library root.
    const rootAssignments = await libraryRootAssignments(lib.resolved.id);
    spend(1);
    heartbeat();
    classifyAssignments(lib.key, "(library root)", rootAssignments);

    // 5b/5c. Every segment folder directly under the root, and one level of department folders
    // beneath each — the segment root is where reconciliation is DOCUMENTED to break inheritance
    // for this library (everything deeper is expected to inherit); checking one level further
    // down catches an unexpected, manually-broken inheritance point.
    const segFolders = await childFolders(lib.resolved.rootUrl);
    spend(1);
    heartbeat();
    if (segFolders === undefined) {
      console.log(`  (could not list segment folders under ${lib.resolved.rootUrl})`);
      continue;
    }
    for (const seg of segFolders) {
      if (budget <= 0) break;
      const segAssignments = await folderAssignments(seg.ServerRelativeUrl);
      spend(1);
      heartbeat();
      classifyAssignments(lib.key, seg.Name, segAssignments);

      const deptFolders = await childFolders(seg.ServerRelativeUrl);
      spend(1);
      heartbeat();
      if (deptFolders === undefined) continue;
      for (const dept of deptFolders) {
        if (budget <= 0) break;
        const isUnique = await uniquePerms(dept.ServerRelativeUrl);
        spend(1);
        heartbeat();
        if (isUnique === true) {
          unexpectedInheritanceBreaks.push({
            libKey: lib.key,
            path: `${seg.Name}/${dept.Name}`,
          });
          const deptAssignments = await folderAssignments(dept.ServerRelativeUrl);
          spend(1);
          heartbeat();
          classifyAssignments(lib.key, `${seg.Name}/${dept.Name}`, deptAssignments);
        }
      }
    }
    if (budget <= 0) {
      console.log(
        `  ⚠ HIT THE ${MAX_WALK_REQUESTS}-REQUEST BUDGET partway through ${lib.key} — not every ` +
          `folder was checked. Re-run with a higher MAX_WALK_REQUESTS, or scope to one segment, ` +
          `before treating this as a complete sweep.`,
      );
    }
  }

  /* ── 6. Report ────────────────────────────────────────────────────────────────────────────── */
  console.log("\n=== REPORT ===");

  console.log(
    `\n--- C-Level (GLOBAL/SEGVIEW) groups holding MORE than "Read" (${wrongLevelFindings.length}) ---`,
  );
  if (wrongLevelFindings.length === 0) {
    console.log('  ✅ NONE FOUND. Every GLOBAL/SEGVIEW group checked holds exactly "Read", nowhere.');
  } else {
    console.log("  ⚠⚠ THIS IS THE EXACT SYMPTOM ASKED ABOUT — an account behind one of these can");
    console.log("     edit, not just read, the folder named:");
    wrongLevelFindings.forEach((f) => {
      console.log(`    [${f.libKey}] ${f.where} — "${f.title}" (id ${f.gid}): ${f.levelNames}`);
    });
  }

  console.log(
    `\n--- OTHER groups (not GLOBAL/SEGVIEW, not Owners) holding ANY grant (${strayGroupFindings.length}) ---`,
  );
  console.log(
    "  (Per LIBRARY_ROLES.Archive/.ArchiveHC, nothing but GLOBAL/SEGVIEW should ever be granted",
  );
  console.log(
    "   here at all — but a group holding only Read/Restricted View is not an escalation on its",
  );
  console.log(
    "   own, since that is the SAME level GLOBAL/SEGVIEW themselves get. It is only worth chasing",
  );
  console.log(
    "   for TWO reasons: (a) it may be a group your Group Map does not actually list as GLOBAL/",
  );
  console.log(
    "   SEGVIEW despite its name implying it should be (a classification question, not a leak),",
  );
  console.log("   or (b) it holds something wider than Read/Restricted View, which IS a real leak.");
  if (strayGroupFindings.length === 0) {
    console.log("  (none)");
  } else {
    strayGroupFindings.forEach((f) => {
      const bindingNames = f.levelNames.split(", ");
      const wider = bindingNames.some((n) => HARMLESS_LEVELS.indexOf(n) === -1);
      console.log(
        `    [${f.libKey}] ${f.where} — "${f.title}" (id ${f.gid}): ${f.levelNames}` +
          (wider
            ? "  ⚠⚠ WIDER THAN Read/Restricted View — this IS an escalation"
            : "  (Read/Restricted View only — not an escalation, see note above)"),
      );
    });
  }

  console.log(
    `\n--- Department folders that unexpectedly broke their own inheritance (${unexpectedInheritanceBreaks.length}) ---`,
  );
  console.log(
    "  (Documented behaviour: everything below the SEGMENT root inherits from it. A folder here",
  );
  console.log("   has its own ACL for some other reason — checked above for what it actually holds.)");
  if (unexpectedInheritanceBreaks.length === 0) {
    console.log("  (none)");
  } else {
    unexpectedInheritanceBreaks.forEach((f) => console.log(`    [${f.libKey}] ${f.path}`));
  }

  console.log("\n=== Summary ===");
  console.log(`  ${checked} folder(s)/checks performed.`);
  // `wrongLevelFindings` is already filtered to "wider than Read/Restricted View" — that check was
  // fixed 2026-09-23 after a false positive (see HARMLESS_LEVELS' own comment above). A stray group
  // only counts as a REAL escalation if it ALSO holds something wider; a stray Read/Restricted-View-
  // only presence is a classification question, not a leak.
  const strayEscalations = strayGroupFindings.filter((f) =>
    f.levelNames.split(", ").some((n) => HARMLESS_LEVELS.indexOf(n) === -1),
  );
  const realEscalations = wrongLevelFindings.length + strayEscalations.length;
  if (realEscalations === 0) {
    console.log(
      "\n✅ NOTHING WIDER THAN Read/Restricted View WAS FOUND ON THE ARCHIVE FOLDERS THEMSELVES" +
        (strayGroupFindings.length > 0
          ? ` (${strayGroupFindings.length} stray Read/Restricted-View-only group presence(s) ` +
            "listed above — worth understanding for classification hygiene, not urgent)."
          : "."),
    );
    console.log(
      "   If a real account still shows edit-like behaviour on an archive document, check next:",
    );
    console.log(
      "   (a) is that account ALSO a member of `CRS Owners` or a Site Collection Administrator —",
    );
    console.log(
      "   Full Control bypasses every folder grant regardless of any C-Level role; (b) is what was",
    );
    console.log(
      "   actually tested the app's own approver/admin screens (a different permission entirely",
    );
    console.log("   from the raw folder ACL this script checked) rather than a direct SharePoint edit.");
  } else {
    console.log(
      `\n⚠⚠ ${realEscalations} REAL ESCALATION(S) FOUND — genuinely wider than Read/Restricted View.`,
    );
    console.log(
      "⏭ NEXT: for each one, delete-and-recreate is the durable fix (SharePoint drops a deleted",
    );
    console.log(
      "  principal's role assignments; unticking the level by hand gets silently reissued by the",
    );
    console.log(
      "  next Folder Reconciliation run, since folder-scope reconciliation only ever ADDS a missing",
    );
    console.log("  grant and never downgrades or removes one already there).");
  }
})();

/*
 * repair-stale-delete-roles.js — companion to check-stale-delete-roles.js. Browser console, run
 * as a SITE ADMINISTRATOR (removing role assignments and deleting list rows needs Manage
 * Permissions / Manage Lists, which an ordinary uploader/approver account does not hold).
 *
 * ⚠⚠ THIS SCRIPT WRITES. DEFAULTS TO A DRY RUN. Read the report it prints, and only set
 * `DO_WRITE = true` below and run it again once you are satisfied every flagged item is real.
 *
 * WHY THIS EXISTS
 * ----------------
 * `check-stale-delete-roles.js` found 393 of 857 mapped groups (across GHO, MHO, TNS) still
 * carrying a `DEL`/`DELS`/`DELHC`/`DELSHC` role that no persona has been granted since the
 * 2026-09-17 proxy-deletion redesign (`shared/groupMapModel.ts`). Every one of those groups can
 * still delete files AND FOLDERS directly on the live site, bypassing the proxy-deletion mechanism
 * entirely.
 *
 * The documented repair (CLAUDE.md, "STALE APPROVER GROUPS STILL GRANT CRS Delete") is
 * delete-the-group-and-recreate-it — durable, but it loses every group's membership (must be
 * exported and re-added by hand) and briefly drops every OTHER role the group correctly holds
 * (APR/UPL/SHARE) until the next reconciliation run puts them back.
 *
 * THIS SCRIPT IS A LIGHTER ALTERNATIVE: it removes only the ONE thing that is actually wrong —
 *   1. The stale `DEL`/`DELS`/`DELHC`/`DELSHC` ROW on each group's `CRS Group Map` entry (its
 *      other rows — APR, UPL, SHARE, whatever else it correctly holds — are left completely
 *      alone), so Folder Reconciliation stops re-granting the permission on its next run.
 *   2. The LIVE `CRS Delete` (or `DMS Delete`) role-assignment BINDING those groups currently hold
 *      on department/unit folders — using the same role-definition-specific removal
 *      (`removeroleassignment(principalid=X,roledefid=Y)`) this project's own `FolderManager.tsx`
 *      already uses for exactly this class of repair (`removeSingleRoleBinding`, kept "unreachable
 *      on purpose" there after its own one-time archive-access-narrowing use on 2026-09-02/03).
 *      Any OTHER permission level the same group holds on the same folder (CRS Upload, CRS
 *      Approve, CRS Share) is UNTOUCHED — this removes one binding out of potentially several on
 *      the same role assignment, never the whole thing.
 * No group is deleted. No membership is lost. No other grant is disturbed.
 *
 * WHAT IT DOES NOT DO
 *   - It does not touch groups whose role set matches no persona for OTHER reasons (section 5 of
 *     the check script) — only the ones carrying a stale DELETE role.
 *   - It does not touch Archive/ArchiveHC — those are already narrowed to C-Level only
 *     (`LIBRARY_ROLES.Archive`) and were already swept once (2026-09-02/03).
 *   - It does not recreate anything reconciliation would otherwise have granted — if a group's
 *     ENTIRE role set was just the stale delete role (no APR/UPL/SHARE alongside it), it is left
 *     holding nothing, same as any other unmapped group.
 *
 * HOW A "LIVE GRANT TO STRIP" IS FOUND
 * -------------------------------------
 * Rather than trying to re-derive which folders a stale Group Map row's fan-out logic *should*
 * have reached (a Head of Department's row sits on the DEPARTMENT term and fans down to every unit
 * beneath it — replicating that offline risks missing folders, or worse, guessing wrong), this
 * script reads the LIVE role assignments directly: it walks every department and unit folder under
 * each affected segment, in every one of the four libraries a delete role could ever land on
 * (Staging, StagingHC, Documents, DocumentsHC — never Archive/ArchiveHC, see above), and for each
 * folder checks whether one of the 393 flagged groups holds a `CRS Delete` binding there. That is
 * ground truth regardless of how the grant got there.
 *
 * ⚠ IT NEVER DESCENDS BELOW UNIT. Reconciliation never grants below Unit either — the below-Unit
 * folders inherit — so there is nothing to find there, and walking that deep would be enormous for
 * no reason.
 *
 * HOW TO RUN
 *   1. F12 → Console on the CRS site, signed in as a site administrator.
 *   2. Paste this whole file, Enter. `DO_WRITE` is `false` — nothing is written on this pass.
 *   3. Read the printed report: the Group Map rows about to be deleted, and the live folder grants
 *      about to be stripped, grouped by segment.
 *   4. If it looks right, set `DO_WRITE = true` at the top and run the WHOLE file again. It will
 *      re-read everything fresh (never trusts anything from the dry run) before writing.
 *   5. OPTIONAL, for a staged rollout on a large finding: set `WRITE_SEGMENTS = ["MHO"]` (or
 *      whichever segment you want to confirm first) alongside `DO_WRITE = true`. Detection and the
 *      report still cover every affected segment either way — this only narrows what actually gets
 *      WRITTEN on that pass, so you can check one segment on the live site before doing the rest.
 *      Leave it as `[]` to write everything the report found in one pass.
 *
 * ⚠ THE ROW DELETIONS HAPPEN BEFORE THE LIVE-GRANT STRIPS, DELIBERATELY. Fixing the source (the
 * Group Map row) first means that even if this run is interrupted partway through stripping live
 * grants, nothing left over can be silently re-granted by a Folder Reconciliation run in the
 * meantime — the row that would have told it to is already gone. Re-running this script afterwards
 * picks up wherever it left off; every step here is safe to repeat (a row already deleted 404s
 * harmlessly on a second attempt, a binding already removed likewise).
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────────────────────── */
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const DO_WRITE = false; // ⚠ set to true ONLY after reviewing the dry-run report below
  /* Leave empty to write EVERY segment the report found. Set e.g. ["MHO"] to write only that
     segment's rows/grants on this pass — a staged rollout, so you can confirm one segment looks
     right on the live site before doing the rest. Detection and the report above always cover
     EVERY affected segment regardless of this — it only narrows what section 9 actually WRITES. */
  const WRITE_SEGMENTS = [];

  /* ── Resolve the site — same three-route fallback as check-stale-delete-roles.js ──────────── */
  const resolveWeb = () => {
    if (SITE_OVERRIDE) return SITE_OVERRIDE.replace(/\/+$/, "");
    try {
      if (typeof _spPageContextInfo !== "undefined" && _spPageContextInfo.webAbsoluteUrl) {
        return _spPageContextInfo.webAbsoluteUrl.replace(/\/+$/, "");
      }
    } catch (e) {
      /* not defined on this page — fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
    return (location.origin + (m ? m[1] : "")).replace(/\/+$/, "");
  };
  const web = resolveWeb();
  console.log("Site:", web, "| DO_WRITE:", DO_WRITE);

  const bust = (u) =>
    (u.indexOf("?") === -1 ? "?" : "&") +
    `_=${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;

  /* ⚠ 429/503 BACKOFF. This script issues far more requests than the read-only check script (a
     folder walk on top of the same Group Map read), so a throttle mid-walk is a real possibility.
     Five attempts, capped exponential backoff — same shape as `withThrottleRetry` elsewhere in
     this project. */
  const withRetry = async (fn) => {
    let attempt = 0;
    for (;;) {
      const res = await fn();
      if (res.status !== 429 && res.status !== 503) return res;
      attempt++;
      if (attempt >= 5) return res;
      const ra = Number(res.headers && res.headers.get && res.headers.get("Retry-After"));
      const waitMs = (isFinite(ra) && ra > 0 ? ra * 1000 : Math.min(1000 * 2 ** attempt, 15000));
      await new Promise((r) => setTimeout(r, waitMs));
    }
  };

  const get = async (url) => {
    const r = await withRetry(() =>
      fetch(web + url + bust(url), {
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

  let digestCache;
  const getDigest = async () => {
    if (digestCache) return digestCache;
    const r = await fetch(web + "/_api/contextinfo", {
      method: "POST",
      headers: { Accept: "application/json;odata=nometadata" },
    });
    if (!r.ok) throw new Error(`Could not get a request digest — HTTP ${r.status}`);
    const j = await r.json();
    digestCache = j.FormDigestValue;
    return digestCache;
  };

  const postAction = async (url) => {
    const digest = await getDigest();
    const r = await withRetry(() =>
      fetch(web + url, {
        method: "POST",
        headers: { Accept: "application/json;odata=nometadata", "X-RequestDigest": digest },
      }),
    );
    return r;
  };

  const deleteListItem = async (listUrl, id) => {
    const digest = await getDigest();
    const r = await withRetry(() =>
      fetch(`${web}${listUrl}/items(${id})`, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          "X-RequestDigest": digest,
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
        },
      }),
    );
    return r;
  };

  const resolveList = async (candidates) => {
    for (const t of candidates) {
      try {
        const probe = await get(
          `/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=Title,ItemCount`,
        );
        if (probe && probe.Title) return probe;
      } catch (e) {
        /* try the next candidate */
      }
    }
    throw new Error(`None of these list titles resolved: ${candidates.join(", ")}`);
  };

  /* Same probe shape as `resolveList`, but also pulls the library's root server-relative URL so
     the folder walk can build paths under it. Returns undefined rather than throwing — an
     unresolved HC library, for instance, is a normal state on a non-HC site. */
  const resolveLibrary = async (candidates) => {
    for (const t of candidates) {
      try {
        const probe = await get(
          `/_api/web/lists/getbytitle('${encodeURIComponent(t)}')` +
            `?$expand=RootFolder&$select=Title,RootFolder/ServerRelativeUrl`,
        );
        if (probe && probe.Title && probe.RootFolder) {
          return { title: probe.Title, rootUrl: probe.RootFolder.ServerRelativeUrl };
        }
      } catch (e) {
        /* try the next candidate */
      }
    }
    return undefined;
  };

  /* ── Role normalisation and the CURRENT persona role sets — copied verbatim from
     check-stale-delete-roles.js. Keep the two files in sync; if `groupMapModel.ts` changes again,
     update BOTH before trusting either script's output. ────────────────────────────────────── */
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
  const STALE_DELETE_ROLES = ["DEL", "DELS", "DELHC", "DELSHC"];

  /* ── The four library candidate lists — copied verbatim from shared/naming.ts. This client
     renames libraries routinely; an unlisted title fails SILENTLY (the library is skipped as
     "not present"), so keep this in sync with naming.ts if it ever gains a new candidate. ────── */
  const LIBRARY_CANDIDATES = ["Approval for Document", "Approval Document", "ApprovalDocument", "Staging"];
  const HC_APPROVAL_CANDIDATES = [
    "Approval for Highly Confidential Document", "HC Approval Document", "HC Approval Documents",
    "HCApprovalDocument", "Highly Confidential Approval Document",
  ];
  const DOCUMENTS_CANDIDATES = [
    "Documents", "Restricted & Confidential Document", "Restricted and Confidential Document",
    "Shared Documents",
  ];
  const HC_DOCUMENTS_CANDIDATES = [
    "HC Documents", "HC Document", "HCDocuments", "Highly Confidential Document",
    "Highly Confidential Documents",
  ];

  /* ── 1. Segments — for labelling only ──────────────────────────────────────────────────── */
  console.log("\n=== 1. Segments (CRS Config) ===");
  const configList = await resolveList(["CRS Config", "DMS Config"]);
  const modeRes = await get(
    `/_api/web/lists/getbytitle('${encodeURIComponent(configList.Title)}')/items` +
      `?$select=Id,ModeLabel,StagingFolder,TermSetGuid,SortOrder,Levels,ConfigType` +
      `&$filter=ConfigType eq 'mode'&$orderby=SortOrder&$top=5000`,
  );
  const allModeRows = modeRes.value || [];
  const codeByGuid = {};
  allModeRows
    .filter((r) => (r.TermSetGuid || "").trim() && (r.Levels || "").trim())
    .forEach((r) => {
      const g = (r.TermSetGuid || "").trim().toLowerCase();
      if (codeByGuid[g] !== undefined) return;
      codeByGuid[g] = (r.StagingFolder || "").trim();
    });
  console.log(`  ${Object.keys(codeByGuid).length} segment code(s) resolved.`);

  /* ── 2. Group Map rows, rolled up per GroupId, AND the specific stale ROW IDS to delete ──── */
  console.log("\n=== 2. Group Map rows (CRS Group Map) ===");
  const gmList = await resolveList(["CRS Group Map", "DMS Group Map"]);
  const gmListUrl = `/_api/web/lists/getbytitle('${encodeURIComponent(gmList.Title)}')`;
  const gmRes = await get(`${gmListUrl}/items?$select=Id,GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`);
  const gmRows = gmRes.value || [];
  console.log(
    `  ${gmRows.length} row(s) read` +
      (gmRows.length >= 5000 ? "  ⚠ AT THE CAP — some rows may be missing, page this manually" : "."),
  );

  const byGroup = {}; // GroupId -> { roles: Set, segments: Set }
  const staleRowsByGroup = {}; // GroupId -> [{ Id, Role }] — ONLY the stale rows, for deletion
  gmRows.forEach((r) => {
    const gid = (r.GroupId ?? "").trim();
    const role = normalizeRoleValue(r.Role ?? "");
    if (!gid || !role) return;
    const entry = byGroup[gid] || (byGroup[gid] = { roles: new Set(), segments: new Set() });
    entry.roles.add(role);
    const seg = (r.Segment ?? "").trim().toLowerCase();
    if (seg) entry.segments.add(seg);
    if (STALE_DELETE_ROLES.indexOf(role) !== -1) {
      (staleRowsByGroup[gid] || (staleRowsByGroup[gid] = [])).push({ Id: r.Id, Role: role });
    }
  });

  /* ── 3. Live SharePoint group titles ───────────────────────────────────────────────────── */
  console.log("\n=== 3. Live group titles (sitegroups) ===");
  const allGroupsRes = await get(`/_api/web/sitegroups?$select=Id,Title&$top=5000`);
  const allGroups = allGroupsRes.value || [];
  const titleById = {};
  allGroups.forEach((g) => {
    titleById[String(g.Id)] = g.Title;
  });
  console.log(`  ${allGroups.length} group(s) on the site.`);

  /* ── 4. Find every group carrying a stale delete role ──────────────────────────────────── */
  console.log("\n=== 4. Groups carrying a stale delete role ===");
  const staleDeleteGroups = [];
  Object.keys(byGroup).forEach((gid) => {
    const entry = byGroup[gid];
    const roles = Array.from(entry.roles);
    const staleFound = roles.filter((r) => STALE_DELETE_ROLES.indexOf(r) !== -1);
    if (staleFound.length === 0) return;
    const title = titleById[gid];
    const segCodes = Array.from(entry.segments)
      .map((g) => codeByGuid[g] || g)
      .filter((v, i, a) => a.indexOf(v) === i);
    staleDeleteGroups.push({
      gid,
      title: title !== undefined ? title : `(id ${gid} — group no longer exists on the site)`,
      gone: title === undefined,
      segCodes,
      staleFound: staleFound.sort(),
    });
  });
  staleDeleteGroups.sort(
    (a, b) => (a.segCodes[0] || "").localeCompare(b.segCodes[0] || "") || a.title.localeCompare(b.title),
  );

  if (staleDeleteGroups.length === 0) {
    console.log("  ✅ NOTHING TO DO. No group on this site carries a stale delete role.");
    return;
  }
  console.log(`  ${staleDeleteGroups.length} group(s) found.`);

  /* ── 5. Resolve the "CRS Delete" (or "DMS Delete") role definition ────────────────────────
     Every one of DEL/DELS/DELHC/DELSHC resolves to the SAME permission level — the separation is
     which LIBRARY the grant lands on (`LIBRARY_ROLES`), never the level name. So there is exactly
     ONE role definition to find and strip, wherever it turns up on a flagged group. */
  console.log("\n=== 5. Resolving the delete permission level ===");
  const roleDefsRes = await get(`/_api/web/roledefinitions?$select=Name,Id&$top=5000`);
  const roleDefs = roleDefsRes.value || [];
  const deleteRoleDef =
    roleDefs.find((d) => d.Name === "CRS Delete") || roleDefs.find((d) => d.Name === "DMS Delete");
  if (!deleteRoleDef) {
    console.log(
      "  ✗ Could not find a 'CRS Delete' or 'DMS Delete' permission level on this site. " +
        "Nothing can be safely stripped without knowing which level to target — aborting.",
    );
    return;
  }
  console.log(`  Found "${deleteRoleDef.Name}" (id ${deleteRoleDef.Id}).`);

  /* ── 6. Resolve the four libraries a delete role could ever land on ──────────────────────── */
  console.log("\n=== 6. Resolving libraries ===");
  const libraries = [
    { key: "Staging", candidates: LIBRARY_CANDIDATES },
    { key: "StagingHC", candidates: HC_APPROVAL_CANDIDATES },
    { key: "Documents", candidates: DOCUMENTS_CANDIDATES },
    { key: "DocumentsHC", candidates: HC_DOCUMENTS_CANDIDATES },
  ];
  for (const lib of libraries) {
    lib.resolved = await resolveLibrary(lib.candidates);
    console.log(
      lib.resolved
        ? `  ${lib.key}: "${lib.resolved.title}"`
        : `  ${lib.key}: not present on this site (normal if this site has no HC pair).`,
    );
  }
  const resolvedLibraries = libraries.filter((l) => l.resolved);

  /* ── 7. Walk each affected segment's department/unit folders, in every resolved library,
     looking for a flagged group holding "CRS Delete" ──────────────────────────────────────── */
  console.log("\n=== 7. Walking live folder permissions ===");
  const staleGroupIds = new Set(staleDeleteGroups.filter((g) => !g.gone).map((g) => g.gid));
  const affectedSegments = Array.from(
    new Set(staleDeleteGroups.flatMap((g) => g.segCodes)),
  ).filter((s) => s.length > 0);
  console.log(`  Segments to walk: ${affectedSegments.join(", ") || "(none — every finding is a GLOBAL/site-entry row, nothing to walk)"}`);

  const roleAssignmentsForFolder = async (path) => {
    const j = await get(
      `/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments` +
        `?$expand=Member,RoleDefinitionBindings` +
        `&$select=PrincipalId,Member/Title,RoleDefinitionBindings/Name,RoleDefinitionBindings/Id` +
        `&@f='${encodeURIComponent(path)}'`,
    );
    return j.__notFound ? undefined : j.value || [];
  };
  const childFolders = async (path) => {
    const j = await get(
      `/_api/web/GetFolderByServerRelativeUrl(@f)/Folders?$select=Name,ServerRelativeUrl&@f='${encodeURIComponent(path)}'`,
    );
    return j.__notFound ? undefined : j.value || [];
  };

  const liveGrantFindings = []; // { segCode, libKey, path, gid, groupTitle }
  const MAX_FOLDER_WALK_REQUESTS = 6000;
  let requestBudget = MAX_FOLDER_WALK_REQUESTS;
  const spend = (n) => {
    requestBudget -= n;
    return requestBudget > 0;
  };

  /* ⚠ PROGRESS HEARTBEAT. This walk is several hundred sequential live REST calls with NOTHING
     printed in between otherwise — found live 2026-09-22, where a dry run on GHO+MHO sat silent
     for minutes and read as stuck even though it was working the whole time. A count every 25
     folders is cheap and is the difference between "is this still running?" and just watching it
     go. */
  let foldersChecked = 0;
  const heartbeat = (segCode, libKey) => {
    foldersChecked++;
    if (foldersChecked % 25 === 0) {
      console.log(`  … ${foldersChecked} folder(s) checked so far (currently ${segCode} / ${libKey}, ${liveGrantFindings.length} finding(s) so far, ${requestBudget} request(s) left in budget)`);
    }
  };

  const checkFolder = async (segCode, libKey, path) => {
    const assignments = await roleAssignmentsForFolder(path);
    if (!spend(1)) return;
    heartbeat(segCode, libKey);
    if (!assignments) return; // no ACL to read on this exact object — nothing to find
    for (const ra of assignments) {
      const gid = String(ra.PrincipalId ?? "");
      if (!staleGroupIds.has(gid)) continue;
      const bindings = ra.RoleDefinitionBindings || [];
      const hasDelete = bindings.some((b) => b.Id === deleteRoleDef.Id || b.Name === deleteRoleDef.Name);
      if (!hasDelete) continue;
      liveGrantFindings.push({
        segCode,
        libKey,
        path,
        gid,
        groupTitle: (ra.Member && ra.Member.Title) || titleById[gid] || `id ${gid}`,
      });
    }
  };

  for (const segCode of affectedSegments) {
    for (const lib of resolvedLibraries) {
      if (!spend(1)) break;
      const segRoot = `${lib.resolved.rootUrl}/${segCode}`;
      const depts = await childFolders(segRoot);
      if (!spend(1)) break;
      if (depts === undefined) {
        console.log(`  (${segCode} / ${lib.key}) segment folder not found — skipping.`);
        continue;
      }
      for (const dept of depts) {
        if (requestBudget <= 0) break;
        await checkFolder(segCode, lib.key, dept.ServerRelativeUrl);
        const units = await childFolders(dept.ServerRelativeUrl);
        if (!spend(1)) break;
        if (units === undefined) continue;
        for (const unit of units) {
          if (requestBudget <= 0) break;
          await checkFolder(segCode, lib.key, unit.ServerRelativeUrl);
        }
      }
    }
  }
  if (requestBudget <= 0) {
    console.log(
      `  ⚠ HIT THE ${MAX_FOLDER_WALK_REQUESTS}-REQUEST WALK BUDGET — some folders were not checked. ` +
        `Re-run scoped to fewer segments, or raise MAX_FOLDER_WALK_REQUESTS, before trusting this as complete.`,
    );
  }
  console.log(`  ${liveGrantFindings.length} live "CRS Delete" grant(s) found to strip.`);

  /* ── 8. REPORT ──────────────────────────────────────────────────────────────────────────── */
  console.log("\n=== REPORT ===");
  const rowDeletions = [];
  staleDeleteGroups.forEach((g) => {
    (staleRowsByGroup[g.gid] || []).forEach((row) => {
      rowDeletions.push({
        gid: g.gid,
        title: g.title,
        segCodes: g.segCodes.join(", ") || "(none)",
        itemId: row.Id,
        role: row.Role,
      });
    });
  });
  /* ⚠ FULL DETAIL IS ON `window`, NEVER ONLY FORCE-PRINTED. Found live 2026-09-22: past a few
     hundred rows, Chrome's console collapses `console.table` into an inert "Array(N)" reference
     instead of an interactive table, so at this site's real scale (1000+ rows) the two calls below
     are not actually reviewable as printed. Same pattern this project's own
     `diagnose-site-members-leak.js` already uses (`window.__leakReport`) for exactly this reason —
     keep the small SUMMARY on screen, keep the full list on window for filtering/exporting. */
  window.__staleDeleteRepair = { rowDeletions, liveGrantFindings };
  console.log(
    "\n⚠ Full detail is on window.__staleDeleteRepair — { rowDeletions, liveGrantFindings }. " +
      "Past a few hundred rows, console.table often renders as a collapsed 'Array(N)' rather than " +
      "an interactive table; inspect it instead with, e.g.:\n" +
      "    console.table(window.__staleDeleteRepair.liveGrantFindings.filter(f => f.segCode === 'GHO'))\n" +
      "    copy(JSON.stringify(window.__staleDeleteRepair, null, 2))  // clipboard, for pasting elsewhere\n" +
      "    window.__staleDeleteRepair.liveGrantFindings.filter(f => f.groupTitle.includes('<a group you recognise>'))",
  );

  const countBy = (rows, keyFn) => {
    const counts = {};
    rows.forEach((r) => {
      const k = keyFn(r);
      counts[k] = (counts[k] || 0) + 1;
    });
    return Object.keys(counts)
      .sort()
      .map((k) => ({ key: k, count: counts[k] }));
  };

  console.log(`\n-- Group Map rows to delete (${rowDeletions.length}) — by segment + role --`);
  console.table(countBy(rowDeletions, (r) => `${r.segCodes} / ${r.role}`).map((c) => ({ "Segment / Role": c.key, Count: c.count })));

  console.log(`\n-- Live "CRS Delete" grants to strip (${liveGrantFindings.length}) — by segment + library --`);
  console.table(countBy(liveGrantFindings, (f) => `${f.segCode} / ${f.libKey}`).map((c) => ({ "Segment / Library": c.key, Count: c.count })));

  console.log(`\n-- Live "CRS Delete" grants to strip — by GROUP (top 30 by count) --`);
  console.table(
    countBy(liveGrantFindings, (f) => f.groupTitle)
      .sort((a, b) => b.count - a.count)
      .slice(0, 30)
      .map((c) => ({ Group: c.key, "Folders affected": c.count })),
  );

  const goneGroupsWithRows = staleDeleteGroups.filter((g) => g.gone && (staleRowsByGroup[g.gid] || []).length > 0);
  if (goneGroupsWithRows.length > 0) {
    console.log(
      `\n⚠ ${goneGroupsWithRows.length} of the flagged groups no longer exist on the site — their ` +
        `stale Group Map rows will still be deleted (pure cleanup; nothing to strip live, since the ` +
        `principal is already gone).`,
    );
  }

  console.log(`\n${rowDeletions.length} row deletion(s) + ${liveGrantFindings.length} live grant removal(s) planned.`);

  /* ── 9. WRITE ───────────────────────────────────────────────────────────────────────────── */
  if (!DO_WRITE) {
    console.log(
      "\nDRY RUN — nothing was changed. Review the two tables above. If they look right, set " +
        "DO_WRITE = true at the top of this script and run the whole file again to apply them.",
    );
    return;
  }

  const writeSegFilter = new Set(WRITE_SEGMENTS.map((s) => (s ?? "").trim()).filter(Boolean));
  const rowsToWrite =
    writeSegFilter.size === 0
      ? rowDeletions
      : rowDeletions.filter((r) => r.segCodes.split(", ").some((s) => writeSegFilter.has(s)));
  const grantsToWrite =
    writeSegFilter.size === 0
      ? liveGrantFindings
      : liveGrantFindings.filter((f) => writeSegFilter.has(f.segCode));
  if (writeSegFilter.size > 0) {
    console.log(
      `\n⚠ WRITE_SEGMENTS is set to [${Array.from(writeSegFilter).join(", ")}] — writing ONLY those ` +
        `segments this pass: ${rowsToWrite.length} of ${rowDeletions.length} row deletion(s), ` +
        `${grantsToWrite.length} of ${liveGrantFindings.length} live grant removal(s). Everything else ` +
        `stays untouched until you re-run with a different (or empty) WRITE_SEGMENTS.`,
    );
  }

  console.log("\n=== 9a. Deleting stale Group Map rows ===");
  let rowsOk = 0, rowsFailed = 0;
  for (const r of rowsToWrite) {
    try {
      const res = await deleteListItem(gmListUrl, r.itemId);
      if (res.ok || res.status === 204 || res.status === 404) {
        rowsOk++;
      } else {
        rowsFailed++;
        console.log(`  ✗ row ${r.itemId} (${r.title}, ${r.role}) — HTTP ${res.status}`);
      }
    } catch (e) {
      rowsFailed++;
      console.log(`  ✗ row ${r.itemId} (${r.title}, ${r.role}) — ${String(e)}`);
    }
  }
  console.log(`  ${rowsOk} deleted, ${rowsFailed} failed.`);

  console.log("\n=== 9b. Stripping live \"CRS Delete\" grants ===");
  let grantsOk = 0, grantsFailed = 0;
  for (const f of grantsToWrite) {
    try {
      const res = await postAction(
        `/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments/` +
          `removeroleassignment(principalid=${f.gid},roledefid=${deleteRoleDef.Id})` +
          `?@f='${encodeURIComponent(f.path)}'`,
      );
      if (res.ok || res.status === 204) {
        grantsOk++;
        console.log(`  ✓ ${f.groupTitle} — ${f.path}`);
      } else {
        grantsFailed++;
        console.log(`  ✗ ${f.groupTitle} — ${f.path} — HTTP ${res.status}`);
      }
    } catch (e) {
      grantsFailed++;
      console.log(`  ✗ ${f.groupTitle} — ${f.path} — ${String(e)}`);
    }
  }
  console.log(`  ${grantsOk} stripped, ${grantsFailed} failed.`);

  console.log(
    `\n=== Done: ${rowsOk}/${rowsToWrite.length} rows deleted, ${grantsOk}/${grantsToWrite.length} ` +
      `live grants stripped. ===\n` +
      (rowsFailed || grantsFailed
        ? "⚠ Some items failed — re-run this whole script (dry run first) to see what is left; " +
          "every step here is safe to repeat."
        : "Re-run check-stale-delete-roles.js to confirm the tenant now reads clean, then run Folder " +
          "Reconciliation for the affected segments as a final check."),
  );
})();

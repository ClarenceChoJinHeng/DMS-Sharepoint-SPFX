/*
 * Independent, read-only scan: which groups on this site still carry a role the current persona
 * model has NOT granted since 2026-09-17 (`DEL` / `DELS` / `DELHC` / `DELSHC`)?
 * -----------------------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, renamed or deleted.
 *
 * Run it in the browser console, ON THE CRS SITE. Any signed-in account with Read on the CRS Group
 * Map, CRS Config and the site groups is enough — it does not need to be an admin.
 *
 * WHY THIS EXISTS (found live, 2026-09-21, on `NBPOLHO_CDS_UPSUPPORT_APPROVER` and its HC twin):
 * the 2026-09-17 proxy-deletion redesign removed DEL/DELS/DELHC/DELSHC from every persona in
 * `PERSONAS` (see `shared/groupMapModel.ts`) — as of that change, NO persona holds any of the four
 * roles below. But that edit is CODE-ONLY. A group created before it still has its OLD Group Map
 * rows, still literally saying `DELS`/`DEL` (or the HC equivalents), and Folder Reconciliation just
 * applies whatever is in those rows with no awareness the persona definition ever changed — so it
 * keeps re-granting `CRS Delete` on that folder, every run, silently. CLAUDE.md's own words on the
 * one group confirmed live: *"every pre-2026-09-17 Head of Unit / approver group across the ENTIRE
 * tenant can still delete files and folders directly, completely bypassing the proxy-deletion
 * mechanism that redesign exists to enforce."* This script exists to learn how many, before working
 * through them one at a time.
 *
 * THE FIX FOR EACH ONE FOUND IS NOT "untick the permission by hand" — that gets silently UNDONE by
 * the next Folder Reconciliation run, because the underlying Group Map row still says the stale
 * role and reconciliation only ever GRANTS, never revokes, at folder scope. The only durable fix:
 * note the group's members (export first — deleting the group loses them), DELETE the group,
 * re-create it via Group Management's persona picker (which writes the CURRENT role set with no
 * delete role at all), re-add the members, re-run Folder Reconciliation. SharePoint drops a deleted
 * principal's role assignments; a removed Group Map row alone does not.
 *
 * WHAT THIS SCRIPT DOES:
 *   1. Reads every "mode" row in CRS Config, to label results by segment code.
 *   2. Reads every row in CRS Group Map, rolls them up per GroupId into one role set each — matching
 *      exactly how the app itself decides what a group grants (one row = one role, deduped, upper-
 *      cased, aliases normalised the same way `normalizeRoleValue` does).
 *   3. Reads every group's LIVE SharePoint title (never trusts the stored `GroupName` column, which
 *      goes stale on a rename — the 1.0.162.0 lesson, joined here on `GroupId` only).
 *   4. For every group whose role set includes DEL, DELS, DELHC or DELSHC, reports it by name,
 *      segment and its exact role set — this is the tenant-wide scope of the live risk.
 *   5. As a secondary, lower-priority finding: any OTHER group whose role set matches no current
 *      persona at all (the same "these roles match no persona exactly" check Group Management's own
 *      Quick Search already makes) — usually harmless (a hand-mapped one-off, an extra `ENTRY` row),
 *      listed separately so the stale-delete finding is never buried under it.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  /* Leave blank to auto-detect. Set it if the auto-detected site below is wrong, e.g.
     "https://sdguthrie.sharepoint.com/sites/CRS" */
  const SITE_OVERRIDE = "";

  /* ── Resolve the site ──────────────────────────────────────────────────────
     Same three-route fallback the other scripts in this project use. `_spPageContextInfo` is
     missing on some modern pages, and parsing the page path breaks the moment the script is run
     from a library or a settings page — which is how an earlier script in this repo ended up
     fetching the page's own HTML and failing with "Unexpected token '<'". */
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

  /* ⚠ CACHE-BUSTED. This project has repeatedly been bitten by a plain `fetch` GET being answered
     from the browser's own HTTP cache — `Cache-Control`/`Pragma` alone are not always enough to
     force revalidation, and a unique query parameter is the one thing a cache cannot match against a
     prior response. Same `bust()` pattern several screens in this app already use for this reason. */
  const bust = (u) =>
    (u.indexOf("?") === -1 ? "?" : "&") +
    `_=${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  const get = async (url) => {
    const r = await fetch(url + bust(url), {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
      },
      cache: "no-store",
    });
    const type = r.headers.get("content-type") || "";
    if (type.indexOf("json") === -1) {
      /* A 200 carrying HTML means the ADDRESS is wrong, not the query — set SITE_OVERRIDE. */
      throw new Error(
        `Expected JSON and got "${type}" from ${url}\n` +
          `The site resolved to ${web}. If that is wrong, set SITE_OVERRIDE at the top and re-run.`,
      );
    }
    if (!r.ok) throw new Error(`HTTP ${r.status} from ${url}`);
    return r.json();
  };

  const resolveList = async (candidates) => {
    for (const t of candidates) {
      try {
        const probe = await get(
          `${web}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=Title,ItemCount`,
        );
        if (probe && probe.Title) return probe;
      } catch (e) {
        /* try the next candidate */
      }
    }
    throw new Error(`None of these list titles resolved: ${candidates.join(", ")}`);
  };

  /* ── Role normalisation — the same rules as `normalizeRoleValue` in groupMapModel.ts ──────────
     Short codes pass through unchanged; a handful of legacy long-form spellings alias to one of
     them; anything else is left as-is (so it visibly fails to match rather than silently becoming
     some role nobody asked for). */
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

  /* ── The CURRENT persona role sets — a snapshot of `PERSONAS` in groupMapModel.ts as of the
     2026-09-17 proxy-deletion redesign. NOT ONE OF THEM HOLDS DEL, DELS, DELHC OR DELSHC — that is
     the whole fact this script exists to check against live data. If `groupMapModel.ts` changes
     again, re-copy this list from there before trusting a "no persona matches" verdict below. */
  const PERSONAS = [
    { key: "clevel_global", roles: ["GLOBAL"] },
    { key: "clevel_segment", roles: ["SEGVIEW"] },
    { key: "hod", roles: ["DEPTVIEW", "SHARE", "SHAREHC"] },
    { key: "hou", roles: ["APR", "SHARE", "UPL"] },
    { key: "hou_hc", roles: ["APRHC", "SHARE", "UPLHC", "SHAREHC"] },
    { key: "pic", roles: ["UPL"] },
    { key: "pic_hc", roles: ["UPLHC"] },
    { key: "employee_hc", roles: ["MEMBER", "MEMBERHC"] },
    { key: "employee", roles: ["MEMBER"] },
  ];
  const roleSetKey = (roles) => {
    const seen = [];
    for (const r of roles ?? []) {
      const v = (r ?? "").trim().toUpperCase();
      if (v && seen.indexOf(v) === -1) seen.push(v);
    }
    return seen.sort().join("+");
  };
  const PERSONA_KEYS = {};
  PERSONAS.forEach((p) => {
    PERSONA_KEYS[roleSetKey(p.roles)] = p.key;
  });
  const personaKeyForRoles = (roles) => PERSONA_KEYS[roleSetKey(roles)]; // undefined = no match

  const STALE_DELETE_ROLES = ["DEL", "DELS", "DELHC", "DELSHC"];

  /* ── 1. Segments — code from every "mode" row, for labelling only ─────────────────── */
  console.log("\n=== 1. Segments (CRS Config) ===");
  const configList = await resolveList(["CRS Config", "DMS Config"]);
  console.log(`  List: "${configList.Title}" (${configList.ItemCount} items)`);
  const modeRes = await get(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(configList.Title)}')/items` +
      `?$select=Id,ModeLabel,StagingFolder,TermSetGuid,SortOrder,Levels,ConfigType` +
      `&$filter=ConfigType eq 'mode'&$orderby=SortOrder&$top=5000`,
  );
  const allModeRows = modeRes.value || [];
  // First match per guid wins, matching the app's own `loadModes()` (lowest SortOrder, Levels
  // non-blank) — a duplicate mode row must not silently relabel results under the wrong code.
  const codeByGuid = {};
  allModeRows
    .filter((r) => (r.TermSetGuid || "").trim() && (r.Levels || "").trim())
    .forEach((r) => {
      const g = (r.TermSetGuid || "").trim().toLowerCase();
      if (codeByGuid[g] !== undefined) return;
      codeByGuid[g] = (r.StagingFolder || "").trim();
    });
  console.log(`  ${Object.keys(codeByGuid).length} segment code(s) resolved.`);

  /* ── 2. Group Map rows, rolled up per GroupId ──────────────────────────────────────── */
  console.log("\n=== 2. Group Map rows (CRS Group Map) ===");
  const gmList = await resolveList(["CRS Group Map", "DMS Group Map"]);
  console.log(`  List: "${gmList.Title}" (${gmList.ItemCount} items)`);
  /* ⚠ CAPPED AT $top=5000, matching how the app itself reads this list. A truncated read here would
     UNDER-report affected groups (missed rows, not invented ones) — the safe direction for a
     confirmation script, but if `gmList.ItemCount` is anywhere near 5000, treat a clean result with
     suspicion and page it by hand. */
  const gmRes = await get(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(gmList.Title)}')/items` +
      `?$select=GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
  );
  const gmRows = gmRes.value || [];
  console.log(
    `  ${gmRows.length} row(s) read` +
      (gmRows.length >= 5000
        ? "  ⚠ AT THE CAP — some rows may be missing, page this manually"
        : "."),
  );

  const byGroup = {}; // GroupId -> { roles: Set, storedName, segments: Set }
  gmRows.forEach((r) => {
    const gid = (r.GroupId ?? "").trim();
    const role = normalizeRoleValue(r.Role ?? "");
    if (!gid || !role) return;
    const entry =
      byGroup[gid] ||
      (byGroup[gid] = {
        roles: new Set(),
        storedName: (r.GroupName ?? "").trim(),
        segments: new Set(),
      });
    entry.roles.add(role);
    const seg = (r.Segment ?? "").trim().toLowerCase();
    if (seg) entry.segments.add(seg);
  });
  console.log(`  ${Object.keys(byGroup).length} distinct group(s) hold at least one mapping row.`);

  /* ── 3. Live SharePoint group titles — never trust the stored GroupName column, which a rename
     leaves stale (the 1.0.162.0 lesson: only `GroupId` survives a rename). ─────────────────────── */
  console.log("\n=== 3. Live group titles (sitegroups) ===");
  const allGroupsRes = await get(`${web}/_api/web/sitegroups?$select=Id,Title&$top=5000`);
  const allGroups = allGroupsRes.value || [];
  console.log(
    `  ${allGroups.length} group(s) on the site` +
      (allGroups.length >= 5000 ? "  ⚠ AT THE CAP — page this manually" : "."),
  );
  const titleById = {};
  allGroups.forEach((g) => {
    titleById[String(g.Id)] = g.Title;
  });

  /* ── 4. The priority finding: any group carrying a STALE DELETE role ──────────────────────── */
  console.log("\n=== 4. Groups carrying a role no current persona holds — DEL / DELS / DELHC / DELSHC ===");
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
      roles: roles.sort(),
      staleFound: staleFound.sort(),
    });
  });
  staleDeleteGroups.sort((a, b) =>
    (a.segCodes[0] || "").localeCompare(b.segCodes[0] || "") ||
    a.title.localeCompare(b.title),
  );

  if (staleDeleteGroups.length === 0) {
    console.log("  ✅ NONE FOUND. No group on this site carries DEL, DELS, DELHC or DELSHC.");
  } else {
    console.log(
      `  ⚠⚠ ${staleDeleteGroups.length} GROUP(S) STILL CARRY A ROLE NO PERSONA HAS GRANTED SINCE ` +
        `2026-09-17. Every one of these can delete files AND FOLDERS directly, bypassing the proxy-` +
        `deletion mechanism entirely:\n`,
    );
    staleDeleteGroups.forEach((g) => {
      console.log(
        `    ⚠ [id ${g.gid}] "${g.title}"` +
          `${g.segCodes.length ? `  (segment: ${g.segCodes.join(", ")})` : "  (no segment — GLOBAL/site-entry row)"}`,
      );
      console.log(`        full role set: ${g.roles.join(", ")}`);
      console.log(`        stale delete role(s): ${g.staleFound.join(", ")}`);
    });
  }

  /* ── 5. Secondary, lower-priority finding: matches no persona at all ───────────────────────── */
  console.log(
    "\n=== 5. Other groups whose role set matches no current persona exactly (usually harmless — hand-mapped, or an extra role) ===",
  );
  const otherMismatches = [];
  Object.keys(byGroup).forEach((gid) => {
    const entry = byGroup[gid];
    const roles = Array.from(entry.roles);
    const hasStaleDelete = roles.some((r) => STALE_DELETE_ROLES.indexOf(r) !== -1);
    if (hasStaleDelete) return; // already reported above — do not report it twice
    const match = personaKeyForRoles(roles);
    if (match !== undefined) return;
    const title = titleById[gid];
    const segCodes = Array.from(entry.segments)
      .map((g) => codeByGuid[g] || g)
      .filter((v, i, a) => a.indexOf(v) === i);
    otherMismatches.push({
      gid,
      title: title !== undefined ? title : `(id ${gid} — group no longer exists on the site)`,
      segCodes,
      roles: roles.sort(),
    });
  });
  otherMismatches.sort((a, b) => a.title.localeCompare(b.title));
  if (otherMismatches.length === 0) {
    console.log("  (none — every other mapped group matches a current persona exactly)");
  } else {
    otherMismatches.forEach((g) => {
      console.log(
        `    [id ${g.gid}] "${g.title}"${g.segCodes.length ? `  (segment: ${g.segCodes.join(", ")})` : ""} — roles: ${g.roles.join(", ")}`,
      );
    });
  }

  /* ── 6. Summary ─────────────────────────────────────────────────────────────────────────────── */
  console.log("\n=== Summary ===");
  const affectedSegments = Array.from(
    new Set(staleDeleteGroups.flatMap((g) => (g.segCodes.length ? g.segCodes : ["(none)"]))),
  ).sort();
  console.log(`  ${Object.keys(byGroup).length} mapped group(s) checked.`);
  console.log(
    `  ${staleDeleteGroups.length} group(s) carry a stale delete role` +
      (staleDeleteGroups.length
        ? `, across ${affectedSegments.length} segment(s): ${affectedSegments.join(", ")}.`
        : "."),
  );
  console.log(`  ${otherMismatches.length} other group(s) match no persona (lower priority).`);
  console.log(
    staleDeleteGroups.length === 0
      ? "\n✅ THE TENANT IS CLEAN — no group can bypass the proxy-deletion mechanism."
      : `\n⏭ NEXT: for each group named in section 4 above — export its members (Group Management),` +
          ` DELETE the group, re-create it via the persona picker, re-add the members, then re-run` +
          ` Folder Reconciliation for the affected segment(s).`,
  );
})();

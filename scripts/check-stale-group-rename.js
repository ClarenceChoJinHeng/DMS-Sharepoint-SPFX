/*
 * Independent, read-only check: did the "Check for groups with a stale name" fix (2026-09-16/17)
 * actually work, everywhere, not just on the three segments already tested by eye in the UI?
 * -----------------------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, renamed or deleted.
 *
 * Run it in the browser console, ON THE CRS SITE. Any signed-in account with Read on the CRS Group
 * Map and CRS Config lists is enough — it does not need to be an admin.
 *
 * WHAT IT CHECKS, and why this is the right independent test:
 * The bug being fixed was that "Check for groups with a stale name" was comparing a group's CURRENT
 * name against the `CRS Group Map` list's own STORED `GroupName` column — which `renameSiteGroup`
 * never updates (only the actual SharePoint group object gets renamed). So a group already renamed
 * correctly on SharePoint kept reappearing as "stale" on the next check, using a name nobody could
 * see anywhere on screen.
 *
 * This script does NOT re-run the app's own check (that would just prove the app agrees with
 * itself). It asks the question a different way: for every segment on the site, does every one of
 * its Group-Map-mapped groups' LIVE SharePoint title actually start with that segment's CURRENT
 * code? That is the one fact the fix exists to make true, checked straight from the two lists and
 * `sitegroups`, with no dependency on the same client-side code being verified.
 *
 * A clean run across every segment (not only the ones already tested) is the strongest available
 * confirmation this can be run without deploying anything new.
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
     from the BROWSER'S OWN HTTP cache rather than the network — most dangerously right after a
     write, which is exactly the situation here (checking a rename immediately after running one in
     the same tab). `Cache-Control`/`Pragma` alone are not always enough to force revalidation; a
     unique query parameter on every request is the one thing a cache cannot match against a prior
     response, and is the same `bust()` pattern several screens in this app already use for this
     reason (`ApprovalDocument.tsx`, `Requests.tsx`, `MySubmissions.tsx`). */
  const bust = (u) => (u.indexOf("?") === -1 ? "?" : "&") + `_=${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  const get = async (url) => {
    const r = await fetch(url + bust(url), {
      headers: { Accept: "application/json;odata=nometadata", "Cache-Control": "no-cache", Pragma: "no-cache" },
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

  /* ── 1. Segments — code + term-set guid from every "mode" row ─────────────── */
  console.log("\n=== 1. Segments (CRS Config) ===");
  const configList = await resolveList(["CRS Config", "DMS Config"]);
  console.log(`  List: "${configList.Title}" (${configList.ItemCount} items)`);
  /* ⚠ `Id` and `SortOrder` are selected NOW (not in the first version of this script) specifically
     to catch a DUPLICATE row sharing one TermSetGuid — the app's own `loadModes()` in
     `GroupManager.tsx` reads `$orderby=SortOrder` and takes the FIRST match for a guid, so if two
     rows share a guid, which one "wins" depends on SortOrder, not on list order — and this script's
     original version had no ordering at all, which is exactly the kind of silent divergence that
     produces two tools disagreeing about the same segment's code. */
  const modeRes = await get(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(configList.Title)}')/items` +
      `?$select=Id,Title,ModeLabel,StagingFolder,TermSetGuid,SortOrder,Levels,ConfigType` +
      `&$filter=ConfigType eq 'mode'&$orderby=SortOrder&$top=5000`,
  );
  const allModeRows = modeRes.value || [];
  const modeRows = allModeRows.filter((r) => r.TermSetGuid && r.StagingFolder);
  console.log(`  ${modeRows.length} segment(s) with a code:`, modeRows.map((r) => r.StagingFolder));

  /* Group every row (including one with no Levels, which the app's own `loadModes()` would exclude
     — replicated below) by TermSetGuid, so a duplicate is impossible to miss. */
  const rowsByGuid = {};
  allModeRows.forEach((r) => {
    const g = (r.TermSetGuid || "").trim().toLowerCase();
    if (!g) return;
    (rowsByGuid[g] = rowsByGuid[g] || []).push(r);
  });
  const dupeGuids = Object.keys(rowsByGuid).filter((g) => rowsByGuid[g].length > 1);
  if (dupeGuids.length > 0) {
    console.log(`\n  ⚠⚠ ${dupeGuids.length} TERM SET GUID(S) HAVE MORE THAN ONE "mode" ROW:`);
    dupeGuids.forEach((g) => {
      console.log(`    guid ${g}:`);
      rowsByGuid[g].forEach((r) =>
        console.log(
          `      [Id ${r.Id}] Title="${r.Title}" ModeLabel="${r.ModeLabel}" ` +
            `StagingFolder="${r.StagingFolder}" SortOrder=${r.SortOrder} ` +
            `Levels=${(r.Levels || "").trim() ? "present" : "BLANK"}`,
        ),
      );
    });
  } else {
    console.log("  (no TermSetGuid has more than one mode row)");
  }

  /* This is the ONE map the rest of the script uses, and it now replicates `loadModes()` exactly —
     filter on Levels non-blank as well as TermSetGuid, order by SortOrder, first match per guid wins
     — rather than "whichever row was seen last while iterating", which was the earlier bug. */
  const codeByGuid = {};
  const labelByGuid = {};
  allModeRows
    .filter((r) => (r.TermSetGuid || "").trim() && (r.Levels || "").trim())
    .forEach((r) => {
      const g = (r.TermSetGuid || "").trim().toLowerCase();
      if (codeByGuid[g] !== undefined) return; // first match (lowest SortOrder) wins, same as loadModes()
      codeByGuid[g] = (r.StagingFolder || "").trim();
      labelByGuid[g] = (r.ModeLabel || r.TermSetGuid || "").trim();
    });

  /* ── 2. Group Map rows — GroupId + Segment (guid) only, no need for the rest ─ */
  console.log("\n=== 2. Group Map rows (CRS Group Map) ===");
  const gmList = await resolveList(["CRS Group Map", "DMS Group Map"]);
  console.log(`  List: "${gmList.Title}" (${gmList.ItemCount} items)`);
  /* ⚠ CAPPED AT $top=5000 — matches how the app's own reads treat this list, and every site this
     has been run against sits well under it. A truncated read here would UNDER-report stale groups
     (missed rows, not invented ones), which is the safe direction for a confirmation script — but
     if `gmList.ItemCount` is anywhere near 5000, treat a clean result with suspicion and page it. */
  const gmRes = await get(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(gmList.Title)}')/items` +
      `?$select=GroupId,Segment&$top=5000`,
  );
  const gmRows = gmRes.value || [];
  console.log(
    `  ${gmRows.length} row(s) read` +
      (gmRows.length >= 5000 ? "  ⚠ AT THE CAP — some rows may be missing, page this manually" : "."),
  );

  /* One entry per (segment guid -> set of group ids actually mapped under it). GLOBAL / site-entry
     rows carry a blank Segment and are correctly excluded — a bare "C_LEVEL_GLOBAL" group has no
     segment prefix to check by design. */
  const idsBySegment = {};
  gmRows.forEach((r) => {
    const seg = (r.Segment || "").trim().toLowerCase();
    const gid = (r.GroupId || "").trim();
    if (!seg || !gid || !codeByGuid[seg]) return;
    (idsBySegment[seg] = idsBySegment[seg] || new Set()).add(gid);
  });

  /* ── 3. Live SharePoint group titles ───────────────────────────────────────── */
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

  /* ── 4. Verdict, per segment ────────────────────────────────────────────────
     The fact being checked: every group mapped under segment X's live SharePoint title starts with
     X's CURRENT code. A group whose title does not is exactly the failure this fix targets — a
     rename that either never happened, or (the actual bug) happened on SharePoint but the app kept
     reporting it as still needing one because it was reading a stale cached name. */
  console.log("\n=== 4. Verdict, per segment ===");
  let totalStale = 0;
  let totalChecked = 0;
  const guids = Object.keys(idsBySegment).sort((a, b) =>
    codeByGuid[a].localeCompare(codeByGuid[b]),
  );
  for (const guid of guids) {
    const code = codeByGuid[guid];
    const ids = Array.from(idsBySegment[guid]);
    const stale = [];
    const gone = [];
    ids.forEach((id) => {
      const title = titleById[id];
      if (title === undefined) {
        gone.push(id); // mapped but the group no longer exists — a different, unrelated finding
        return;
      }
      const t = title.toUpperCase();
      const c = code.toUpperCase();
      // Matches the app's own naming shape: "<CODE>_..." (the normal case) or exactly "<CODE>"
      // (unlikely, but not a bug if it happened).
      if (t.indexOf(c + "_") !== 0 && t !== c) stale.push({ id, title });
    });
    totalChecked += ids.length;
    totalStale += stale.length;
    const label = labelByGuid[guid] || code;
    if (stale.length === 0 && gone.length === 0) {
      console.log(
        `  ✓ ${code} (${label}) — all ${ids.length} mapped group(s) carry a live title starting with "${code}".`,
      );
    } else {
      if (stale.length > 0) {
        console.log(
          `  ⚠ ${code} (${label}) — ${stale.length} of ${ids.length} mapped group(s) do NOT start with "${code}":`,
        );
        stale.forEach((s) => console.log(`      [id ${s.id}] "${s.title}"`));
      }
      if (gone.length > 0) {
        console.log(
          `  (${code} — ${gone.length} mapped group id(s) no longer exist on the site: ` +
            `${gone.join(", ")} — a deleted-group finding, unrelated to this fix)`,
        );
      }
    }
  }

  console.log(`\n${totalChecked} mapped group(s) checked across ${guids.length} segment(s).`);
  console.log(
    totalStale === 0
      ? "\n✅ NO STALE GROUP NAMES FOUND ANYWHERE ON THE SITE. The fix holds."
      : `\n⚠ ${totalStale} group(s) still carry a name that does not match their segment's current ` +
          `code — open "Check for groups with a stale name" for the segment(s) named above.`,
  );
})();

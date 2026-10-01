/*
 * diagnose-below-unit-levels.js — READ ONLY. Browser console.
 *
 * WHY THIS EXISTS
 * ----------------
 * The CRS Term Abbreviations screen refused to load for "Test New Segment" with:
 *   "Could not load 'Test New Segment' — the term store returned HTTP 404 for the
 *    'Showroom Branch' level's own term set."
 * That comes from `AbbreviationManager.tsx`: a coded BELOW-UNIT level that declares its own
 * SHARED term set (rather than drawing per-unit child terms) gets its own read, and if that
 * set's GUID does not resolve, the WHOLE screen fails to load for the segment — deliberately,
 * because silently skipping an unreadable level would report every code as filled in while a
 * whole level's rows sat unlisted.
 *
 * There is no folder called "Showroom Branch" anywhere in the library — which is CONSISTENT
 * with the 404, not contradictory: if the term set never resolved, reconciliation could never
 * enumerate its terms, so it could never have created a folder for any of them either.
 *
 * This script reads the mode row's `Levels` (and `PendingLevels`, if staged — same precedence
 * the app uses: PendingLevels wins when present) directly from CRS Config, lists every
 * below-Unit level, and for each one that declares its own `termSet`, probes that term set GUID
 * against the term store to confirm exactly which one is broken and how.
 *
 * HOW TO RUN
 *   1. Set SEGMENT_KEY below to the mode row's Title (e.g. "mode_test_new_segment") OR its
 *      StagingFolder (e.g. "TNS") OR its ModeLabel (e.g. "Test New Segment") — any of the three
 *      is matched.
 *   2. Open the site in a browser, signed in as an administrator.
 *   3. F12 -> Console -> paste this whole file -> Enter.
 *   4. Read the table — the broken level's row will show a non-200 status.
 *
 * ⚠ EVERY REQUEST IS A GET. It creates nothing, changes nothing, deletes nothing.
 */
(async () => {
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const SEGMENT_KEY = "Test New Segment"; // Title, StagingFolder, or ModeLabel — any matches

  // Same robust resolution every script here uses since the 2026-09-10 incident: relying on
  // location.pathname alone broke on a page whose path did not start with /sites|/teams, and a
  // wrong web means every request below silently fetches the PAGE's own HTML at HTTP 200 —
  // which then throws an unhelpful JSON parse error instead of naming the real problem.
  const web = (() => {
    if (SITE_OVERRIDE) return SITE_OVERRIDE.replace(/\/+$/, "");
    try {
      if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) {
        return window._spPageContextInfo.webAbsoluteUrl.replace(/\/+$/, "");
      }
    } catch {
      /* fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    return (location.origin + (m ? m[1] : "")).replace(/\/+$/, "");
  })();
  console.log("Site:", web || "(root)");

  const get = async (url) => {
    const r = await fetch(web + url, { headers: { Accept: "application/json;odata=nometadata" } });
    const ct = r.headers.get("content-type") || "";
    if (!ct.includes("application/json")) {
      return { __error: `HTTP ${r.status}, non-JSON response (wrong site/address?)` };
    }
    if (!r.ok) return { __error: `HTTP ${r.status}` };
    return await r.json();
  };

  // This client renames lists — same discovery pattern as every other script here.
  const resolveList = async (suffix) => {
    for (const prefix of ["CRS", "DMS"]) {
      const title = `${prefix} ${suffix}`;
      const r = await get(`/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Title`);
      if (r.__error === undefined) return title;
    }
    return undefined;
  };

  const configList = await resolveList("Config");
  if (!configList) {
    console.log("Could not resolve the Config list (tried 'CRS Config' and 'DMS Config').");
    return;
  }
  console.log("Config list:", configList);

  const norm = (v) => (v || "").trim().toLowerCase();

  const modes = await get(
    `/_api/web/lists/getbytitle('${encodeURIComponent(configList)}')/items` +
      `?$select=Id,Title,ModeLabel,StagingFolder,Levels,PendingLevels&$filter=ConfigType eq 'mode'&$top=200`,
  );
  if (modes.__error !== undefined) {
    console.log(`Could not read ${configList}: ${modes.__error}`);
    return;
  }
  const seg = (modes.value || []).find(
    (m) =>
      norm(m.Title) === norm(SEGMENT_KEY) ||
      norm(m.StagingFolder) === norm(SEGMENT_KEY) ||
      norm(m.ModeLabel) === norm(SEGMENT_KEY),
  );
  if (!seg) {
    console.log(
      `No mode row matched "${SEGMENT_KEY}". Rows found:`,
      (modes.value || []).map((m) => `${m.Title} (${m.ModeLabel || "no label"})`),
    );
    return;
  }
  console.log("Matched segment:", { Title: seg.Title, ModeLabel: seg.ModeLabel, StagingFolder: seg.StagingFolder });

  const staged = (seg.PendingLevels || "").trim();
  const live = (seg.Levels || "").trim();
  const usingSource = staged ? "PendingLevels (STAGED — not yet applied)" : "Levels (live)";
  console.log("Reading from:", usingSource);
  if (staged) {
    console.log(
      "⚠ A structure change is staged for this segment. The abbreviation screen reads the STAGED " +
        "chain too (same precedence used here), so if 'Showroom Branch' only exists in the staged " +
        "chain, it was added there and not yet applied.",
    );
  }

  let levels;
  try {
    levels = JSON.parse(staged || live || "[]");
    if (!Array.isArray(levels)) throw new Error("not an array");
  } catch (e) {
    console.log(`Could not parse the Levels JSON: ${e.message}`);
    console.log("Raw value:", staged || live);
    return;
  }

  console.log(`\n${levels.length} level(s) total in the chain:\n`);
  console.log(JSON.stringify(levels, null, 2));

  const belowUnit = levels.filter((l) => l && l.permissioned === false);
  const shared = belowUnit.filter((l) => (l.termSet || "").trim());
  const perUnit = belowUnit.filter((l) => !(l.termSet || "").trim());

  console.log(
    `\n${belowUnit.length} below-Unit level(s): ${shared.length} shared (own term set), ` +
      `${perUnit.length} per-unit (no term set — draws from each unit's own child terms).\n`,
  );
  if (perUnit.length > 0) {
    console.log("Per-unit levels (nothing to probe — these read fine unless a UNIT's own terms are missing):",
      perUnit.map((l) => l.label));
  }

  if (shared.length === 0) {
    console.log("No shared below-Unit levels declare a term set — nothing to probe.");
    return;
  }

  console.log("\n── Probing each shared level's own term set ──\n");
  const results = [];
  for (const lvl of shared) {
    const setGuid = (lvl.termSet || "").trim();
    const res = await get(`/_api/v2.1/termStore/sets/${setGuid}?$select=id,localizedNames`);
    results.push({
      Level: lvl.label,
      "Term set GUID": setGuid,
      Status: res.__error !== undefined ? res.__error : "OK",
      "Set name": res.__error !== undefined ? "" : ((res.localizedNames || [])[0] || {}).name || "(no name)",
    });
  }
  console.table(results);

  const broken = results.filter((r) => r.Status !== "OK");
  if (broken.length > 0) {
    console.log(
      `\n⚠ ${broken.length} of ${shared.length} shared level(s) FAILED to resolve. This is exactly what ` +
        `blocks the abbreviation screen for this segment. Fix: on the Folder levels screen for this ` +
        `segment, either point the broken level at a real, existing term set, or remove and re-add it ` +
        `pointing at one that exists.`,
    );
  } else {
    console.log("\nAll shared level term sets resolved. If the screen is still failing, re-check for a stale bundle (hard refresh) before assuming a data problem.");
  }
})();

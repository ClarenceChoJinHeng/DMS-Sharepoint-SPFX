/*
 * find-orphaned-abbreviations.js — READ ONLY. Browser console, run on ONE site at a time.
 *
 * WHY THIS EXISTS
 * ----------------
 * Reconciliation reports orphaned `CRS Term Abbreviation` rows in its log, but only as scrolling
 * text mixed in with everything else the run did. This asks the same question directly and prints
 * a clean list — each row's Item Id, Title, Level and Abbreviation — so you can go delete exactly
 * those rows in list settings without hunting through a run log.
 *
 * It uses the SAME rule reconciliation itself uses: walk every segment's LIVE term tree
 * (TermSetGuid, every depth) and collect every term GUID that still resolves. Any abbreviation row
 * whose TermGuid is NOT in that set is orphaned — its term was deleted and never re-created under
 * the same name.
 *
 * HOW TO RUN
 *   1. Open the site in a browser, signed in as an administrator.
 *   2. F12 -> Console -> paste this whole file -> Enter.
 *   3. Read the printed list. Each row names its Item Id — delete those in list settings.
 *
 * ⚠ EVERY REQUEST IS A GET. It creates nothing, changes nothing, deletes nothing. It does NOT
 *   delete anything for you, on purpose — this codebase never auto-deletes an abbreviation row from
 *   a script, because a wrong guess here would take a unit's ONLY copy of its folder-naming rule
 *   with it. Read the list, then delete by hand.
 *
 * ⚠ IF A SEGMENT'S TERM WALK CANNOT COMPLETE (a throttled request, an unreadable term set), that
 *   segment is reported SEPARATELY as "walk incomplete" and its abbreviation rows are NOT judged
 *   either way — an incomplete walk must never be read as "this term does not exist". Same
 *   fail-closed rule the app itself uses for this exact question.
 */
(async () => {
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

  // The list titles are discovered, never assumed — this client renames lists ("DMS" -> "CRS")
  // and a hardcoded title fails SILENTLY. Same two prefixes the app itself tries.
  const resolveList = async (suffix) => {
    for (const prefix of ["CRS", "DMS"]) {
      const title = `${prefix} ${suffix}`;
      const r = await get(`/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Title`);
      if (r.__error === undefined) return title;
    }
    return undefined;
  };

  const configList = await resolveList("Config");
  const abbrevList = await resolveList("Term Abbreviation");
  console.log("Config list:", configList || "NOT FOUND", "| Abbreviation list:", abbrevList || "NOT FOUND");
  if (!configList || !abbrevList) {
    console.log("Could not resolve one of the two lists this needs — nothing else to check.");
    return;
  }

  // Every "mode" row = one segment. TermSetGuid is what its whole term tree hangs off.
  const modes = await get(
    `/_api/web/lists/getbytitle('${encodeURIComponent(configList)}')/items?$select=Title,StagingFolder,TermSetGuid&$filter=ConfigType eq 'mode'&$top=200`,
  );
  if (modes.__error !== undefined) {
    console.log(`Could not read ${configList}: HTTP ${modes.__error}`);
    return;
  }
  const segments = (modes.value || []).filter((m) => (m.TermSetGuid || "").trim().length > 0);
  console.log(`${segments.length} segment(s) found:`, segments.map((s) => s.StagingFolder || s.Title));

  const MAX_REQUESTS_PER_SEGMENT = 400;

  /** Every term GUID under one segment's set, at any depth. Same walk the app itself uses. */
  const walkSegment = async (setGuid) => {
    const guids = new Set();
    let requests = 0;
    let frontier = [""];
    try {
      while (frontier.length > 0) {
        const next = [];
        const kidLists = await Promise.all(
          frontier.map(async (parentId) => {
            if (requests >= MAX_REQUESTS_PER_SEGMENT) return [];
            requests++;
            const url = parentId
              ? `/_api/v2.1/termStore/sets/${setGuid}/terms/${parentId}/children?$select=id`
              : `/_api/v2.1/termStore/sets/${setGuid}/children?$select=id`;
            const res = await get(url);
            if (res.__error !== undefined) throw new Error(`HTTP ${res.__error}`);
            return res.value || [];
          }),
        );
        if (requests >= MAX_REQUESTS_PER_SEGMENT) return { complete: false, guids };
        for (const kids of kidLists) {
          for (const k of kids) {
            const id = (k.id || "").toLowerCase();
            if (id && !guids.has(id)) {
              guids.add(id);
              next.push(k.id);
            }
          }
        }
        frontier = next;
      }
      return { complete: true, guids };
    } catch {
      return { complete: false, guids };
    }
  };

  const liveGuids = new Set();
  const incompleteSegments = [];
  for (const seg of segments) {
    const walk = await walkSegment(seg.TermSetGuid.trim());
    for (const g of walk.guids) liveGuids.add(g);
    if (!walk.complete) incompleteSegments.push(seg.StagingFolder || seg.Title);
  }
  console.log(`Walked ${liveGuids.size} live term(s) across ${segments.length} segment(s).`);
  if (incompleteSegments.length > 0) {
    console.log(
      `⚠ Could not fully walk: ${incompleteSegments.join(", ")}. Abbreviation rows that would have`,
      "matched terms in those segments are NOT judged below — re-run once the term store responds.",
    );
  }

  const rows = await get(
    `/_api/web/lists/getbytitle('${encodeURIComponent(abbrevList)}')/items?$select=Id,TermGuid,Title,Level,Abbreviation&$top=5000`,
  );
  if (rows.__error !== undefined) {
    console.log(`Could not read ${abbrevList}: HTTP ${rows.__error}`);
    return;
  }

  const orphaned = (rows.value || []).filter((r) => {
    const guid = (r.TermGuid || "").trim().toLowerCase();
    return guid.length > 0 && !liveGuids.has(guid);
  });

  if (orphaned.length === 0) {
    console.log("No orphaned abbreviation rows found.");
    return;
  }

  console.log(`\n${orphaned.length} ORPHANED ROW(S) — delete these by hand in list settings:\n`);
  console.table(
    orphaned.map((r) => ({
      Id: r.Id,
      Title: r.Title,
      Level: r.Level,
      Abbreviation: r.Abbreviation,
      TermGuid: r.TermGuid,
    })),
  );
})();

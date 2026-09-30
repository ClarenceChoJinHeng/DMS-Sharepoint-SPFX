// Lists rows left behind by retired segments or deleted terms. READ ONLY: it deletes nothing.
// Checks CRS Term Abbreviation, CRS Folder Map and CRS Group Map against the LIVE segments
// (CRS Config "mode" rows) and their term sets, including shared levels named in Levels/PendingLevels.
// Run: open the CRS site as a site owner, F12 → Console, paste, Enter. Then delete the listed rows by hand.
(async () => {
  // ---- CONFIG ----
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const MAX_TERM_REQUESTS = 3000;

  // ---- HELPERS ----
  function resolveWeb() {
    if (SITE_OVERRIDE) return SITE_OVERRIDE;
    try {
      if (window._spPageContextInfo) return window._spPageContextInfo.webAbsoluteUrl;
    } catch (e) {
      // not on a classic page
    }
    const match = location.pathname.match(/^\/(sites|teams)\/[^/]+/);
    return location.origin + (match ? match[0] : "");
  }
  const web = resolveWeb();
  const bust = () => "_=" + Date.now() + Math.random().toString(36).slice(2);
  const lower = (s) => (s || "").trim().toLowerCase().replace(/[{}]/g, "");

  async function getJson(pathOrUrl) {
    const base = pathOrUrl.startsWith("http") ? pathOrUrl : web + pathOrUrl;
    const url = base + (base.includes("?") ? "&" : "?") + bust();
    const res = await fetch(url, { headers: { Accept: "application/json;odata=nometadata", "Cache-Control": "no-cache" } });
    const type = res.headers.get("content-type") || "";
    if (!res.ok || !type.includes("json")) return { ok: false, status: res.status };
    return { ok: true, data: await res.json() };
  }

  // Every page of a list or term collection. undefined = could not read (never "empty").
  async function getAll(path) {
    const rows = [];
    let next = path;
    while (next) {
      const r = await getJson(next);
      if (!r.ok) return undefined;
      rows.push(...(r.data.value || []));
      next = r.data["odata.nextLink"] || r.data["@odata.nextLink"] || "";
    }
    return rows;
  }

  async function findList(suffix) {
    for (const prefix of ["CRS", "DMS"]) {
      const title = prefix + " " + suffix;
      const r = await getJson("/_api/web/lists/getbytitle('" + encodeURIComponent(title) + "')?$select=Title");
      if (r.ok) return title;
    }
    return undefined;
  }
  const itemsUrl = (title, select) => "/_api/web/lists/getbytitle('" + encodeURIComponent(title) + "')/items?$select=" + select + "&$top=500";

  // Reads with the extra columns first, then without them if one does not exist.
  async function getAllLadder(title, selects) {
    for (const select of selects) {
      const rows = await getAll(itemsUrl(title, select));
      if (rows !== undefined) return rows;
    }
    return undefined;
  }

  // ---- 1. LIVE SEGMENTS ----
  const configList = await findList("Config");
  const abbrevList = await findList("Term Abbreviation");
  const folderMapList = await findList("Folder Map");
  const groupMapList = await findList("Group Map");
  console.log("Site:", web);
  console.log("Lists:", { configList, abbrevList, folderMapList, groupMapList });
  if (!configList) {
    console.error("CRS Config not found. Nothing concluded.");
    return;
  }

  const modeRows = await getAllLadder(configList, [
    "Title,ConfigType,TermSetGuid,StagingFolder,Levels,PendingLevels",
    "Title,ConfigType,TermSetGuid,StagingFolder,Levels",
  ]);
  if (modeRows === undefined) {
    console.error("Could not read CRS Config. Nothing concluded.");
    return;
  }
  const modes = modeRows.filter((r) => lower(r.ConfigType) === "mode");
  const liveFolders = new Set(modes.map((m) => lower(m.StagingFolder)).filter(Boolean));
  console.log("Live segments:", modes.map((m) => m.StagingFolder || m.Title).join(", "));

  // Term sets to walk: each segment's own set + every shared-level set in Levels / PendingLevels.
  const setGuids = new Set();
  for (const m of modes) {
    if (lower(m.TermSetGuid)) setGuids.add(lower(m.TermSetGuid));
    for (const field of ["Levels", "PendingLevels"]) {
      let levels = [];
      try {
        levels = JSON.parse(m[field] || "[]");
      } catch (e) {
        levels = [];
      }
      for (const lvl of Array.isArray(levels) ? levels : []) {
        if (lvl && lower(lvl.termSet)) setGuids.add(lower(lvl.termSet));
      }
    }
  }

  // ---- 2. WALK EVERY LIVE TERM ----
  const liveTerms = new Set(setGuids); // segment-tier Group Map rows point at the set itself
  let requests = 0;
  let walkComplete = true;
  for (const setGuid of setGuids) {
    let frontier = [""];
    while (frontier.length > 0 && walkComplete) {
      const next = [];
      for (const parentId of frontier) {
        if (requests >= MAX_TERM_REQUESTS) {
          walkComplete = false;
          break;
        }
        requests++;
        const path = parentId
          ? "/_api/v2.1/termStore/sets/" + setGuid + "/terms/" + parentId + "/children?$select=id"
          : "/_api/v2.1/termStore/sets/" + setGuid + "/children?$select=id";
        const kids = await getAll(path);
        if (kids === undefined) {
          console.warn("Could not read term set " + setGuid + " (parent " + (parentId || "top") + ")");
          walkComplete = false;
          break;
        }
        for (const k of kids) {
          const id = lower(k.id);
          if (id && !liveTerms.has(id)) {
            liveTerms.add(id);
            next.push(k.id);
          }
        }
      }
      frontier = next;
    }
  }
  console.log("Walked " + setGuids.size + " term set(s), " + liveTerms.size + " live term(s), " + requests + " request(s).");
  if (!walkComplete) {
    console.error("STOPPED: the term store could not be read completely. Judging now would mark live rows as leftovers. Run again later. Nothing concluded.");
    return;
  }

  // ---- 3. FIND LEFTOVERS ----
  const result = {};

  if (abbrevList) {
    const rows = await getAllLadder(abbrevList, ["Id,Title,Abbreviation,Level,TermGuid", "Id,Title,Abbreviation,TermGuid"]);
    if (rows === undefined) console.error("Could not read " + abbrevList + ".");
    else result.abbreviations = rows
      .filter((r) => lower(r.TermGuid) && !liveTerms.has(lower(r.TermGuid)))
      .map((r) => ({ list: abbrevList, Id: r.Id, Title: r.Title, Abbreviation: r.Abbreviation, Level: r.Level, TermGuid: r.TermGuid }));
  }

  if (folderMapList) {
    const rows = await getAllLadder(folderMapList, ["Id,Title,TermGuid,Section"]);
    if (rows === undefined) console.error("Could not read " + folderMapList + ".");
    else result.folderMap = rows
      .filter((r) => (lower(r.TermGuid) && !liveTerms.has(lower(r.TermGuid))) || (lower(r.Section) && !liveFolders.has(lower(r.Section))))
      .map((r) => ({ list: folderMapList, Id: r.Id, Title: r.Title, Section: r.Section, TermGuid: r.TermGuid }));
  }

  if (groupMapList) {
    const rows = await getAllLadder(groupMapList, ["Id,GroupName,Role,Scope,Segment,UnitTermGuid", "Id,GroupName,Role,Segment,UnitTermGuid"]);
    if (rows === undefined) console.error("Could not read " + groupMapList + ".");
    else {
      const scopeOf = (r) => lower(Array.isArray(r.Scope) ? r.Scope[0] : r.Scope) || "folder";
      const folderRows = rows.filter((r) => scopeOf(r) === "folder" && lower(r.UnitTermGuid));
      const leftovers = folderRows.filter((r) => !liveTerms.has(lower(r.UnitTermGuid)));
      result.groupMap = leftovers.map((r) => ({ list: groupMapList, Id: r.Id, GroupName: r.GroupName, Role: r.Role, Segment: r.Segment, UnitTermGuid: r.UnitTermGuid }));
      // Groups whose ONLY folder rows are leftovers: candidates to delete in Group Management.
      const liveGroupNames = new Set(folderRows.filter((r) => liveTerms.has(lower(r.UnitTermGuid))).map((r) => lower(r.GroupName)));
      result.groupsOnlyInRetiredSegments = [...new Set(leftovers.map((r) => r.GroupName))]
        .filter((g) => !liveGroupNames.has(lower(g)))
        .map((g) => ({ GroupName: g }));
    }
  }

  // ---- 4. REPORT ----
  window.__result = result;
  for (const key of ["abbreviations", "folderMap", "groupMap", "groupsOnlyInRetiredSegments"]) {
    const rows = result[key];
    if (rows === undefined) continue;
    console.log("\n" + key + ": " + rows.length + " leftover(s)");
    if (rows.length > 0) console.table(rows);
  }
  console.log("\nREAD ONLY: nothing was changed. Delete list rows by Id in each list; delete groups in Group Management.");
})();

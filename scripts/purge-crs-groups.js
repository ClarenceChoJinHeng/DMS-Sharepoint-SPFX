// Launch-day group purge: deletes every CRS unit group and its CRS Group Map rows, and empties the
// site-entry group (CRS_SITE_MEMBERS). DRY RUN unless the DO_ flags are true.
// Never touches: the site's Owners / Members / Visitors groups, or any group that is not a CRS group.
// Rebuild afterwards with Group Management → Create Group (bulk), then Folder Reconciliation.
// Run: open the CRS site as a site owner, F12 → Console, paste, Enter.
(async () => {
  // ---- CONFIG ----
  const EXPECTED_SITE = "/sites/CRS";     // refuses to run anywhere else
  const DO_DELETE_GROUPS = false;         // delete the CRS unit groups (members and permissions go with them)
  const DO_DELETE_GROUP_MAP_ROWS = false; // recycle their CRS Group Map rows
  const DO_EMPTY_SITE_ENTRY = false;      // remove every person from CRS_SITE_MEMBERS (the group itself stays)

  const SITE_ENTRY_TITLES = ["GDC_SITE_MEMBERS", "CRS_SITE_MEMBERS", "DMS_SITE_MEMBERS"]; // shared/naming.ts GROUP_CANDIDATE_PREFIXES
  // Every role suffix CRS has used for a unit group name (shared/groupMapModel.ts ROLE_SUFFIXES).
  const ROLE_SUFFIXES = [
    "_DELETER_DOCUMENTS", "_DELETER_STAGING", "_APPROVER", "_UPLOADER", "_C_LEVEL", "_SEGVIEW",
    "_DELS", "_APR", "_UPL", "_DEL", "_UPL_HIGHLY_CONFIDENTIAL", "_UPL_HC", "_DELS_HIGHLY_CONFIDENTIAL",
    "_DELS_HC", "_VIEWER_HIGHLY_CONFIDENTIAL", "_EMPLOYEE", "_HOD", "_VIEWER_HC", "_DEPARTMENT_VIEWER",
    "_DEPTVIEW", "_APR_HIGHLY_CONFIDENTIAL", "_APR_HC", "_APPROVER_HIGHLY_CONFIDENTIAL",
    "_UPLOADER_HIGHLY_CONFIDENTIAL", "_VIEWER", "_HC",
  ];
  const EXACT_NAMES = ["C_LEVEL_GLOBAL"];

  // ---- HELPERS ----
  function resolveWeb() {
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
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const up = (s) => (s || "").trim().toUpperCase();

  async function getJson(path) {
    const url = web + path + (path.includes("?") ? "&" : "?") + bust();
    const res = await fetch(url, { headers: { Accept: "application/json;odata=nometadata", "Cache-Control": "no-cache" } });
    const type = res.headers.get("content-type") || "";
    if (!res.ok || !type.includes("json")) return { ok: false, status: res.status };
    return { ok: true, data: await res.json() };
  }

  async function getAll(path) {
    const rows = [];
    let next = path;
    while (next) {
      const r = await getJson(next);
      if (!r.ok) return undefined;
      rows.push(...(r.data.value || []));
      next = r.data["odata.nextLink"] ? r.data["odata.nextLink"].replace(web, "") : "";
    }
    return rows;
  }

  let digest = "";
  async function getDigest() {
    const res = await fetch(web + "/_api/contextinfo", { method: "POST", headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) throw new Error("Could not get a request digest (HTTP " + res.status + ")");
    digest = (await res.json()).FormDigestValue;
  }

  async function post(path) {
    for (let attempt = 0; attempt < 4; attempt++) {
      const res = await fetch(web + path, {
        method: "POST",
        headers: { Accept: "application/json;odata=nometadata", "X-RequestDigest": digest },
      });
      if (res.status === 429 || res.status === 503) {
        await sleep(Number(res.headers.get("Retry-After") || 5) * 1000);
        continue;
      }
      return res;
    }
    return { ok: false, status: 429 };
  }

  const isCrsGroupName = (title) => {
    const t = up(title);
    return EXACT_NAMES.includes(t) || ROLE_SUFFIXES.some((s) => t.endsWith(s));
  };

  // ---- SAFETY ----
  if (!web.toLowerCase().endsWith(EXPECTED_SITE.toLowerCase())) {
    console.error("STOPPED: this is " + web + ", not " + EXPECTED_SITE + ". Nothing was changed.");
    return;
  }
  console.log("Site:", web);
  console.log("DO_DELETE_GROUPS:", DO_DELETE_GROUPS, "| DO_DELETE_GROUP_MAP_ROWS:", DO_DELETE_GROUP_MAP_ROWS, "| DO_EMPTY_SITE_ENTRY:", DO_EMPTY_SITE_ENTRY);

  // ---- READ ----
  const groups = await getAll("/_api/web/sitegroups?$select=Id,Title&$top=5000");
  const owner = await getJson("/_api/web/AssociatedOwnerGroup?$select=Id,Title");
  const member = await getJson("/_api/web/AssociatedMemberGroup?$select=Id,Title");
  const visitor = await getJson("/_api/web/AssociatedVisitorGroup?$select=Id,Title");
  if (groups === undefined || !owner.ok) {
    console.error("Could not read the site's groups or its Owners group. Nothing was changed.");
    return;
  }
  const protectedIds = new Set([owner, member, visitor].filter((g) => g.ok).map((g) => g.data.Id));

  let mapTitle;
  for (const t of ["CRS Group Map", "DMS Group Map"]) {
    if ((await getJson("/_api/web/lists/getbytitle('" + encodeURIComponent(t) + "')?$select=Title")).ok) {
      mapTitle = t;
      break;
    }
  }
  let mapRows = [];
  if (mapTitle) {
    const base = "/_api/web/lists/getbytitle('" + encodeURIComponent(mapTitle) + "')/items?$top=500&$select=";
    mapRows = (await getAll(base + "Id,GroupId,GroupName,Role,Scope")) || (await getAll(base + "Id,GroupId,GroupName,Role"));
    if (mapRows === undefined) {
      console.error("Could not read " + mapTitle + ". Nothing was changed.");
      return;
    }
  }
  const mappedIds = new Set(mapRows.map((r) => String(r.GroupId || "").trim()).filter(Boolean));

  const siteEntry = groups.find((g) => SITE_ENTRY_TITLES.includes(up(g.Title)));
  const toDelete = groups.filter(
    (g) =>
      !protectedIds.has(g.Id) &&
      !(siteEntry && g.Id === siteEntry.Id) &&
      !up(g.Title).endsWith("_SITE_MEMBERS") &&
      (mappedIds.has(String(g.Id)) || isCrsGroupName(g.Title)),
  );
  const deleteIds = new Set(toDelete.map((g) => String(g.Id)));
  const deleteTitles = new Set(toDelete.map((g) => up(g.Title)));
  const liveIds = new Set(groups.map((g) => String(g.Id)));
  const siteEntryRowIds = new Set(siteEntry ? [String(siteEntry.Id)] : []);
  const rowsToDelete = mapRows.filter((r) => {
    const gid = String(r.GroupId || "").trim();
    if (siteEntryRowIds.has(gid)) return false;
    return deleteIds.has(gid) || deleteTitles.has(up(r.GroupName)) || (gid && !liveIds.has(gid));
  });
  let siteEntryUsers = [];
  if (siteEntry) {
    siteEntryUsers = (await getAll("/_api/web/sitegroups(" + siteEntry.Id + ")/users?$select=Id,Title,LoginName")) || [];
  }

  // ---- DRY-RUN REPORT ----
  const kept = groups.filter((g) => !deleteIds.has(String(g.Id)));
  console.log("Groups to DELETE: " + toDelete.length);
  console.table(toDelete.map((g) => ({ Id: g.Id, Title: g.Title, inGroupMap: mappedIds.has(String(g.Id)) })));
  console.log("Groups KEPT: " + kept.length);
  console.table(kept.map((g) => ({ Id: g.Id, Title: g.Title, why: protectedIds.has(g.Id) ? "site Owners/Members/Visitors" : siteEntry && g.Id === siteEntry.Id ? "site entry (members emptied)" : "not a CRS group" })));
  console.log((mapTitle || "Group Map") + " rows to delete: " + rowsToDelete.length + " of " + mapRows.length);
  console.log("Site-entry group: " + (siteEntry ? siteEntry.Title + " (" + siteEntryUsers.length + " members to remove)" : "NOT FOUND"));
  window.__result = { toDelete, kept, rowsToDelete, siteEntryUsers };

  if (!DO_DELETE_GROUPS && !DO_DELETE_GROUP_MAP_ROWS && !DO_EMPTY_SITE_ENTRY) {
    console.log("DRY RUN only. Nothing was changed. Check both tables, then set the DO_ flags.");
    return;
  }
  if (!siteEntry) {
    console.error("STOPPED: the site-entry group was not found, so its access rows cannot be protected. Nothing was changed.");
    return;
  }

  await getDigest();
  const failures = [];

  if (DO_DELETE_GROUPS) {
    for (const g of toDelete) {
      const res = await post("/_api/web/sitegroups/removebyid(" + g.Id + ")");
      console.log((res.ok ? "✓" : "✗ HTTP " + res.status) + " deleted group " + g.Title);
      if (!res.ok) failures.push("group " + g.Title);
    }
  }

  if (DO_DELETE_GROUP_MAP_ROWS && mapTitle) {
    for (const r of rowsToDelete) {
      const res = await post("/_api/web/lists/getbytitle('" + encodeURIComponent(mapTitle) + "')/items(" + r.Id + ")/recycle()");
      if (!res.ok) failures.push("Group Map row " + r.Id);
    }
    console.log("✓ Group Map rows processed: " + rowsToDelete.length);
  }

  if (DO_EMPTY_SITE_ENTRY && siteEntry) {
    for (const u of siteEntryUsers) {
      const res = await post("/_api/web/sitegroups(" + siteEntry.Id + ")/users/removebyid(" + u.Id + ")");
      if (!res.ok) failures.push("site-entry member " + (u.Title || u.LoginName));
    }
    console.log("✓ " + siteEntry.Title + ": " + siteEntryUsers.length + " member(s) processed");
  }

  if (failures.length > 0) {
    console.warn(failures.length + " item(s) failed. They are on window.__failures. Run the same step again to retry.");
    window.__failures = failures;
  } else {
    console.log("Done, no failures.");
  }
})();

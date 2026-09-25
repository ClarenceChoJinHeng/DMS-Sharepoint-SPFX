// <What this checks>. READ ONLY unless DO_WRITE is true.
// Run: open the CRS site, F12 → Console, paste, Enter.
(async () => {
  // ---- CONFIG ----
  const SITE_OVERRIDE = "";      // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const DO_WRITE = false;        // repairs only: keep false until the dry run looks right

  // ---- HELPERS ----
  function resolveWeb() {
    if (SITE_OVERRIDE) return SITE_OVERRIDE;
    try {
      if (window._spPageContextInfo) return window._spPageContextInfo.webAbsoluteUrl;
    } catch (e) {
      // not on a classic page; fall through
    }
    const match = location.pathname.match(/^\/(sites|teams)\/[^/]+/);
    return location.origin + (match ? match[0] : "");
  }
  const web = resolveWeb();
  const bust = () => "_=" + Date.now() + Math.random().toString(36).slice(2);

  async function getJson(path) {
    const url = web + path + (path.includes("?") ? "&" : "?") + bust();
    const res = await fetch(url, {
      headers: { Accept: "application/json;odata=nometadata", "Cache-Control": "no-cache" },
    });
    const type = res.headers.get("content-type") || "";
    if (!res.ok || !type.includes("json")) {
      return { ok: false, status: res.status, body: await res.text() };
    }
    return { ok: true, data: await res.json() };
  }

  // Reads every page. Returns undefined if any page fails (unknown, never "empty").
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

  // ---- MAIN ----
  console.log("Site:", web, "| DO_WRITE:", DO_WRITE);
  // Example:
  // const rows = await getAll("/_api/web/lists/getbytitle('CRS Submissions')/items?$select=Id,Title&$top=500");
  // if (rows === undefined) { console.warn("Could not read the list. Nothing concluded."); return; }
  // window.__result = rows;
  // console.table(rows.slice(0, 20));
})();

// Counts documents whose Document Type / Year / Confidentiality value is stored with a NUMBER as its name
// (Label === WssId), across all six libraries. READ ONLY.
// Run: open the CRS site (any page, wait until it has fully loaded), F12 → Console, paste, Enter.
(async () => {
  // ---- CONFIG ----
  const SITE_OVERRIDE = "";      // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const LIBRARY_SEGMENTS = ["ApprovalDocument", "Shared Documents", "HCApprovalDocument", "HCDocuments", "Archive", "HCArchive"];
  const FIELDS = ["Document_x0020_Type", "Year", "Confidentiality_x0020_Level"];

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
  const web = resolveWeb().replace(/\/+$/, "");
  const webPath = web.replace(/^https:\/\/[^/]+/, "");
  const bust = () => "_=" + Date.now() + Math.random().toString(36).slice(2);

  async function getJson(path) {
    const base = path.indexOf("http") === 0 ? path : web + path;
    const url = base + (base.includes("?") ? "&" : "?") + bust();
    try {
      const res = await fetch(url, {
        headers: { Accept: "application/json;odata=nometadata", "Cache-Control": "no-cache" },
      });
      const type = res.headers.get("content-type") || "";
      if (!res.ok || !type.includes("json")) return { ok: false, status: res.status };
      return { ok: true, data: await res.json() };
    } catch (e) {
      return { ok: false, status: "network: " + e.message };
    }
  }

  // Reads every page. rows is undefined if any page fails (unknown, never "empty").
  async function getAll(path) {
    const rows = [];
    let next = path;
    while (next) {
      const r = await getJson(next);
      if (!r.ok) return { rows: undefined, status: r.status };
      rows.push(...(r.data.value || []));
      if (rows.length % 500 === 0) console.log("  …", rows.length, "items read");
      next = r.data["odata.nextLink"] || "";
    }
    return { rows };
  }

  const isNumberName = (v) => !!v && typeof v === "object" && v.Label !== undefined && String(v.Label) === String(v.WssId);

  // ---- MAIN ----
  console.log("Site:", web, "| READ ONLY");
  const summary = [];
  const bad = [];
  for (const seg of LIBRARY_SEGMENTS) {
    const q = "?@u='" + encodeURIComponent(webPath + "/" + seg) + "'";
    const head = await getJson("/_api/web/getlist(@u)" + q + "&$select=Title,ItemCount");
    if (!head.ok) {
      summary.push({ library: seg, documents: "could not read (" + head.status + ")" });
      continue;
    }
    console.log("Reading", head.data.Title, "(" + head.data.ItemCount + " items incl. folders)");
    const read = await getAll("/_api/web/getlist(@u)/items" + q + "&$select=Id,FileLeafRef,FSObjType," + FIELDS.join(",") + "&$top=500");
    if (read.rows === undefined) {
      summary.push({ library: head.data.Title, documents: "could not read (" + read.status + ")" });
      continue;
    }
    const docs = read.rows.filter((r) => r.FSObjType === 0);
    const counts = { library: head.data.Title, documents: docs.length };
    for (const f of FIELDS) {
      const hits = docs.filter((d) => isNumberName(d[f]));
      counts[f] = hits.length;
      for (const h of hits) {
        bad.push({ library: head.data.Title, id: h.Id, file: h.FileLeafRef, field: f, storedName: h[f].Label, termGuid: h[f].TermGuid });
      }
    }
    summary.push(counts);
  }
  window.__result = { summary, bad };
  console.table(summary);
  if (bad.length) {
    console.log("Documents with a number stored as the term name (first 20; all on window.__result.bad):");
    console.table(bad.slice(0, 20));
  } else if (summary.every((s) => typeof s.documents === "number")) {
    console.log("No document has a number stored as a term name, in all six libraries.");
  } else {
    console.log("No number names found, but some libraries could not be read (see table): nothing concluded for those.");
  }
})();

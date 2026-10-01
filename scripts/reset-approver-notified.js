// Clears ApproverNotifiedAt on one upload's CRS Submissions rows, so the bundled approver email sends again (testing).
// Dry run by default: it lists recent uploads first. Set SUBMISSION_REF, then DO_WRITE = true, and run again.
// Run: open the site, F12 → Console, paste, Enter.
(async () => {
  // ---- CONFIG ----
  const SITE_OVERRIDE = "";      // e.g. "https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting"
  const SUBMISSION_REF = "";     // e.g. "SUB-20260930-XXXX" (from the dry-run table)
  const DO_WRITE = false;        // keep false until the dry run looks right
  const DAYS_BACK = 3;           // the bundled flow only looks back 3 days

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
      if (!r.ok) { console.warn("Read failed:", r.status, (r.body || "").slice(0, 300)); return undefined; }
      rows.push(...(r.data.value || []));
      next = r.data["odata.nextLink"] ? r.data["odata.nextLink"].replace(web, "") : "";
    }
    return rows;
  }

  async function getDigest() {
    const res = await fetch(web + "/_api/contextinfo", {
      method: "POST",
      headers: { Accept: "application/json;odata=nometadata" },
    });
    return res.ok ? (await res.json()).FormDigestValue : undefined;
  }

  // ---- MAIN ----
  console.log("Site:", web, "| SUBMISSION_REF:", SUBMISSION_REF || "(none)", "| DO_WRITE:", DO_WRITE);
  const since = new Date(Date.now() - DAYS_BACK * 86400000).toISOString();
  const list = "/_api/web/lists/getbytitle('CRS%20Submissions')";
  const rows = await getAll(
    list + "/items?$select=Id,SubmissionRef,FileName,LibraryTitle,ApproverNotifiedAt,UploadedAt" +
      "&$filter=UploadedAt ge datetime'" + since + "'&$top=500",
  );
  if (rows === undefined) { console.warn("Could not read CRS Submissions. Nothing changed."); return; }

  const byRef = {};
  for (const r of rows) {
    const ref = r.SubmissionRef || "(no ref)";
    byRef[ref] = byRef[ref] || { SubmissionRef: ref, files: 0, notified: 0, libraries: new Set() };
    byRef[ref].files += 1;
    if (r.ApproverNotifiedAt) byRef[ref].notified += 1;
    byRef[ref].libraries.add(r.LibraryTitle || "?");
  }
  console.table(Object.values(byRef).map((g) => ({ ...g, libraries: [...g.libraries].join(" + ") })));

  if (!SUBMISSION_REF) { console.log("Dry run only. Put a SubmissionRef from the table in SUBMISSION_REF and run again."); return; }
  const targets = rows.filter((r) => r.SubmissionRef === SUBMISSION_REF && r.ApproverNotifiedAt);
  console.log(targets.length + " row(s) of " + SUBMISSION_REF + " have ApproverNotifiedAt set and would be cleared.");
  window.__result = targets;
  if (!DO_WRITE || targets.length === 0) { console.log("Nothing written (DO_WRITE is false or no rows)."); return; }

  const digest = await getDigest();
  if (!digest) { console.warn("Could not get a request digest. Nothing changed."); return; }
  let done = 0;
  for (const r of targets) {
    const res = await fetch(web + list + "/items(" + r.Id + ")", {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
      body: JSON.stringify({ ApproverNotifiedAt: null }),
    });
    if (!res.ok) console.warn("Row " + r.Id + " failed:", res.status, (await res.text()).slice(0, 200));
    else done += 1;
    if (done % 25 === 0) console.log("cleared " + done + " / " + targets.length);
  }
  console.log("Done. Cleared " + done + " of " + targets.length + ".");
})();

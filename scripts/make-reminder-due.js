// Makes one upload's pending files due for the bundled approver reminder now (testing). READ ONLY unless DO_WRITE is true.
// Dry run lists pending uploads. Set SUBMISSION_ID, then DO_WRITE = true, and run again.
// ONLY_ONE_FILE = true makes just one file due, to prove the reminder still covers (and resets) the whole upload.
// Run: open the site, F12 → Console, paste, Enter.
(async () => {
  // ---- CONFIG ----
  const SITE_OVERRIDE = "";      // e.g. "https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting"
  const SUBMISSION_ID = "";      // from the dry-run table
  const ONLY_ONE_FILE = true;    // false = every pending file of the upload
  const DO_WRITE = false;        // keep false until the dry run looks right
  const LIBRARY_SEGMENTS = ["ApprovalDocument", "HCApprovalDocument"];

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
  const webPath = new URL(web).pathname.replace(/\/$/, "");
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

  // Libraries by URL segment, never by title.
  const listAlias = (segment) => "@u='" + encodeURIComponent(webPath + "/" + segment) + "'";

  // ---- MAIN ----
  console.log("Site:", web, "| SUBMISSION_ID:", SUBMISSION_ID || "(none)", "| DO_WRITE:", DO_WRITE);
  const pending = [];
  for (const segment of LIBRARY_SEGMENTS) {
    const rows = await getAll(
      "/_api/web/GetList(@u)/items?" + listAlias(segment) +
        "&$select=Id,FileLeafRef,SubmissionId,Created,NextReminderAt,UnitTid" +
        "&$filter=FSObjType eq 0 and OData__ModerationStatus eq 2&$top=500",
    );
    if (rows === undefined) { console.warn("Could not read " + segment + ". Nothing changed."); return; }
    for (const r of rows) pending.push({ ...r, segment });
  }

  const byId = {};
  for (const r of pending) {
    const id = r.SubmissionId || "(no SubmissionId)";
    byId[id] = byId[id] || { SubmissionId: id, files: 0, libraries: new Set(), created: r.Created, nextReminder: new Set() };
    byId[id].files += 1;
    byId[id].libraries.add(r.segment);
    byId[id].nextReminder.add(r.NextReminderAt || "(blank)");
  }
  console.table(Object.values(byId).map((g) => ({
    ...g, libraries: [...g.libraries].join(" + "), nextReminder: [...g.nextReminder].join(", "),
  })));

  if (!SUBMISSION_ID) { console.log("Dry run only. Put a SubmissionId from the table in SUBMISSION_ID and run again."); return; }
  const upload = pending.filter((r) => r.SubmissionId === SUBMISSION_ID);
  const targets = ONLY_ONE_FILE ? upload.slice(0, 1) : upload;
  console.log(targets.length + " of " + upload.length + " pending file(s) of " + SUBMISSION_ID + " would get NextReminderAt = now - 1 min:");
  console.table(targets.map((r) => ({ Id: r.Id, library: r.segment, file: r.FileLeafRef, NextReminderAt: r.NextReminderAt })));
  window.__result = targets;
  if (!DO_WRITE || targets.length === 0) { console.log("Nothing written (DO_WRITE is false or no files)."); return; }

  const digest = await getDigest();
  if (!digest) { console.warn("Could not get a request digest. Nothing changed."); return; }
  const dueNow = new Date(Date.now() - 60000).toISOString();
  let done = 0;
  for (const r of targets) {
    const res = await fetch(web + "/_api/web/GetList(@u)/items(" + r.Id + ")?" + listAlias(r.segment), {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
      body: JSON.stringify({ NextReminderAt: dueNow }),
    });
    if (!res.ok) console.warn("File " + r.Id + " failed:", res.status, (await res.text()).slice(0, 200));
    else done += 1;
    if (done % 25 === 0) console.log("set " + done + " / " + targets.length);
  }
  console.log("Done. Set " + done + " of " + targets.length + ". The reminder flow picks it up on its next 30-minute run (or Test > Manually).");
})();

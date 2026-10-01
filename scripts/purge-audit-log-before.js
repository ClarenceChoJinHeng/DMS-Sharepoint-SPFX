// Recycles CRS Audit Log rows whose event time is BEFORE a cutoff. Rows at or after the cutoff are kept.
// DRY RUN unless DO_RECYCLE is true. Recycled rows go to the recycle bin (restorable); this never empties it.
// Run: open the CRS site as a site owner, F12 → Console, paste, Enter.
(async () => {
  // ---- CONFIG ----
  const EXPECTED_SITE = "/sites/CRS"; // refuses to run anywhere else
  const CUTOFF = "2026-09-30T05:14:00+08:00"; // Malaysia time (UTC+8); rows before this go
  const DO_RECYCLE = false; // keep false until the dry run looks right
  const LIST_TITLES = ["CRS Audit Log", "DMS Audit Log"];

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
  const local = (iso) =>
    iso ? new Date(iso).toLocaleString("en-GB", { timeZone: "Asia/Kuala_Lumpur" }) : "(none)";

  async function getJson(path) {
    const url = web + path + (path.includes("?") ? "&" : "?") + bust();
    const res = await fetch(url, {
      headers: { Accept: "application/json;odata=nometadata", "Cache-Control": "no-cache" },
    });
    const type = res.headers.get("content-type") || "";
    if (!res.ok || !type.includes("json")) return { ok: false, status: res.status, body: await res.text() };
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

  // POST, retrying on throttling (429/503).
  async function post(path, digest) {
    for (let attempt = 0; attempt < 4; attempt++) {
      const res = await fetch(web + path, {
        method: "POST",
        headers: { Accept: "application/json;odata=nometadata", "X-RequestDigest": digest },
      });
      if (res.status === 429 || res.status === 503) {
        const wait = Number(res.headers.get("Retry-After") || 5) * 1000;
        console.warn("Throttled, waiting " + wait / 1000 + "s");
        await sleep(wait);
        continue;
      }
      return res;
    }
    return { ok: false, status: 429 };
  }

  // ---- SAFETY ----
  if (!web.toLowerCase().endsWith(EXPECTED_SITE.toLowerCase())) {
    console.error("STOPPED: this is " + web + ", not " + EXPECTED_SITE + ". Nothing was changed.");
    return;
  }
  const cutoff = new Date(CUTOFF);
  if (isNaN(cutoff.getTime())) { console.error("STOPPED: CUTOFF is not a valid date. Nothing was changed."); return; }
  console.log("Site:", web, "| cutoff:", local(cutoff.toISOString()), "(Malaysia time) | DO_RECYCLE:", DO_RECYCLE);

  // ---- READ ----
  let listTitle = "";
  for (const t of LIST_TITLES) {
    const r = await getJson("/_api/web/lists/getbytitle('" + encodeURIComponent(t) + "')?$select=Title");
    if (r.ok) { listTitle = t; break; }
  }
  if (!listTitle) { console.error("No Audit Log list found. Nothing was changed."); return; }
  const base = "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')";
  const rows = await getAll(base + "/items?$select=Id,EventTime,Created,EventType,ItemName,ActorName&$top=500");
  if (rows === undefined) { console.error("Could not read " + listTitle + ". Nothing was changed."); return; }

  // Event time decides; a row without one falls back to when it was written.
  const when = (r) => new Date(r.EventTime || r.Created);
  const before = rows.filter((r) => when(r) < cutoff).sort((a, b) => when(a) - when(b));
  const kept = rows.filter((r) => !(when(r) < cutoff)).sort((a, b) => when(a) - when(b));
  const show = (r) => ({ Id: r.Id, when: local(r.EventTime || r.Created), event: r.EventType, item: r.ItemName, by: r.ActorName });

  console.log(listTitle + ": " + rows.length + " rows. To recycle (before cutoff): " + before.length + ". Kept: " + kept.length + ".");
  if (before.length) {
    console.log("Oldest and newest rows that WOULD BE RECYCLED:");
    console.table([before[0], before[before.length - 1]].map(show));
  }
  if (kept.length) {
    console.log("Earliest rows that are KEPT:");
    console.table(kept.slice(0, 5).map(show));
  }
  window.__result = { before, kept };

  if (!DO_RECYCLE || before.length === 0) { console.log("DRY RUN (or nothing to do). Nothing was changed."); return; }

  // ---- RECYCLE ----
  const digest = await getDigest();
  if (!digest) { console.error("Could not get a request digest. Nothing was changed."); return; }
  const failures = [];
  let done = 0;
  for (const r of before) {
    const res = await post(base + "/items(" + r.Id + ")/recycle()", digest);
    if (!res.ok) failures.push(r.Id + " (HTTP " + res.status + ")");
    done++;
    if (done % 25 === 0) console.log("  " + done + "/" + before.length);
  }
  console.log("Done. Recycled " + (done - failures.length) + " of " + before.length + ". Kept " + kept.length + ".");
  if (failures.length) { console.warn(failures.length + " failed; see window.__failures. Run again to retry."); window.__failures = failures; }
})();

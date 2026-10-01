// Launch-day purge: removes ALL test data from the CRS site. DRY RUN unless the DO_ flags are true.
// Deletes: every file and folder in the 6 document libraries, every row in CRS Submissions,
// CRS Requests and CRS Audit Log, then (last, separately) empties the recycle bin.
// Keeps: CRS Config, Group Map, Folder Map, Term Abbreviation, groups, pages, the app.
// Run: open the CRS site as a site owner, F12 → Console, paste, Enter.
(async () => {
  // ---- CONFIG ----
  const EXPECTED_SITE = "/sites/CRS"; // refuses to run anywhere else
  const DO_LIBRARIES = false; // step 1: recycle all files + folders in the 6 libraries
  const DO_LISTS = false; // step 1: recycle all rows in CRS Submissions + CRS Requests
  const DO_AUDIT = true; // step 2: recycle all CRS Audit Log rows (run ~15 min after step 1)
  const DO_EMPTY_RECYCLE_BIN = true; // step 3: PERMANENT. Empties both recycle bin stages

  const LIB_SEGMENTS = [
    "ApprovalDocument",
    "Shared Documents",
    "HCApprovalDocument",
    "HCDocuments",
    "Archive",
    "HCArchive",
  ];
  const LISTS = {
    submissions: ["CRS Submissions", "DMS Submissions"],
    requests: ["CRS Requests", "DMS Requests"],
    audit: ["CRS Audit Log", "DMS Audit Log"],
  };

  // ---- HELPERS ----
  function resolveWeb() {
    try {
      if (window._spPageContextInfo)
        return window._spPageContextInfo.webAbsoluteUrl;
    } catch (e) {
      // not on a classic page
    }
    const match = location.pathname.match(/^\/(sites|teams)\/[^/]+/);
    return location.origin + (match ? match[0] : "");
  }
  const web = resolveWeb();
  const bust = () => "_=" + Date.now() + Math.random().toString(36).slice(2);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  async function getJson(path) {
    const url = web + path + (path.includes("?") ? "&" : "?") + bust();
    const res = await fetch(url, {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Cache-Control": "no-cache",
      },
    });
    const type = res.headers.get("content-type") || "";
    if (!res.ok || !type.includes("json"))
      return { ok: false, status: res.status, body: await res.text() };
    return { ok: true, data: await res.json() };
  }

  async function getAll(path) {
    const rows = [];
    let next = path;
    while (next) {
      const r = await getJson(next);
      if (!r.ok) return undefined;
      rows.push(...(r.data.value || []));
      next = r.data["odata.nextLink"]
        ? r.data["odata.nextLink"].replace(web, "")
        : "";
    }
    return rows;
  }

  let digest = "";
  async function getDigest() {
    const res = await fetch(web + "/_api/contextinfo", {
      method: "POST",
      headers: { Accept: "application/json;odata=nometadata" },
    });
    if (!res.ok)
      throw new Error(
        "Could not get a request digest (HTTP " + res.status + ")",
      );
    digest = (await res.json()).FormDigestValue;
  }

  // POST, retrying on throttling (429/503).
  async function post(path) {
    for (let attempt = 0; attempt < 4; attempt++) {
      const res = await fetch(web + path, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          "X-RequestDigest": digest,
        },
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

  function segmentOf(serverRelativeUrl) {
    const parts = (serverRelativeUrl || "").split("/").filter(Boolean);
    return decodeURIComponent(parts[parts.length - 1] || "");
  }

  // ---- SAFETY ----
  if (!web.toLowerCase().endsWith(EXPECTED_SITE.toLowerCase())) {
    console.error(
      "STOPPED: this is " +
        web +
        ", not " +
        EXPECTED_SITE +
        ". Nothing was changed.",
    );
    return;
  }
  console.log("Site:", web);
  console.log(
    "DO_LIBRARIES:",
    DO_LIBRARIES,
    "| DO_LISTS:",
    DO_LISTS,
    "| DO_AUDIT:",
    DO_AUDIT,
    "| DO_EMPTY_RECYCLE_BIN:",
    DO_EMPTY_RECYCLE_BIN,
  );

  // ---- FIND THE LIBRARIES AND LISTS ----
  const allLists = await getAll(
    "/_api/web/lists?$select=Id,Title,ItemCount,BaseTemplate,RootFolder/ServerRelativeUrl&$expand=RootFolder&$top=500",
  );
  if (allLists === undefined) {
    console.error("Could not read the site's lists. Nothing was changed.");
    return;
  }
  const libs = LIB_SEGMENTS.map((seg) => ({
    seg,
    list: allLists.find(
      (l) =>
        l.BaseTemplate === 101 &&
        segmentOf(
          l.RootFolder && l.RootFolder.ServerRelativeUrl,
        ).toLowerCase() === seg.toLowerCase(),
    ),
  }));
  const pickList = (titles) =>
    allLists.find((l) =>
      titles.some((t) => t.toLowerCase() === (l.Title || "").toLowerCase()),
    );
  const lists = {
    submissions: pickList(LISTS.submissions),
    requests: pickList(LISTS.requests),
    audit: pickList(LISTS.audit),
  };

  // ---- DRY-RUN REPORT ----
  const plan = [];
  for (const l of libs) {
    if (!l.list) {
      plan.push({
        what: "Library /" + l.seg,
        found: "NOT FOUND",
        items: "-",
        topFolders: "-",
        rootFiles: "-",
      });
      continue;
    }
    const folders = await getAll(
      "/_api/web/lists(guid'" +
        l.list.Id +
        "')/RootFolder/Folders?$select=Name,UniqueId,ItemCount&$top=500",
    );
    const files = await getAll(
      "/_api/web/lists(guid'" +
        l.list.Id +
        "')/RootFolder/Files?$select=Name,UniqueId&$top=500",
    );
    if (folders === undefined || files === undefined) {
      console.error(
        "Could not read the top of " + l.list.Title + ". Nothing was changed.",
      );
      return;
    }
    l.folders = folders.filter((f) => f.Name !== "Forms");
    l.files = files;
    plan.push({
      what: "Library " + l.list.Title,
      found: "/" + l.seg,
      items: l.list.ItemCount,
      topFolders: l.folders.map((f) => f.Name).join(", ") || "(none)",
      rootFiles: l.files.length,
    });
  }
  for (const key of Object.keys(lists)) {
    const l = lists[key];
    plan.push({
      what: "List " + (l ? l.Title : LISTS[key][0]),
      found: l ? "yes" : "NOT FOUND",
      items: l ? l.ItemCount : "-",
      topFolders: "-",
      rootFiles: "-",
    });
  }
  console.table(plan);
  window.__result = { plan, libs, lists };

  if (!DO_LIBRARIES && !DO_LISTS && !DO_AUDIT && !DO_EMPTY_RECYCLE_BIN) {
    console.log(
      "DRY RUN only. Nothing was changed. Check the table, then set the DO_ flags for step 1.",
    );
    return;
  }

  await getDigest();
  const failures = [];

  // ---- STEP 1a: LIBRARIES (recycle each top-level folder with everything under it, then root files) ----
  if (DO_LIBRARIES) {
    for (const l of libs) {
      if (!l.list) continue;
      for (const f of l.folders) {
        const res = await post(
          "/_api/web/GetFolderById('" + f.UniqueId + "')/recycle()",
        );
        console.log(
          (res.ok ? "✓" : "✗ HTTP " + res.status) +
            " folder " +
            l.list.Title +
            "/" +
            f.Name,
        );
        if (!res.ok) failures.push(l.list.Title + "/" + f.Name);
      }
      for (const f of l.files) {
        const res = await post(
          "/_api/web/GetFileById('" + f.UniqueId + "')/recycle()",
        );
        console.log(
          (res.ok ? "✓" : "✗ HTTP " + res.status) +
            " file " +
            l.list.Title +
            "/" +
            f.Name,
        );
        if (!res.ok) failures.push(l.list.Title + "/" + f.Name);
      }
    }
  }

  // ---- LIST ROWS (recycle every item) ----
  async function purgeList(l, label) {
    if (!l) {
      console.warn(label + " not found, skipped.");
      return;
    }
    const rows = await getAll(
      "/_api/web/lists(guid'" + l.Id + "')/items?$select=Id&$top=500",
    );
    if (rows === undefined) {
      console.error("Could not read " + l.Title + ". It was not purged.");
      failures.push(l.Title + " (read failed)");
      return;
    }
    console.log(l.Title + ": " + rows.length + " rows to recycle");
    let done = 0;
    for (const r of rows) {
      const res = await post(
        "/_api/web/lists(guid'" + l.Id + "')/items(" + r.Id + ")/recycle()",
      );
      if (!res.ok)
        failures.push(l.Title + " row " + r.Id + " (HTTP " + res.status + ")");
      done++;
      if (done % 25 === 0)
        console.log("  " + l.Title + ": " + done + "/" + rows.length);
    }
    console.log("✓ " + l.Title + ": " + done + " rows processed");
  }

  if (DO_LISTS) {
    await purgeList(lists.submissions, "CRS Submissions");
    await purgeList(lists.requests, "CRS Requests");
  }

  // ---- STEP 2: AUDIT LOG (last, after the deletion-audit flows have finished writing) ----
  if (DO_AUDIT) await purgeList(lists.audit, "CRS Audit Log");

  // ---- STEP 3: EMPTY THE RECYCLE BIN (PERMANENT) ----
  if (DO_EMPTY_RECYCLE_BIN) {
    const first = await post("/_api/web/RecycleBin/DeleteAll()");
    console.log(
      (first.ok ? "✓" : "✗ HTTP " + first.status) +
        " first-stage recycle bin emptied",
    );
    const second = await post(
      "/_api/site/RecycleBin/DeleteAllSecondStageItems()",
    );
    console.log(
      (second.ok ? "✓" : "✗ HTTP " + second.status) +
        " second-stage recycle bin emptied",
    );
    if (!first.ok || !second.ok) failures.push("recycle bin");
  }

  // ---- CHECK WHAT IS LEFT ----
  const after = await getAll(
    "/_api/web/lists?$select=Id,Title,ItemCount&$top=500",
  );
  if (after !== undefined) {
    const ids = [
      ...libs.map((l) => l.list && l.list.Id),
      ...Object.values(lists).map((l) => l && l.Id),
    ].filter(Boolean);
    console.table(
      after
        .filter((l) => ids.includes(l.Id))
        .map((l) => ({ list: l.Title, itemsLeft: l.ItemCount })),
    );
  }
  if (failures.length > 0) {
    console.warn(
      failures.length +
        " item(s) failed. They are on window.__failures. Run the same step again to retry.",
    );
    window.__failures = failures;
  } else {
    console.log("Done, no failures.");
  }
})();

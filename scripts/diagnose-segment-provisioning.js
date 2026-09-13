/*
 * diagnose-segment-provisioning.js — READ ONLY. Browser console, run as a SITE ADMINISTRATOR.
 *
 * WHY THIS EXISTS
 * ----------------
 * One segment's upload-form dropdown went completely empty right after a Folder Reconciliation run,
 * while other segments still work — so this is NOT the whole-library ACL wipeout
 * `diagnose-blank-libraries.js` checks for. It is per-segment provisioning.
 *
 * The upload form offers a unit ONLY when BOTH of these hold (shared/segmentReadiness.ts):
 *   1. Its leaf term has a `CRS Folder Map` row (a term with none is dropped — no folder, no path).
 *   2. A live AddListItems probe against that folder's ACL succeeds.
 * If EVERY unit under a segment fails either check, the whole segment silently empties itself with
 * no error anywhere — which is exactly the reported symptom.
 *
 * This reconstructs both checks for one segment, admin-side (admins issue no probes themselves, so
 * this reads the raw ACL instead and reports what an ordinary uploader's group would find there).
 * It ALSO tracks whether the term-tree walk itself completed cleanly — a transient term-store
 * failure mid-walk (503s) can make live units look unprovisioned when nothing is actually wrong with
 * the data underneath (see the 2026-08-20 "503 made six live units strays" incident this project has
 * already hit once).
 *
 * HOW TO RUN
 *   1. Set SEGMENT_KEY below to the mode row's Title (e.g. "mode_project_cars") OR its StagingFolder
 *      (e.g. "PCAR") — either is matched.
 *   2. Open the site in a browser, signed in as a SITE ADMINISTRATOR.
 *   3. F12 -> Console -> paste this whole file -> Enter.
 *   4. Read the summary table and the checklist at the bottom.
 *
 * ⚠ EVERY REQUEST IS A GET. It creates nothing, changes nothing, deletes nothing.
 * ⚠ `unknown`/`null` MEANS "COULD NOT READ", NEVER "NOT THERE" — the two are reported separately
 *   throughout, because they point at opposite fixes.
 */
(async () => {
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const SEGMENT_KEY = "PCAR"; // fill in: the mode row's Title OR its StagingFolder, e.g. "PCAR"

  const derived = (() => {
    const m = location.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    return location.origin + (m ? m[1] : "");
  })();
  const web = (SITE_OVERRIDE || derived).replace(/\/+$/, "");
  console.log("reading site:", web || "(root)");

  if (!SEGMENT_KEY.trim()) {
    console.log(
      "Set SEGMENT_KEY at the top of this script to the segment's Title or StagingFolder first.",
    );
    return;
  }

  const get = async (url) => {
    try {
      const r = await fetch(web + url, {
        headers: { Accept: "application/json;odata=nometadata" },
      });
      if (!r.ok) return { __error: r.status };
      return await r.json();
    } catch (e) {
      return { __error: String(e && e.message ? e.message : e) };
    }
  };

  // Same discovery pattern as every other script here — this client renames lists.
  const resolveList = async (suffix) => {
    for (const prefix of ["CRS", "DMS"]) {
      const title = `${prefix} ${suffix}`;
      const r = await get(
        `/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Title`,
      );
      if (r.__error === undefined) return title;
    }
    return undefined;
  };

  const configList = await resolveList("Config");
  const folderMapList = await resolveList("Folder Map");
  const abbrevList = await resolveList("Term Abbreviation");
  console.log(
    "Config:",
    configList || "NOT FOUND",
    "| Folder Map:",
    folderMapList || "NOT FOUND",
    "| Abbreviations:",
    abbrevList || "NOT FOUND",
  );
  if (!configList || !folderMapList) {
    console.log(
      "Could not resolve the lists this needs — nothing else to check.",
    );
    return;
  }

  // ── The mode row ────────────────────────────────────────────────────────────
  const modes = await get(
    `/_api/web/lists/getbytitle('${encodeURIComponent(configList)}')/items` +
      `?$select=Id,Title,ModeLabel,StagingFolder,TermSetGuid,Levels,PendingLevels&$filter=ConfigType eq 'mode'&$top=200`,
  );
  if (modes.__error !== undefined) {
    console.log(`Could not read ${configList}: HTTP ${modes.__error}`);
    return;
  }
  const norm = (v) => (v || "").trim().toLowerCase();
  const seg = (modes.value || []).find(
    (m) =>
      norm(m.Title) === norm(SEGMENT_KEY) ||
      norm(m.StagingFolder) === norm(SEGMENT_KEY),
  );
  if (!seg) {
    console.log(
      `No mode row matched "${SEGMENT_KEY}". Rows found:`,
      (modes.value || []).map((m) => m.Title),
    );
    return;
  }
  console.log("Matched segment:", {
    Title: seg.Title,
    ModeLabel: seg.ModeLabel,
    StagingFolder: seg.StagingFolder,
  });
  console.log(
    "Has PendingLevels staged:",
    (seg.PendingLevels || "").trim().length > 0
      ? "YES — a below-Unit structure change is staged and not yet live (unrelated to Dept/Unit provisioning below)"
      : "no",
  );

  const setGuid = (seg.TermSetGuid || "").trim();
  if (!setGuid) {
    console.log("This mode row has no TermSetGuid — nothing to walk.");
    return;
  }

  // ── Walk the PERMISSIONED tiers only (Department -> Unit): every segment family in this project
  //    is exactly two permissioned levels, so depth 2 is the constant, not a guess. Deeper terms
  //    (SubUnit, shared-folder values) are below-Unit and never provisioned by reconciliation. ────
  const MAX_REQUESTS = 400;
  let requests = 0;
  let incomplete = false;

  const children = async (termId) => {
    if (requests >= MAX_REQUESTS) {
      incomplete = true;
      return [];
    }
    requests++;
    const url = termId
      ? `/_api/v2.1/termStore/sets/${setGuid}/terms/${termId}/children?$select=id,labels`
      : `/_api/v2.1/termStore/sets/${setGuid}/children?$select=id,labels`;
    const res = await get(url);
    if (res.__error !== undefined) {
      incomplete = true;
      return [];
    }
    return res.value || [];
  };

  const labelOf = (t) =>
    (t.labels && t.labels[0] && t.labels[0].name) || "(no label)";

  const depts = await children();
  const units = [];
  for (const d of depts) {
    const kids = await children(d.id);
    for (const u of kids)
      units.push({ deptLabel: labelOf(d), unitLabel: labelOf(u), id: u.id });
  }
  console.log(
    `Term tree: ${depts.length} department(s), ${units.length} unit(s) total.`,
  );
  if (incomplete) {
    console.log(
      "⚠ THE TERM-TREE WALK DID NOT COMPLETE CLEANLY (a request failed or the request cap was hit). " +
        "Units below may be under-reported. Re-run this script before trusting a unit reported as " +
        "'no Folder Map row' — that could mean the walk simply never reached it.",
    );
  }

  // ── Folder Map: every row, keyed both ways ─────────────────────────────────
  const fmRows = await get(
    `/_api/web/lists/getbytitle('${encodeURIComponent(folderMapList)}')/items` +
      `?$select=Id,FolderUniqueId,TermGuid,Section&$top=5000`,
  );
  let fmUnreadable = false;
  const byTermGuid = {};
  if (fmRows.__error !== undefined) {
    fmUnreadable = true;
    console.log(
      `⚠ Could not read ${folderMapList}: HTTP ${fmRows.__error} — every unit below will read UNKNOWN.`,
    );
  } else {
    for (const r of fmRows.value || []) {
      const g = norm(r.TermGuid);
      if (g) byTermGuid[g] = r;
    }
    console.log(`Folder Map: ${fmRows.value.length} row(s) total on the list.`);
  }

  // ── Per unit: does its term have a Folder Map row, does the folder resolve, does its ACL carry
  //    a real (non-Limited-Access) role? ─────────────────────────────────────
  const results = [];
  for (const u of units) {
    const row = fmUnreadable ? undefined : byTermGuid[norm(u.id)];
    if (fmUnreadable) {
      results.push({ ...u, state: "UNKNOWN — Folder Map unreadable" });
      continue;
    }
    if (!row) {
      results.push({ ...u, state: "NO FOLDER MAP ROW" });
      continue;
    }
    const folderRes = await get(
      `/_api/web/GetFolderById('${row.FolderUniqueId}')` +
        `?$select=ServerRelativeUrl&$expand=ListItemAllFields/RoleAssignments,ListItemAllFields/RoleAssignments/Member,ListItemAllFields/RoleAssignments/RoleDefinitionBindings` +
        `&$select=ServerRelativeUrl,ListItemAllFields/RoleAssignments/Member/Title,ListItemAllFields/RoleAssignments/RoleDefinitionBindings/Name`,
    );
    if (folderRes.__error !== undefined) {
      results.push({
        ...u,
        state: `FOLDER DOES NOT RESOLVE (HTTP ${folderRes.__error}) — Folder Map row points at a dead UniqueId`,
      });
      continue;
    }
    const assignments =
      (folderRes.ListItemAllFields &&
        folderRes.ListItemAllFields.RoleAssignments) ||
      [];
    const realGrants = assignments
      .map((a) => ({
        principal: a.Member ? a.Member.Title : "(unknown)",
        roles: (a.RoleDefinitionBindings || []).map((b) => b.Name),
      }))
      .filter((a) => a.roles.some((r) => r !== "Limited Access"));
    results.push({
      ...u,
      state:
        realGrants.length > 0
          ? "OK"
          : "FOLDER RESOLVES BUT NO REAL GRANT — reconciliation has not (re)granted it",
      url: folderRes.ServerRelativeUrl,
      grants: realGrants,
    });
  }

  console.log(`\n── ${SEGMENT_KEY} — ${results.length} unit(s) checked ──\n`);
  console.table(
    results.map((r) => ({
      Department: r.deptLabel,
      Unit: r.unitLabel,
      State: r.state,
      Path: r.url || "",
      Grants: (r.grants || [])
        .map((g) => `${g.principal}:${g.roles.join("+")}`)
        .join(", "),
    })),
  );

  // console.table truncates long cells with "…" — the Grants column is exactly where that bites,
  // since a group name plus its role list is long. Dump it untruncated too, same run, no scope
  // issue (results only lives inside this IIFE, so it cannot be read back afterward).
  console.log("\n── grants, untruncated ──");
  console.log(
    JSON.stringify(
      results.map((r) => ({
        unit: `${r.deptLabel}/${r.unitLabel}`,
        state: r.state,
        grants: (r.grants || []).map((g) => `${g.principal}:${g.roles.join("+")}`),
      })),
      null,
      2,
    ),
  );

  const bad = results.filter((r) => r.state !== "OK");
  console.log(
    `\n${bad.length} of ${results.length} unit(s) would NOT be offered in the upload form.`,
  );
  console.log(`
── READ IT AGAINST THIS ──
- "NO FOLDER MAP ROW" on every unit -> reconciliation has not written rows for this segment at all
  (never run for it, or errored before reaching it). Re-run Folder Reconciliation, scoped to this
  segment, and re-run this script afterward.
- "FOLDER DOES NOT RESOLVE" -> the row exists but points at a UniqueId that no longer resolves
  (the folder was deleted/recycled since). This is the orphan-repair pass's job; a full
  reconciliation run should report and clean these up.
- "FOLDER RESOLVES BUT NO REAL GRANT" -> the folder is real and the Folder Map row is correct, but
  no uploader/approver group holds anything beyond Limited Access on it — reconciliation created the
  folder but the grant step did not complete (or was skipped by scope). Re-run reconciliation.
- If EVERY unit shows one of the above, that is the direct cause of the empty dropdown: the
  provisioned-segment filter in the upload form legitimately has nothing left to offer.
- If the term walk above reported incomplete, re-run this script once before trusting a "NO FOLDER
  MAP ROW" verdict — it may just mean the walk did not reach that unit.
`);
})();

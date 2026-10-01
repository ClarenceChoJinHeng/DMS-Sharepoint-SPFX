/*
 * Companion to `check-stale-delete-roles.js` — exports the SAFETY NET the established repair path
 * requires before any of those groups are touched: every stale group's CURRENT members, as a CSV.
 * -----------------------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, renamed or deleted. It only DOWNLOADS a
 * file — the one side effect a browser console script can have with no write to SharePoint at all.
 *
 * Run it in the browser console, ON THE CRS SITE, after `check-stale-delete-roles.js` has already
 * told you how many groups are affected. Live tenant run, 2026-09-21: 393 of 857 mapped groups,
 * across GHO / MHO / TNS — far too many to review one at a time in a chat window, which is exactly
 * why this exists: it turns that list into one CSV, opens in Excel, sorted by segment.
 *
 * WHY THIS MATTERS BEFORE DELETING ANYTHING: CLAUDE.md's own established repair path for a stale
 * group is "note the group's members (export first — deleting the group loses them), DELETE the
 * group, re-create it via Group Management's persona picker, re-add the members, re-run Folder
 * Reconciliation." Deleting a SharePoint group destroys its membership permanently — there is no
 * undo, and the whole point of recreating it is a FRESH group with a FRESH id, so the members are
 * not recoverable from the group afterward. This CSV is that "note the members" step, done for
 * every affected group in one pass instead of by hand, one at a time, across 393 groups.
 *
 * THE FASTER BULK PATH THIS ENABLES, rather than one group at a time:
 *   1. Open this CSV, filter to one segment (GHO / MHO / TNS).
 *   2. On Group Management, filter the group list to that segment's stale groups and use
 *      "Select all N shown" -> "Delete N selected" (typed DELETE confirmation) to remove them all
 *      in one sequential run — this ALREADY EXISTS in the app (built 2026-08-25).
 *   3. Run Bulk Provisioning ("Create missing groups") for that segment. Since the stale groups
 *      are now genuinely gone, it creates them fresh with the CURRENT persona role set — no DEL/
 *      DELS/DELHC/DELSHC — and maps their Group Map rows.
 *   4. Re-add each group's members from this CSV (Group Management's own member editor).
 *   5. Re-run Folder Reconciliation for the segment.
 *
 * WHAT IT DOES:
 *   1. Re-runs the exact same stale-role detection as `check-stale-delete-roles.js` (CRS Config for
 *      segment codes, CRS Group Map rolled up per GroupId, live `sitegroups` for titles).
 *   2. Reads every group's CURRENT members via the proven keyset-paged `sitegroups?$expand=Users`
 *      method (`fetchAllGroupMembers` in `shared/spGroups.ts`) — never the slow, throttle-prone
 *      one-request-per-group approach this project has already been bitten by twice.
 *   3. Writes one CSV row per (stale group, member) — or one row with blank member columns for a
 *      group that currently has nobody in it, so an empty group is never silently dropped from the
 *      export. Sorted by segment, then group name.
 *   4. Triggers a browser download of the CSV. Nothing else happens.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  /* Leave blank to auto-detect. Set it if the auto-detected site below is wrong, e.g.
     "https://sdguthrie.sharepoint.com/sites/CRS" */
  const SITE_OVERRIDE = "";

  /* ── Resolve the site ────────────────────────────────────────────────────
     Same three-route fallback every script in this project uses. */
  const resolveWeb = () => {
    if (SITE_OVERRIDE) return SITE_OVERRIDE.replace(/\/+$/, "");
    try {
      if (
        typeof _spPageContextInfo !== "undefined" &&
        _spPageContextInfo.webAbsoluteUrl
      ) {
        return _spPageContextInfo.webAbsoluteUrl.replace(/\/+$/, "");
      }
    } catch (e) {
      /* not defined on this page — fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
    return (location.origin + (m ? m[1] : "")).replace(/\/+$/, "");
  };
  const web = resolveWeb();
  console.log("Site:", web);

  /* ⚠ CACHE-BUSTED. See `check-stale-delete-roles.js`'s own note — a plain GET can be answered from
     the browser's own HTTP cache, and a unique query parameter is the one thing that cannot match a
     prior response. */
  const bust = (u) =>
    (u.indexOf("?") === -1 ? "?" : "&") +
    `_=${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  const get = async (url) => {
    const r = await fetch(url + bust(url), {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
      },
      cache: "no-store",
    });
    const type = r.headers.get("content-type") || "";
    if (type.indexOf("json") === -1) {
      throw new Error(
        `Expected JSON and got "${type}" from ${url}\n` +
          `The site resolved to ${web}. If that is wrong, set SITE_OVERRIDE at the top and re-run.`,
      );
    }
    if (!r.ok) throw new Error(`HTTP ${r.status} from ${url}`);
    return r.json();
  };

  const resolveList = async (candidates) => {
    for (const t of candidates) {
      try {
        const probe = await get(
          `${web}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=Title,ItemCount`,
        );
        if (probe && probe.Title) return probe;
      } catch (e) {
        /* try the next candidate */
      }
    }
    throw new Error(`None of these list titles resolved: ${candidates.join(", ")}`);
  };

  /* ── Role normalisation — identical to `check-stale-delete-roles.js` ──────────────────────── */
  const SHORT_ROLE_CODES = [
    "MEMBER", "UPL", "APR", "DEL", "DELS", "SEGVIEW", "UPLHC", "APRHC",
    "DELSHC", "DELHC", "SHAREHC", "MEMBERHC", "DEPTVIEW", "SHARE", "GLOBAL", "ENTRY",
  ];
  const ROLE_ALIASES = {
    UPLOADER: "UPL",
    APPROVER: "APR",
    DELETER_DOCUMENTS: "DEL",
    DELETER_STAGING: "DELS",
    MEMBER: "MEMBER",
    VIEWER: "MEMBER",
    SEGMENTVIEW: "SEGVIEW",
    DEPARTMENTVIEW: "DEPTVIEW",
    VIEWER_HC: "MEMBERHC",
  };
  const normalizeRoleValue = (raw) => {
    const v = (raw ?? "").trim().toUpperCase().replace(/[\s-]+/g, "_");
    if (v.length === 0) return "";
    if (SHORT_ROLE_CODES.indexOf(v) !== -1) return v;
    return ROLE_ALIASES[v] ?? v;
  };
  const STALE_DELETE_ROLES = ["DEL", "DELS", "DELHC", "DELSHC"];

  /* ── 1. Segments — code from every "mode" row, for labelling only ─────────────────── */
  const configList = await resolveList(["CRS Config", "DMS Config"]);
  const modeRes = await get(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(configList.Title)}')/items` +
      `?$select=Id,ModeLabel,StagingFolder,TermSetGuid,SortOrder,Levels,ConfigType` +
      `&$filter=ConfigType eq 'mode'&$orderby=SortOrder&$top=5000`,
  );
  const allModeRows = modeRes.value || [];
  const codeByGuid = {};
  allModeRows
    .filter((r) => (r.TermSetGuid || "").trim() && (r.Levels || "").trim())
    .forEach((r) => {
      const g = (r.TermSetGuid || "").trim().toLowerCase();
      if (codeByGuid[g] !== undefined) return; // first match (lowest SortOrder) wins
      codeByGuid[g] = (r.StagingFolder || "").trim();
    });
  console.log(`Segments resolved: ${Object.keys(codeByGuid).length}`);

  /* ── 2. Group Map rows, rolled up per GroupId ──────────────────────────────────────── */
  const gmList = await resolveList(["CRS Group Map", "DMS Group Map"]);
  const gmRes = await get(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(gmList.Title)}')/items` +
      `?$select=GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
  );
  const gmRows = gmRes.value || [];
  if (gmRows.length >= 5000) {
    console.warn("⚠ CRS Group Map read is AT THE CAP (5000) — some rows may be missing. Page it by hand.");
  }
  const byGroup = {}; // GroupId -> { roles: Set, segments: Set }
  gmRows.forEach((r) => {
    const gid = (r.GroupId ?? "").trim();
    const role = normalizeRoleValue(r.Role ?? "");
    if (!gid || !role) return;
    const entry = byGroup[gid] || (byGroup[gid] = { roles: new Set(), segments: new Set() });
    entry.roles.add(role);
    const seg = (r.Segment ?? "").trim().toLowerCase();
    if (seg) entry.segments.add(seg);
  });
  console.log(`Mapped groups checked: ${Object.keys(byGroup).length}`);

  /* ── 3. Live SharePoint group titles ───────────────────────────────────────────────── */
  const titleRes = await get(`${web}/_api/web/sitegroups?$select=Id,Title&$top=5000`);
  const titleById = {};
  (titleRes.value || []).forEach((g) => {
    titleById[String(g.Id)] = g.Title;
  });

  /* ── 4. Which groups are stale ─────────────────────────────────────────────────────── */
  const staleGroups = [];
  Object.keys(byGroup).forEach((gid) => {
    const entry = byGroup[gid];
    const roles = Array.from(entry.roles);
    const staleFound = roles.filter((r) => STALE_DELETE_ROLES.indexOf(r) !== -1);
    if (staleFound.length === 0) return;
    const title = titleById[gid];
    const segCodes = Array.from(entry.segments)
      .map((g) => codeByGuid[g] || g)
      .filter((v, i, a) => a.indexOf(v) === i);
    staleGroups.push({
      gid,
      title: title !== undefined ? title : "",
      gone: title === undefined,
      segCodes,
      roles: roles.sort(),
      staleFound: staleFound.sort(),
    });
  });
  console.log(`Stale groups found: ${staleGroups.length}`);
  if (staleGroups.length === 0) {
    console.log("✅ Nothing to export — no group carries a stale delete role. Run check-stale-delete-roles.js again to confirm.");
    return;
  }
  const alreadyGone = staleGroups.filter((g) => g.gone);
  if (alreadyGone.length > 0) {
    console.warn(
      `⚠ ${alreadyGone.length} of the stale groups have ALREADY BEEN DELETED (their Group Map rows ` +
        `are orphaned, no live title to resolve) — nobody to export members for. These rows must be ` +
        `cleaned up separately; they are still listed in the CSV below with blank member columns so ` +
        `nothing is silently dropped from the count.`,
    );
  }

  /* ── 5. Every group's current members — the keyset-paged `$expand=Users` method, never the
     one-request-per-group approach (an 8-minute load across ~700 groups on this exact site,
     2026-09-02 — see `fetchAllGroupMembers` in shared/spGroups.ts, replicated here verbatim). ──── */
  console.log("Reading group memberships (paged)…");
  const membersById = {};
  {
    const PAGE = 500;
    let afterId = 0;
    let pages = 0;
    let more = true;
    while (more && pages < 20) {
      pages += 1;
      const url =
        `${web}/_api/web/sitegroups?$select=Id,Users/Id,Users/Title,Users/Email,Users/LoginName` +
        `&$expand=Users&$orderby=Id&$filter=Id gt ${afterId}&$top=${PAGE}`;
      const data = await get(url);
      const rows = data.value || [];
      rows.forEach((g) => {
        membersById[g.Id] = (g.Users || []).map((u) => ({
          title: u.Title || "",
          email: u.Email || "",
          loginName: u.LoginName || "",
        }));
        if (g.Id > afterId) afterId = g.Id;
      });
      more = rows.length === PAGE;
    }
    if (more) {
      throw new Error(
        `Member read did not finish after ${pages} pages of ${PAGE} — refusing a partial export. ` +
          `Raise the page cap in this script and re-run.`,
      );
    }
  }
  console.log(`Membership read across ${Object.keys(membersById).length} group(s) on the site.`);

  /* ── 6. Build the CSV — same escaping/CRLF/BOM rules as `shared/groupExportCsv.ts` ─────────── */
  const csvCell = (value) => {
    const v = value ?? "";
    const risky = /^[=+\-@\t\r]/.test(v) ? `'${v}` : v;
    return /[",\r\n]/.test(risky) ? `"${risky.replace(/"/g, '""')}"` : risky;
  };
  const HEADERS = [
    "Segment",
    "Group Id",
    "Group",
    "Full Role Set",
    "Stale Delete Role(s)",
    "Member Name",
    "Member Email",
    "Member Login Name",
  ];
  const lines = [HEADERS.map(csvCell).join(",")];

  staleGroups
    .slice()
    .sort(
      (a, b) =>
        (a.segCodes[0] || "").localeCompare(b.segCodes[0] || "") ||
        (a.title || `(id ${a.gid})`).localeCompare(b.title || `(id ${b.gid})`),
    )
    .forEach((g) => {
      const seg = g.segCodes.length ? g.segCodes.join("/") : "(no segment)";
      const label = g.gone ? `(id ${g.gid} — group already deleted)` : g.title;
      const members = g.gone ? [] : membersById[Number(g.gid)] || [];
      if (members.length === 0) {
        lines.push(
          [seg, g.gid, label, g.roles.join("; "), g.staleFound.join("; "), "", "", ""]
            .map(csvCell)
            .join(","),
        );
      } else {
        members.forEach((m) => {
          lines.push(
            [
              seg,
              g.gid,
              label,
              g.roles.join("; "),
              g.staleFound.join("; "),
              m.title,
              m.email,
              m.loginName,
            ]
              .map(csvCell)
              .join(","),
          );
        });
      }
    });

  const csv = lines.join("\r\n");

  /* ── 7. Download ────────────────────────────────────────────────────────────────────────────
     BOM-prefixed: without it Excel reads the file as ANSI and mangles anything non-ASCII (the
     fullwidth ＆ in unit names, per `shared/groupExportCsv.ts`'s own comment). Revoke on the next
     tick rather than synchronously — revoking immediately can cancel the download in Edge. */
  const d = new Date();
  const pad = (n) => (n < 10 ? `0${n}` : String(n));
  const fileName = `CRS-stale-delete-groups-${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}.csv`;
  const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 0);

  console.log(`\n✅ Downloaded "${fileName}" — ${lines.length - 1} row(s) across ${staleGroups.length} stale group(s).`);
  console.log(
    "Open it, filter to one segment at a time, and work through the bulk-delete → bulk-provision → " +
      "re-add-members → reconcile cycle described at the top of this script.",
  );
})();

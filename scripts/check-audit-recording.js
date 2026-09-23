/*
 * Is the Audit Log actually recording new uploads?
 * -------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, deleted or granted.
 *
 * Run it in the browser console, ON THE CRS SITE, signed in as an ADMINISTRATOR (it reads the
 * whole `CRS Audit Log` list, which is Owners/service-account write-only but readable by an admin).
 *
 * WHY THIS EXISTS: reported live, 2026-09-19, again 2026-09-23 ("Audit log is not tracking
 * upload for bulk upload for some reason") — while testing Bulk Upload both times.
 * `Uploaded`/`Approved`/`Rejected`/`Routed`/`Replaced` rows are written EXCLUSIVELY by
 * Power Automate flows watching the approval libraries (`Audit — approval activity` + its HC
 * clone) and by `Auto-route`/`HC Auto Route` — never by client code, because that list restricts
 * writes to Owners and the service account by design (tamper resistance). So a genuinely missing
 * row means one of those flows is not running, not that anything in this repo's TypeScript is at
 * fault — Bulk Upload's own tagging gap (see check-tagging-status.js) is a SEPARATE, already-known
 * issue and does not explain a missing `Uploaded` row, which fires on library CREATE regardless of
 * whether tagging ever completes.
 *
 * WHAT THIS CHECKS, IN ORDER:
 *   1. Can the `CRS Audit Log` list be found at all, and when is its NEWEST row? A long gap since
 *      "now" (rather than just since a moment ago) is the tell that a flow has stopped altogether,
 *      not that one specific upload was missed.
 *   2. For each approval library (normal + HC), every item created in the last WINDOW_MINUTES —
 *      does an `Uploaded` audit row exist for its UniqueId? This is the direct check: a real file
 *      sitting in the library with no matching row is the gap itself, independent of tagging.
 *
 * KNOWN HISTORICAL CAUSES OF EXACTLY THIS SYMPTOM IN THIS PROJECT (check these BEFORE assuming a
 * new code defect — none of them are fixable from this repo):
 *   - A library was renamed and the flow's own `getbytitle('...')` reference is now stale — this
 *     has happened repeatedly (HC libraries, "Documents" -> "Restricted & Confidential Document").
 *     Compare the title this script reports below against what the flow's trigger/actions say.
 *   - The flow is simply turned OFF, or its connection needs re-authorising.
 *   - A service-account migration is mid-flight (crs@sdguthrie.com -> gdc@sdguthrie.com per the
 *     2026-09-19 runbook) and the flow's connection was disconnected without being reconnected yet.
 *   - The flow's dedupe check ("Already logged") is reading a poisoned blank-ItemUniqueId row and
 *     answering "already logged" for everything — see the 2026-08-23 CLAUDE.md entry on this exact
 *     failure mode.
 * None of these are things this script can fix; it can only tell you which one to go look at.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  // Leave blank to auto-detect. Set it if the auto-detected site below is wrong.
  const SITE_OVERRIDE = "";

  // How far back to look for library items to cross-check against the audit log.
  const WINDOW_MINUTES = 180;

  /* ── Resolve the site (same pattern as every other script here — reading a page path alone
     breaks on a library page whose URL has extra segments after the site, e.g. .../Forms/...) ── */
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

  const get = async (url) => {
    const r = await fetch(url, {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
      },
    });
    const type = r.headers.get("content-type") || "";
    if (type.indexOf("json") === -1) {
      const text = await r.text();
      throw new Error(
        `Expected JSON and got "${type}" (HTTP ${r.status}) from ${url}\n` +
          `The site resolved to ${web} — if that is wrong, set SITE_OVERRIDE at the top.\n` +
          `First 200 chars of response: ${text.slice(0, 200)}`,
      );
    }
    if (!r.ok) throw new Error(`HTTP ${r.status} from ${url}`);
    return r.json();
  };
  const tryGet = async (url) => {
    try {
      return await get(url);
    } catch (e) {
      return { __error: e.message };
    }
  };

  const findList = async (candidates, select) => {
    for (const t of candidates) {
      const probe = await tryGet(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=${select}`,
      );
      if (probe && probe.Title) return probe;
    }
    return undefined;
  };

  /* ── 1. Resolve `CRS Audit Log` and report its newest row ────────────────────────────────── */
  const auditCandidates = ["CRS Audit Log", "GDC Audit Log", "DMS Audit Log"];
  const auditList = await findList(auditCandidates, "Title,ItemCount");
  if (!auditList) {
    console.log(
      "%c✗ Could not find the audit log list under any known name — checked: " +
        auditCandidates.join(", "),
      "color:#a4262c;font-weight:bold",
    );
    console.log("  If it has a different title on this site, add it to auditCandidates above.");
    return;
  }
  console.log(`Audit log: "${auditList.Title}" (${auditList.ItemCount} items)`);

  const newest = await tryGet(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(auditList.Title)}')/items` +
      `?$select=EventTime,EventType,Source,ActorEmail&$orderby=EventTime desc&$top=5`,
  );
  if (newest.__error) {
    console.log("✗ Could not read the audit log's newest rows:", newest.__error);
  } else {
    console.log("\nMost recent 5 audit rows (any type):");
    for (const row of newest.value || []) {
      const ageMin = row.EventTime
        ? Math.round((Date.now() - new Date(row.EventTime).getTime()) / 60000)
        : "?";
      console.log(
        `  ${row.EventTime}  (${ageMin} min ago)  ${row.EventType}  ${row.Source || ""}  ${row.ActorEmail || ""}`,
      );
    }
    if ((newest.value || []).length === 0) {
      console.log("  (list is completely empty)");
    } else {
      const gapMin = Math.round(
        (Date.now() - new Date(newest.value[0].EventTime).getTime()) / 60000,
      );
      if (gapMin > WINDOW_MINUTES) {
        console.log(
          `\n%c⚠ THE NEWEST ROW IN THE WHOLE LIST IS ${gapMin} MINUTES OLD. That points at a flow that ` +
            `has stopped running entirely (turned off, disconnected, or mid-migration), not at one ` +
            `missed upload — check Power Automate's run history for "Audit — approval activity" before ` +
            `anything else.`,
          "color:#a4262c;font-weight:bold",
        );
      }
    }
  }

  /* ── 2. Resolve the approval libraries and cross-check recent items ──────────────────────── */
  const LIB_CANDIDATES = [
    ["Approval for Document", "Approval Document"],
    ["Approval for Highly Confidential Document", "HC Approval Document"],
  ];

  const sinceIso = new Date(Date.now() - WINDOW_MINUTES * 60000).toISOString();

  for (const candidates of LIB_CANDIDATES) {
    const lib = await findList(candidates, "Title,ItemCount");
    if (!lib) {
      console.log(
        `\n⚠ Could not find a library under any of: ${candidates.join(", ")} — skipping.`,
      );
      continue;
    }
    console.log(`\n=== "${lib.Title}" — items created in the last ${WINDOW_MINUTES} min ===`);

    const items = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(lib.Title)}')/items` +
        // `BulkImport` added 2026-09-23 — the marker Bulk Upload alone stamps (see
        // "BULK UPLOAD IS AN UPLOADER TOOL, THROUGH THE APPROVAL LIBRARY" in CLAUDE.md). Reading
        // it here is what lets a "missing" file below be labelled BULK vs FORM, which is the
        // actual question ("why is it bulk specifically") rather than just "is anything missing".
        `?$select=Id,FileLeafRef,FSObjType,Created,GUID,BulkImport` +
        `&$filter=FSObjType eq 0 and Created ge datetime'${sinceIso}'` +
        `&$orderby=Created desc&$top=200`,
    );
    if (items.__error) {
      console.log(`  ✗ Could not read this library: ${items.__error}`);
      continue;
    }
    const files = items.value || [];
    console.log(`  ${files.length} file(s) created in the window.`);
    if (files.length === 0) continue;

    // GUID (the list item's own GUID field) is NOT the file's UniqueId — resolve each file's real
    // UniqueId via the File resource, same as the app itself does everywhere (gotcha: item ids and
    // GUIDs are per-list/per-item; the audit rows are keyed on the FILE's UniqueId, not this GUID).
    const auditHit = async (uniqueId, eventType) => {
      const res = await tryGet(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(auditList.Title)}')/items` +
          `?$select=Id&$filter=ItemUniqueId eq '${uniqueId}' and EventType eq '${eventType}'&$top=1`,
      );
      if (!res || res.__error) return undefined; // could not ask — never read this as "absent"
      return res.value && res.value.length > 0;
    };

    let matched = 0;
    const missing = [];
    for (const f of files) {
      const fileInfo = await tryGet(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(lib.Title)}')/items(${f.Id})` +
          `/File?$select=UniqueId`,
      );
      const uniqueId = fileInfo && !fileInfo.__error ? fileInfo.UniqueId : undefined;
      if (!uniqueId) {
        missing.push({ ...f, reason: "could not resolve the file's UniqueId" });
        continue;
      }
      const hasUploaded = await auditHit(uniqueId, "Uploaded");
      if (hasUploaded === true) {
        matched += 1;
        continue;
      }
      /* ⚠ FOR A BULK-IMPORTED FILE ONLY, ALSO CHECK "Approved"/"Routed" — this is the theory named
         at the top of this file: `CRS — Auto-approve bulk imports` can flip the item to Approved
         within seconds of creation, and `Audit — approval activity` polls, so its own run may see
         the item already Approved and (since 2026-09-01) that flow owns "Uploaded" and "Rejected"
         ONLY — never "Approved", which Auto-route writes separately. If Approved/Routed exist and
         Uploaded does not, that is exactly the race, and it explains "bulk uploads specifically". */
      let raceNote = "";
      if (f.BulkImport === true) {
        const [hasApproved, hasRouted] = await Promise.all([
          auditHit(uniqueId, "Approved"),
          auditHit(uniqueId, "Routed"),
        ]);
        if (hasApproved || hasRouted) {
          raceNote =
            ` — RACE-SHAPED: "${hasApproved ? "Approved" : "Routed"}" exists, "Uploaded" does not. ` +
            `Matches the bulk-import race described at the top of this file.`;
        }
      }
      missing.push({
        ...f,
        uniqueId,
        reason: raceNote
          ? `no matching 'Uploaded' audit row${raceNote}`
          : "no matching 'Uploaded' audit row",
      });
    }
    console.log(`  ${matched} of ${files.length} have a matching "Uploaded" audit row.`);
    if (missing.length > 0) {
      console.log(`  %cMISSING (${missing.length}):`, "color:#a4262c;font-weight:bold");
      for (const m of missing) {
        console.log(
          `    #${m.Id}  ${m.FileLeafRef}  ${m.BulkImport === true ? "[BULK UPLOAD]" : "[form upload]"}` +
            `  created ${m.Created}` +
            (m.uniqueId ? `  (UniqueId ${m.uniqueId})` : "") +
            `  — ${m.reason}`,
        );
      }
      const missingBulk = missing.filter((m) => m.BulkImport === true).length;
      const missingForm = missing.length - missingBulk;
      console.log(
        `  Of the ${missing.length} missing: ${missingBulk} were Bulk Upload files, ${missingForm} were form uploads.`,
      );
    }
  }

  console.log(
    "\nDone. If files show up as MISSING here, the gap is real and is in Power Automate — " +
      "check that flow's own run history next, not this codebase.",
  );
})();

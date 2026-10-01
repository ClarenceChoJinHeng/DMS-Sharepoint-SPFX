/*
 * Trace ONE `CRS Submissions` record end to end — was its file really replaced, or just gone?
 * -------------------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, deleted or granted.
 *
 * Run it in the browser console, ON THE CRS SITE, signed in as an ADMINISTRATOR (a true admin
 * bypasses Draft Item Security, so this can see a pending file others can't).
 *
 * WHY THIS EXISTS: `check-tagging-status.js` can flag a record as "never tagged", but that alone
 * does not tell you WHY — specifically, whether the underlying file was genuinely replaced by a
 * newer upload (in which case the stuck record is a harmless reporting artefact) or whether it was
 * deleted with nothing taking its place (in which case the document may actually be gone). This
 * script answers that directly, by checking every fact rather than inferring from timing/naming:
 *
 *   1. Re-reads the record itself, fresh, right now — its CURRENT TagStatus/ReplacedAt/Modified,
 *      not a stale snapshot from an earlier check.
 *   2. Searches `CRS Submissions` for every OTHER row with the same filename, to find a genuine
 *      newer replacement record if one exists.
 *   3. Searches `CRS Audit Log` for every event tied to the record's own `ItemUniqueId` — this is
 *      the one query that can PROVE (not infer) what actually happened to that exact file, since it
 *      is keyed on the GUID itself rather than matched by name or timing.
 *   4. Checks whether the file's OWN original UniqueId still resolves right now, and separately
 *      whether a live document currently sits at the recorded path (which may be a different file
 *      entirely if a replacement landed there).
 *
 * Prints what was found under each question with no interpretation layered on top — the summary at
 * the end states only what the evidence directly shows, and says explicitly where it doesn't.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  // Leave blank to auto-detect. Set it if the auto-detected site below is wrong.
  const SITE_OVERRIDE = "";

  // The `CRS Submissions` record to trace, and its ORIGINAL `ItemUniqueId` (from the failed flow
  // run's own trigger body) — both are checked, never assumed correct.
  const RECORD_ID = 371;
  const ORIGINAL_UNIQUE_ID = "e56fd1fd-3689-47ef-a5a9-80eaf9238a2d";

  /* ── Resolve the site ──────────────────────────────────────────────────── */
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
      throw new Error(
        `Expected JSON and got "${type}" (HTTP ${r.status}) from ${url}\n` +
          `The site resolved to ${web}. If that is wrong, set SITE_OVERRIDE at the top and re-run.`,
      );
    }
    const body = await r.json().catch(() => undefined);
    return { status: r.status, body };
  };
  const tryGet = async (url) => {
    try {
      return await get(url);
    } catch (e) {
      return { status: undefined, error: e.message };
    }
  };

  const resolveListTitle = async (candidates, label) => {
    for (const t of candidates) {
      const res = await tryGet(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=Title`,
      );
      if (res.status === 200 && res.body && res.body.Title) return res.body.Title;
    }
    console.log(
      `✗ Could not find the ${label} list under any of: ${candidates.join(", ")}`,
    );
    return undefined;
  };

  /* ── 1. Re-read the record itself, fresh ─────────────────────────────────────────────── */
  console.log(`\n=== 1. RECORD #${RECORD_ID}, RIGHT NOW ===`);
  const subListTitle = await resolveListTitle(
    ["CRS Submissions", "GDC Submissions", "DMS Submissions"],
    "submissions",
  );
  if (!subListTitle) return;

  const RECORD_SELECT =
    "Id,FileName,ItemPath,LibraryTitle,UploadedAt,TagStatus,TagError," +
    "WithdrawnAt,ReplacedAt,ReplacedBy,ArchivedAt,SubmissionFileId,ItemUniqueId,Modified";
  const recordRes = await tryGet(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(subListTitle)}')/items(${RECORD_ID})?$select=${RECORD_SELECT}`,
  );
  if (recordRes.status !== 200) {
    console.log(`✗ Could not read record #${RECORD_ID}:`, recordRes.error || recordRes.status);
    return;
  }
  const record = recordRes.body;
  console.log(`  FileName:        ${record.FileName}`);
  console.log(`  ItemPath:        ${record.ItemPath}`);
  console.log(`  LibraryTitle:    ${record.LibraryTitle}`);
  console.log(`  UploadedAt:      ${record.UploadedAt}`);
  console.log(`  Modified:        ${record.Modified}  (has anything touched this row since?)`);
  console.log(`  TagStatus:       ${record.TagStatus || "(blank)"}`);
  console.log(`  TagError:        ${record.TagError || "(blank)"}`);
  console.log(`  ReplacedAt:      ${record.ReplacedAt || "(blank)"}`);
  console.log(`  ReplacedBy:      ${record.ReplacedBy || "(blank)"}`);
  console.log(`  WithdrawnAt:     ${record.WithdrawnAt || "(blank)"}`);
  console.log(`  ArchivedAt:      ${record.ArchivedAt || "(blank)"}`);
  console.log(`  SubmissionFileId:${record.SubmissionFileId}`);
  console.log(
    `  ItemUniqueId:    ${record.ItemUniqueId}  (matches the GUID given? ` +
      `${record.ItemUniqueId === ORIGINAL_UNIQUE_ID ? "YES" : "NO — DIFFERENT"})`,
  );

  /* ── 2. Any OTHER submission record for the SAME filename ────────────────────────────── */
  console.log(`\n=== 2. OTHER "CRS Submissions" ROWS WITH THE SAME FILENAME ===`);
  let allRows = [];
  {
    let url =
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(subListTitle)}')/items` +
      `?$select=${RECORD_SELECT}&$top=5000&$orderby=Id`;
    let pages = 0;
    while (url && pages < 20) {
      const res = await tryGet(url);
      if (res.status !== 200) {
        console.log("✗ Could not read the submissions list:", res.error || res.status);
        break;
      }
      allRows = allRows.concat(res.body.value || []);
      url = res.body["odata.nextLink"];
      pages += 1;
    }
  }
  const siblings = allRows.filter(
    (r) => r.FileName === record.FileName && r.Id !== RECORD_ID,
  );
  if (siblings.length === 0) {
    console.log(
      "  NONE FOUND. No other record in CRS Submissions shares this exact filename — there is no\n" +
        "  visible \"newer replacement record\" for this upload anywhere in the list.",
    );
  } else {
    for (const s of siblings) {
      console.log(
        `  #${s.Id}  UploadedAt: ${s.UploadedAt}  TagStatus: ${s.TagStatus || "(blank)"}  ` +
          `SubmissionFileId: ${s.SubmissionFileId}  ItemUniqueId: ${s.ItemUniqueId}`,
      );
    }
  }

  /* ── 3. Full audit-log timeline for the ORIGINAL file's OWN UniqueId ─────────────────────
     This is the one query that can PROVE what happened, since it is keyed on the GUID itself —
     never matched by filename or timing, which is where an earlier write-up was only inferring. */
  console.log(`\n=== 3. "CRS AUDIT LOG" — EVERY EVENT TIED TO ${ORIGINAL_UNIQUE_ID} ===`);
  const auditListTitle = await resolveListTitle(
    ["CRS Audit Log", "GDC Audit Log", "DMS Audit Log"],
    "audit log",
  );
  if (!auditListTitle) {
    console.log("  Could not resolve the audit log list — skipping this check.");
  } else {
    const auditRes = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(auditListTitle)}')/items` +
        `?$select=EventTime,EventType,ActorEmail,ActorName,ItemName,ItemUniqueId,Details` +
        `&$filter=ItemUniqueId eq '${ORIGINAL_UNIQUE_ID}'&$orderby=EventTime asc&$top=100`,
    );
    if (auditRes.status !== 200) {
      console.log(
        "  ✗ The $filter query failed (",
        auditRes.error || auditRes.status,
        ") — the column may not be filterable this way. Not falling back to a full-list scan\n" +
          "  automatically; re-run with a manual check if this matters.",
      );
    } else {
      const events = (auditRes.body && auditRes.body.value) || [];
      if (events.length === 0) {
        console.log(
          "  NO EVENTS FOUND with this exact ItemUniqueId. Either the audit log never recorded\n" +
            "  anything for this file (a flow gap, not evidence either way), or the \"Deleted\" row\n" +
            "  seen earlier belongs to a DIFFERENT file that merely shares the name and timing.",
        );
      } else {
        for (const e of events) {
          console.log(
            `  ${e.EventTime}  ${e.EventType}  (${e.ActorEmail || e.ActorName || "?"})`,
          );
          if (e.Details) console.log(`    ${e.Details}`);
        }
      }
    }
  }

  /* ── 4. Does the ORIGINAL GUID still resolve, and does a live file sit at the path now ──── */
  console.log(`\n=== 4. LIVE STATE, RIGHT NOW ===`);
  const origRes = await tryGet(
    `${web}/_api/web/GetFileById(guid'${ORIGINAL_UNIQUE_ID}')?$select=Exists`,
  );
  console.log(
    `  GetFileById(${ORIGINAL_UNIQUE_ID}): ` +
      (origRes.status === 200
        ? "STILL RESOLVES — this file is NOT gone."
        : `HTTP ${origRes.status} — this file no longer resolves under its original identity.`),
  );

  if (record.ItemPath) {
    const pathRes = await tryGet(
      // OData alias form, never an inline literal — gotcha #9.
      `${web}/_api/web/GetFileByServerRelativeUrl(@f)?$select=Exists,UniqueId&@f='${encodeURIComponent(record.ItemPath).replace(/'/g, "''")}'`,
    );
    if (pathRes.status === 200) {
      console.log(
        `  A live file currently sits at "${record.ItemPath}", with UniqueId ` +
          `${pathRes.body.UniqueId} — ` +
          (pathRes.body.UniqueId === ORIGINAL_UNIQUE_ID
            ? "the SAME identity as the original record (never actually changed)."
            : "a DIFFERENT identity — something else now occupies this path (consistent with a replacement)."),
      );
    } else {
      console.log(
        `  Nothing currently resolves at "${record.ItemPath}" (HTTP ${pathRes.status}) — ` +
          "no live document sits there right now, under any identity.",
      );
    }
  }

  /* ── 5. Summary — states only what was directly confirmed above ─────────────────────────── */
  console.log(`\n=== SUMMARY — facts only, no interpretation added ===`);
  console.log(
    "  Q: Is the deleted file genuinely the same one this record originally pointed at?\n" +
      "  A: See section 3 — only a positive audit-log match on the exact GUID confirms this.",
  );
  console.log(
    "  Q: Was it replaced (a newer copy exists) or just gone?\n" +
      "  A: See sections 2 and 4 — a sibling record and/or a live file with a different identity at\n" +
      "     the same path means replaced; neither present means it is likely just gone.",
  );
  console.log(
    "  Q: Has anything retried or touched this record since the original failure?\n" +
      "  A: Compare record.Modified above against record.UploadedAt — if they differ, something\n" +
      "     wrote to this row again after it was created.",
  );
})();

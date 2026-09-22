/*
 * Diagnose why an archived file still reads "Deleted" (instead of "Archived") on My Submissions.
 *
 * READ ONLY — every request is a GET, nothing is created, renamed or deleted.
 *
 * WHAT THIS ANSWERS, for one named file sitting in Archive / HC Archive:
 *   1. Does the file itself carry a `SubmissionFileId`? (blank = it predates the 2026-08-22
 *      submission-record feature, or was uploaded by a route that never stamped it — the archive
 *      mover's `ArchivedAt` stamp has NOTHING to join against in that case, and no amount of
 *      re-running the archive flow can fix it retroactively. This is the single most likely cause.)
 *   2. Does a matching row exist on `CRS Submissions` (by that same SubmissionFileId)?
 *   3. Is `ArchivedAt` actually set on that row?
 *
 * My Submissions reads `CRS Submissions` (via `readSubmissionRecords` / `mergeRecords` in
 * `src/shared/submissionRecords.ts`) and shows "Archived" ONLY when that row's `ArchivedAt` column
 * is non-blank — see the `RecordState` precedence comment there. If any of the three checks above
 * comes back negative, that is the actual cause, independent of whether the archive-mover FLOW
 * itself ran correctly.
 *
 * HOW TO RUN
 * ----------
 * Open the site, F12 -> Console, set FILENAME below to the exact name shown in the Archive
 * library (or leave it blank to just list what's in both archive libraries), paste, Enter.
 */
(async () => {
  // ── Configure ────────────────────────────────────────────────────────────
  const FILENAME = ""; // e.g. "test - test - test - 01092026.pdf" — leave blank to list files instead
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  // ─────────────────────────────────────────────────────────────────────────

  // Same resolveWeb() shape as the other scripts in this project (check-archive-column-gap.js,
  // dump-segment-folders.js) — a modern page's URL can be any shape, and a wrong resolution answers
  // 200 with a page's own HTML rather than an obvious error.
  const resolveWeb = () => {
    if (SITE_OVERRIDE) return SITE_OVERRIDE.replace(/\/$/, "");
    try {
      const ctx = window._spPageContextInfo;
      if (ctx && ctx.webAbsoluteUrl) return String(ctx.webAbsoluteUrl).replace(/\/$/, "");
    } catch (e) {
      /* not defined on some modern pages */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
    return location.origin + (m ? m[1] : "");
  };
  const web = resolveWeb();
  const get = async (u) => {
    const r = await fetch(u, { headers: { Accept: "application/json;odata=nometadata" } });
    const ct = r.headers.get("content-type") || "";
    if (!/json/i.test(ct)) {
      throw new Error(
        `expected JSON and got "${ct || "no content-type"}" from ${u}\n` +
          `    The site URL resolved to ${web} — if that is not this site's web, ` +
          `set SITE_OVERRIDE at the top of this script and run it again.`
      );
    }
    const body = await r.json();
    if (!r.ok) throw new Error("HTTP " + r.status + " on " + u + "\n" + JSON.stringify(body));
    return body;
  };

  console.log("%cSite: " + web, "font-weight:bold;font-size:13px");

  // 1. Every document library, matched by URL SEGMENT — never title (gotcha #12, and this client
  //    renames libraries routinely). Archive/HCArchive segments never change even when retitled.
  const lists = (await get(
    `${web}/_api/web/lists?$select=Title,Id&$expand=RootFolder&$filter=BaseTemplate eq 101`
  )).value;
  const bySegment = {};
  for (const l of lists) {
    const seg = decodeURIComponent((l.RootFolder?.ServerRelativeUrl || "").split("/").pop() || "");
    bySegment[seg] = l;
  }
  const archiveLibs = [
    { seg: "Archive", lib: bySegment["Archive"] },
    { seg: "HCArchive", lib: bySegment["HCArchive"] },
  ].filter((x) => x.lib);

  if (archiveLibs.length === 0) {
    console.error("No library at /Archive or /HCArchive on this site — nothing to check.");
    return;
  }
  console.log(
    "Archive libraries found: " + archiveLibs.map((x) => `"${x.lib.Title}" (/${x.seg})`).join(", ")
  );

  // 2. Resolve the item(s) to check.
  const itemFields =
    "Id,FileLeafRef,FileRef,Created,Modified,SubmissionFileId,SubmissionId,BatchId";
  let candidates = [];
  for (const { seg, lib } of archiveLibs) {
    let items;
    try {
      const url = FILENAME
        ? `${web}/_api/web/lists(guid'${lib.Id}')/items` +
          `?$select=${itemFields}&$filter=FileLeafRef eq '${encodeURIComponent(FILENAME.replace(/'/g, "''"))}'`
        : `${web}/_api/web/lists(guid'${lib.Id}')/items?$select=${itemFields}&$top=50&$orderby=Modified desc`;
      items = (await get(url)).value;
    } catch (e) {
      console.warn(`  Could not read "${lib.Title}": ${e.message}`);
      continue;
    }
    for (const it of items) candidates.push({ lib: lib.Title, seg, it });
  }

  if (!FILENAME) {
    console.log("%c\nFiles currently in the archive libraries (most recent 50 per library):", "font-weight:bold");
    console.table(
      candidates.map(({ lib, it }) => ({
        library: lib,
        name: it.FileLeafRef,
        SubmissionFileId: it.SubmissionFileId || "(blank)",
        Modified: it.Modified,
      }))
    );
    console.log('\nSet FILENAME at the top of this script to one of the "name" values above and re-run.');
    return;
  }

  if (candidates.length === 0) {
    console.error(`No file named "${FILENAME}" found in either archive library.`);
    return;
  }

  // 3. Resolve "CRS Submissions" the same way the app does — CRS then DMS prefix.
  let subsList;
  for (const p of ["CRS Submissions", "DMS Submissions"]) {
    try {
      subsList = await get(`${web}/_api/web/lists/getbytitle('${encodeURIComponent(p)}')?$select=Title,Id`);
      break;
    } catch (e) {
      /* try next candidate */
    }
  }
  if (!subsList) {
    console.error(
      'Could not resolve a "CRS Submissions" (or "DMS Submissions") list on this site. ' +
        "Without it, My Submissions has nothing to join the archived file against at all."
    );
    return;
  }
  console.log(`\nSubmission-record list: "${subsList.Title}"`);

  // 4. For each matching file, look up its record by SubmissionFileId and report.
  for (const { lib, seg, it } of candidates) {
    console.log("%c\n" + "=".repeat(70), "color:#888");
    console.log(`%c"${it.FileLeafRef}"  in  "${lib}"  (/${seg})`, "font-weight:bold;font-size:13px");
    console.log(`  path:              ${it.FileRef}`);
    console.log(`  Modified:          ${it.Modified}`);
    console.log(`  SubmissionFileId:  ${it.SubmissionFileId || "(BLANK)"}`);

    if (!it.SubmissionFileId) {
      console.log(
        "%c  ⇒ THIS FILE HAS NO STAMP. The archive mover's `ArchivedAt` stamp has nothing to join " +
          "against — it can only ever write onto a `CRS Submissions` row found BY this id. " +
          "This file will read \"Deleted\" on My Submissions no matter how many times the archive " +
          "flow is re-run, because there is no record for it to update. " +
          "This is expected for anything uploaded before 2026-08-22, or filed by a route that " +
          "never wrote the stamp — not a bug in the archive-mover fix.",
        "color:#a4262c"
      );
      continue;
    }

    const recUrl =
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(subsList.Title)}')/items` +
      `?$select=Id,SubmissionFileId,ArchivedAt,ReplacedAt,WithdrawnAt,FileName,ItemPath,LibraryTitle` +
      `&$filter=SubmissionFileId eq '${encodeURIComponent(it.SubmissionFileId)}'`;
    let recs;
    try {
      recs = (await get(recUrl)).value;
    } catch (e) {
      console.warn(`  Could not read ${subsList.Title}: ${e.message}`);
      continue;
    }

    if (recs.length === 0) {
      console.log(
        "%c  ⇒ NO MATCHING ROW on " + subsList.Title + " for this SubmissionFileId. " +
          "The file carries a stamp but the record it should join to is missing or unreadable — " +
          "worth checking whether the row was ever written (readLibrary/writeSubmissionRecord at " +
          "upload time), separately from anything about the archive mover.",
        "color:#a4262c"
      );
      continue;
    }
    if (recs.length > 1) {
      console.warn(`  ⚠ ${recs.length} rows share this SubmissionFileId — ambiguous, listing all:`);
    }
    for (const r of recs) {
      console.log(`  Submissions row #${r.Id}:`);
      console.log(`    LibraryTitle (as recorded at upload): ${r.LibraryTitle}`);
      console.log(`    ArchivedAt:  ${r.ArchivedAt || "(BLANK)"}`);
      console.log(`    ReplacedAt:  ${r.ReplacedAt || "(blank)"}`);
      console.log(`    WithdrawnAt: ${r.WithdrawnAt || "(blank)"}`);
      if (r.ArchivedAt) {
        console.log(
          "%c    ⇒ ArchivedAt IS set. My Submissions should show \"Archived\" for this file. " +
            "If it still shows \"Deleted\" in the browser, that is very likely a STALE TAB/bundle — " +
            "hard-refresh (Ctrl+Shift+R) the My Submissions page and check again before assuming " +
            "the code is still wrong.",
          "color:#0f6c3f"
        );
      } else if (r.ReplacedAt) {
        console.log('    ⇒ ReplacedAt is set — this record will show "Replaced", not "Deleted".');
      } else if (r.WithdrawnAt) {
        console.log('    ⇒ WithdrawnAt is set — this record will show "Cancelled", not "Deleted".');
      } else {
        console.log(
          "%c    ⇒ NONE of ArchivedAt/ReplacedAt/WithdrawnAt are set on this row, even though the " +
            "file is physically sitting in the archive library right now. This means the archive " +
            "mover's stamp step (GetArchivedFileId → HasStamp → GetArchivedRecordId → " +
            "StampArchivedRecord) did not actually write to THIS row when this file was moved — " +
            "worth re-checking that flow's run history for the run that moved this specific file.",
          "color:#a4262c"
        );
      }
    }
  }
})();

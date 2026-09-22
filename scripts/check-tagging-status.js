/*
 * Is any uploaded document still missing its tags?
 * -------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, deleted or granted.
 *
 * Run it in the browser console, ON THE CRS SITE, signed in as an ADMINISTRATOR (it reads the
 * whole `CRS Submissions` list, not just your own rows).
 *
 * WHY THIS EXISTS: since the 2026-09-18 tag/approve-by-proxy change, `Form.tsx`/`BulkUpload.tsx`
 * no longer write a document's metadata directly. Instead they write a `CRS Submissions` row
 * carrying `TagPayload` (the exact field values that need to land on the document), and
 * `CRS — Apply pending tags` (Power Automate, running as the service account) reads that payload
 * and applies it, then stamps `TagStatus` back onto the row: blank/absent = not yet processed,
 * "Tagged" = applied, "Failed" = see `TagError`. This script checks every recent submission
 * against that state, and — for anything marked "Tagged" — actually re-reads the live document to
 * confirm the fields the payload named are genuinely non-blank, rather than trusting the stamp.
 *
 * FOUR BUCKETS, IN THE ORDER THEY MATTER:
 *   1. FAILED       — TagStatus = "Failed". Real, and TagError says why.
 *   2. NEVER TAGGED — TagStatus is still blank/absent, and the row is older than the grace period
 *                      below. `CRS — Apply pending tags` should have reached it by now and has not.
 *   3. TAGGED BUT BLANK — TagStatus = "Tagged", but re-reading the live item shows one or more of
 *                      the fields the payload named are still empty. This is the case a green
 *                      TagStatus cannot rule out on its own — worth checking because a flow that
 *                      reports success on a partial write is exactly the failure mode this project
 *                      has hit before with `validateUpdateListItem` (gotcha #4: one bad field
 *                      SHOULD fail the whole call, but only if the flow's own HasException check is
 *                      actually reading the response body correctly).
 *   4. STILL SETTLING — blank TagStatus, but uploaded within the grace period. Not a problem; the
 *                      flow just has not had time to run yet. Listed only so a fast re-run some
 *                      minutes later can confirm they cleared, not because they need action now.
 *
 * Anything WITHDRAWN, REPLACED or ARCHIVED is skipped outright — those documents are no longer the
 * live thing the tag payload was ever meant to land on, so a blank TagStatus there is expected and
 * meaningless.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  // Leave blank to auto-detect. Set it if the auto-detected site below is wrong.
  const SITE_OVERRIDE = "";

  // Rows uploaded more recently than this are still allowed to be "not yet tagged" — the flow
  // polls, it does not run instantly. Generous on purpose: understating a problem for a few extra
  // minutes costs nothing, crying wolf about a submission still genuinely in flight costs trust in
  // the tool the next time it is run.
  const GRACE_MINUTES = 10;

  // Cap on how many "Tagged" rows get the expensive live-field re-check (one or two extra requests
  // each). Set to a smaller number for a quick pass, or a larger one for a thorough overnight run.
  const MAX_FIELD_CHECKS = 300;

  // Only check submissions whose recorded path CONTAINS one of these (case-insensitive). Leave
  // empty to check everything, as before. Narrows a big list down to one test's own files — the HC
  // and non-HC copies of the SAME folder both still match, since only the LIBRARY segment differs
  // ("ApprovalDocument" vs "HCApprovalDocument"), never the tail of the path.
  // 2026-09-21: the front-to-back upload/approve test ground, NBPOLHO > CDS > UPSUPPORT > 2024 > Ara
  // > Approval Papers.
  const PATH_FILTER = ["UPSUPPORT/2025/Approval Papers"];

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
        `Expected JSON and got "${type}" from ${url}\n` +
          `The site resolved to ${web}. If that is wrong, set SITE_OVERRIDE at the top and re-run.`,
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

  /* ── 1. Resolve the `CRS Submissions` list — this client renames things, so probe ────────── */
  const listCandidates = [
    "CRS Submissions",
    "GDC Submissions",
    "DMS Submissions",
  ];
  let listTitle;
  for (const t of listCandidates) {
    const probe = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')?$select=Title,ItemCount`,
    );
    if (probe && probe.Title) {
      listTitle = probe.Title;
      console.log(
        `Submissions list: "${listTitle}" (${probe.ItemCount} items)`,
      );
      break;
    }
  }
  if (!listTitle) {
    console.log(
      "✗ Could not find the submissions list under any known name — checked:",
      listCandidates.join(", "),
    );
    return;
  }

  /* ── 2. Read every row (paged, in case the list has grown past one page) ───────────────── */
  const SELECT =
    "Id,SubmissionFileId,FileName,ItemPath,LibraryTitle,UploadedAt,TagPayload,TagStatus,TagError," +
    "WithdrawnAt,ReplacedAt,ArchivedAt";
  let rows = [];
  let url = `${web}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=${SELECT}&$top=5000&$orderby=Id`;
  let pages = 0;
  while (url && pages < 20) {
    const res = await tryGet(url);
    if (res.__error) {
      console.log("✗ Could not read the submissions list:", res.__error);
      return;
    }
    rows = rows.concat(res.value || []);
    url = res["odata.nextLink"];
    pages += 1;
  }
  console.log(`Read ${rows.length} row(s) across ${pages} page(s).`);
  if (pages >= 20) {
    console.log(
      "⚠ Hit the page cap — this run may not cover every row. Re-run with a narrower window if needed.",
    );
  }

  if (PATH_FILTER.length > 0) {
    const before = rows.length;
    const needles = PATH_FILTER.map((f) => f.toLowerCase());
    rows = rows.filter((r) => {
      const path = (r.ItemPath || "").toLowerCase();
      return needles.some((n) => path.indexOf(n) > -1);
    });
    console.log(
      `Narrowed to ${rows.length} of ${before} row(s) matching PATH_FILTER: ${PATH_FILTER.join(", ")}`,
    );
  }

  /* ── 3. Bucket everything that is still live ─────────────────────────────────────────── */
  const now = Date.now();
  const failed = [];
  const neverTagged = [];
  const settling = [];
  const taggedRows = [];
  let skippedGone = 0;

  for (const r of rows) {
    if (r.WithdrawnAt || r.ReplacedAt || r.ArchivedAt) {
      skippedGone += 1;
      continue;
    }
    const uploadedAt = r.UploadedAt ? new Date(r.UploadedAt) : undefined;
    const ageMinutes = uploadedAt
      ? (now - uploadedAt.getTime()) / 60000
      : Infinity;

    if (r.TagStatus === "Failed") {
      failed.push(r);
    } else if (r.TagStatus === "Tagged") {
      taggedRows.push(r);
    } else if (!r.TagStatus) {
      if (ageMinutes < GRACE_MINUTES) settling.push(r);
      else neverTagged.push(r);
    }
    // Any other TagStatus value is unrecognised and deliberately falls through unreported here —
    // this script only knows the three the flow is supposed to write.
  }

  console.log(
    `\nSkipped ${skippedGone} withdrawn/replaced/archived row(s) — not this script's concern.`,
  );

  /* ── 4. Report the two conclusive buckets first ──────────────────────────────────────── */
  console.log(`\n=== FAILED (${failed.length}) ===`);
  for (const r of failed) {
    console.log(`  #${r.Id}  ${r.FileName}`);
    console.log(`    ${r.ItemPath || "(no path recorded)"}`);
    console.log(
      `    Error: ${(r.TagError || "(no TagError recorded)").slice(0, 300)}`,
    );
  }

  console.log(
    `\n=== NEVER TAGGED, older than ${GRACE_MINUTES} minutes (${neverTagged.length}) ===`,
  );
  for (const r of neverTagged) {
    const ageMin = r.UploadedAt
      ? Math.round((now - new Date(r.UploadedAt).getTime()) / 60000)
      : "?";
    console.log(`  #${r.Id}  ${r.FileName}  (uploaded ${ageMin} min ago)`);
    console.log(`    ${r.ItemPath || "(no path recorded)"}`);
    console.log(
      `    TagPayload present: ${r.TagPayload ? "yes" : "NO — nothing for the flow to apply"}`,
    );
  }

  /* ── 5. For "Tagged" rows, re-read the live item and check the payload's own fields ──────
     The join mirrors the app's own: filter the recorded library (and the usual approved-side
     candidates, since an approved document has moved libraries since the row was written) on
     `SubmissionFileId`. `_x005f_` is undone the same way `FieldValuesAsText` needs it elsewhere in
     this project — OData double-encodes the underscore in that response's own keys. */
  const norm = (k) => k.replace(/_x005f_/g, "_");
  const APPROVED_SIDE_CANDIDATES = [
    "Restricted & Confidential Document",
    "Documents",
    "Highly Confidential Document",
    "HC Documents",
    "HC Document",
  ];

  // Fields the upload form treats as OPTIONAL (per `missingForFile` in Form.tsx — only Document
  // Name/Date/Confidentiality are actually required). A blank value here means the uploader chose
  // not to fill it in, not that tagging failed — `composeUploadBase` even drops the filename
  // segment for a blank Project Name/Vendor, which is how you can tell from the filename alone.
  // Internal names, matching `FILE_FIXED_FIELDS` in `shared/documentDetails.ts`.
  const OPTIONAL_FIELDS = [
    "ProjectName",
    "Vendor_x002f_CustomerName",
    "Remark",
    "Keyword",
  ];

  const toCheck = taggedRows.slice(0, MAX_FIELD_CHECKS);
  if (taggedRows.length > MAX_FIELD_CHECKS) {
    console.log(
      `\n⚠ ${taggedRows.length} rows are marked Tagged; only checking the first ${MAX_FIELD_CHECKS} ` +
        `(raise MAX_FIELD_CHECKS at the top for a full pass).`,
    );
  }

  const taggedButBlank = [];
  const blankOptionalOnly = [];
  const taggedNotFound = [];
  let checked = 0;

  for (const r of toCheck) {
    checked += 1;
    if (checked % 25 === 0)
      console.log(`  ...checked ${checked}/${toCheck.length}`);

    let payload;
    try {
      payload = r.TagPayload ? JSON.parse(r.TagPayload) : [];
    } catch (e) {
      taggedButBlank.push({
        r,
        reason: `TagPayload is not valid JSON: ${e.message}`,
      });
      continue;
    }
    if (!Array.isArray(payload) || payload.length === 0) continue;

    const libCandidates = [r.LibraryTitle, ...APPROVED_SIDE_CANDIDATES].filter(
      (v, i, a) => v && a.indexOf(v) === i,
    );
    let found;
    for (const title of libCandidates) {
      const hit = await tryGet(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/items` +
          `?$select=Id&$filter=SubmissionFileId eq '${(r.SubmissionFileId || "").replace(/'/g, "''")}'&$top=1`,
      );
      if (hit && !hit.__error && hit.value && hit.value.length > 0) {
        found = { title, id: hit.value[0].Id };
        break;
      }
    }
    if (!found) {
      taggedNotFound.push(r);
      continue;
    }

    const ft = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(found.title)}')/items(${found.id})/FieldValuesAsText`,
    );
    if (!ft || ft.__error) {
      taggedButBlank.push({
        r,
        reason: "Could not re-read the live item's fields to confirm.",
      });
      continue;
    }
    const plain = {};
    for (const k of Object.keys(ft)) plain[norm(k)] = ft[k];

    const blankFields = [];
    const blankOptionalFields = [];
    for (const entry of payload) {
      const name = entry.FieldName;
      if (!name) continue;
      const val = plain[name];
      const isBlank =
        val === undefined || val === null || String(val).trim() === "";
      if (!isBlank) continue;
      if (OPTIONAL_FIELDS.indexOf(name) > -1) blankOptionalFields.push(name);
      else blankFields.push(name);
    }
    if (blankFields.length > 0) {
      taggedButBlank.push({
        r,
        reason: `Blank on the live item: ${blankFields.join(", ")}`,
        library: found.title,
      });
    } else if (blankOptionalFields.length > 0) {
      blankOptionalOnly.push({
        r,
        fields: blankOptionalFields,
        library: found.title,
      });
    }
  }

  console.log(
    `\n=== TAGGED BUT ONE OR MORE FIELDS ARE STILL BLANK (${taggedButBlank.length}) ===`,
  );
  for (const { r, reason, library } of taggedButBlank) {
    console.log(
      `  #${r.Id}  ${r.FileName}${library ? `  (found in "${library}")` : ""}`,
    );
    console.log(`    ${reason}`);
  }

  console.log(
    `\n=== BLANK, BUT ONLY OPTIONAL FIELDS — NOT A PROBLEM (${blankOptionalOnly.length}) ===`,
  );
  console.log(
    "  (Project Name / Vendor-Customer / Remark / Keyword are all optional on the upload form —",
  );
  console.log(
    "   blank here means the uploader left it empty on purpose, listed only for transparency.)",
  );
  for (const { r, fields, library } of blankOptionalOnly) {
    console.log(
      `  #${r.Id}  ${r.FileName}${library ? `  (found in "${library}")` : ""}`,
    );
    console.log(`    Blank (optional, fine): ${fields.join(", ")}`);
  }

  console.log(
    `\n=== TAGGED, BUT THE LIVE ITEM COULD NOT BE FOUND AT ALL (${taggedNotFound.length}) ===`,
  );
  console.log(
    "  (Checked the recorded library plus the usual approved-side names. Not necessarily",
  );
  console.log(
    "   a problem — the document may have moved to the archive, which this script does not",
  );
  console.log(
    "   search — but worth a manual look if this list is non-empty.)",
  );
  for (const r of taggedNotFound) {
    console.log(`  #${r.Id}  ${r.FileName}  (recorded in "${r.LibraryTitle}")`);
  }

  console.log(
    `\n=== STILL SETTLING, under ${GRACE_MINUTES} minutes old (${settling.length}) ===`,
  );
  console.log(
    "  (Not a problem yet — just uploaded, the flow has not necessarily reached it.)",
  );
  for (const r of settling) {
    const ageMin = r.UploadedAt
      ? Math.round((now - new Date(r.UploadedAt).getTime()) / 60000)
      : "?";
    console.log(`  #${r.Id}  ${r.FileName}  (${ageMin} min ago)`);
  }

  /* ── 6. Summary ───────────────────────────────────────────────────────────────────────── */
  // `blankOptionalOnly` is deliberately NOT in `problems` — those rows have nothing wrong, only an
  // optional field the uploader chose to leave empty. It IS included in "confirmed clean" below,
  // since that is what it is.
  const problems = failed.length + neverTagged.length + taggedButBlank.length;
  console.log(
    `\n=== SUMMARY: ${problems} row(s) need attention ` +
      `(${failed.length} failed, ${neverTagged.length} never tagged, ${taggedButBlank.length} tagged-but-blank). ` +
      `${settling.length} still settling, ${taggedNotFound.length} tagged-but-unresolvable, ` +
      `${taggedRows.length - taggedButBlank.length - taggedNotFound.length} confirmed clean ` +
      `(${blankOptionalOnly.length} of those with an optional field left blank on purpose). ===`,
  );
})();

/*
 * Does re-asserting a taxonomy field's correct value fix a corrupted Label?
 * ---------------------------------------------------------------------------
 * ⚠⚠ THIS SCRIPT CAN WRITE. DEFAULTS TO A DRY RUN. Read the report it prints, and only set
 * `DO_WRITE = true` below and run it again once you are satisfied the "before" state matches
 * what you expect.
 *
 * WHY THIS EXISTS
 * CRS Search's Document Type / Year / Confidentiality filters return "No Result Found" for
 * documents that plainly match — traced (2026-09-24, see CLAUDE.md "CRS SEARCH'S 'Document
 * Type: Term Sheet' FILTER") to the RAW taxonomy field's own `Label` sub-property being stored
 * as the term's numeric WssId ("14") instead of its real text ("Term Sheet"), on the exact
 * document a test targeted. `FieldValuesAsText` resolves the SAME field correctly at the SAME
 * moment, which is the tell: the term store itself is fine, and something about the RAW stored
 * value is stale or unresolved.
 *
 * The staged `TagPayload` on `CRS Submissions` and the raw Inputs of `CRS — Apply pending
 * tags`'s `ApplyTags` action were BOTH confirmed byte-exact correct
 * ("Term Sheet|ffd9961d-f70d-4044-b5b9-a404a1d50f3d", not "14|..."). So the flow sent the right
 * data; whatever went wrong happened on SharePoint's side, after accepting that write — most
 * likely `TaxonomyHiddenList` (a per-site term cache) being cold for a term GUID at write time.
 *
 * THIS SCRIPT TESTS THE ONE CHEAP REPAIR WORTH TRYING FIRST: does simply writing the exact same
 * correct value a SECOND time (through the identical `validateUpdateListItem` mechanism the flow
 * already uses) force SharePoint to resolve and cache the Label properly? If yes, the fix is
 * trivial (retry once, or a small delay). If no, this is a deeper platform issue needing a
 * different mitigation (a forced re-save via the SharePoint UI's own taxonomy picker, waiting for
 * a background cache job, or opening a support case).
 *
 * HOW TO RUN
 *   1. F12 → Console on the CRS site, signed in as a site administrator.
 *   2. Paste this whole file, Enter. `DO_WRITE` is `false` — nothing is written on this pass.
 *      It reads the current raw values and FieldValuesAsText, and prints exactly what a write
 *      pass WOULD send.
 *   3. Read the printed report. If the "before" state matches what you already know (raw Label
 *      shows the WssId digit, FieldValuesAsText shows the real term), set `DO_WRITE = true` and
 *      run the WHOLE file again.
 *   4. Read the "after" report. If the raw Label now matches FieldValuesAsText, the repair
 *      worked — the fix for the flow (or for any other stuck document) is to have it retry once
 *      on the same GUID rather than trust the first write. If it is unchanged, the cache issue
 *      is not fixed by a second identical write and needs a different approach.
 *
 * READ ONLY unless DO_WRITE is true, in which case it makes exactly ONE write per configured
 * item — the SAME `validateUpdateListItem` call shape `CRS — Apply pending tags` already uses,
 * re-asserting values that are supposed to already be correct. Nothing here changes what the
 * document IS tagged as; it only re-sends the same tag.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────────────────────── */
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const DO_WRITE = false; // ⚠ set to true ONLY after reading the dry-run "before" report below

  // One or more documents to check, identified the same way the rest of this project's own
  // request/deletion resolution does — by the stamped `SubmissionFileId`, which survives a
  // document being routed between libraries (a raw UniqueId does not).
  const ITEMS = [{ submissionFileId: "SFI-20260924-JNXK" }];

  /* ── Resolve the site ──────────────────────────────────────────────────────────────────── */
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
  console.log("Site:", web, "| DO_WRITE:", DO_WRITE);

  const bust = (u) => u + (u.indexOf("?") > -1 ? "&" : "?") + "_=" + Date.now();

  const get = async (url) => {
    const r = await fetch(bust(url), {
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

  const getDigest = async () => {
    const r = await fetch(bust(`${web}/_api/contextinfo`), {
      method: "POST",
      headers: { Accept: "application/json;odata=nometadata" },
    });
    const type = r.headers.get("content-type") || "";
    if (type.indexOf("json") === -1) {
      throw new Error(
        `Expected JSON getting a request digest and got "${type}" — the site resolved to ${web}, ` +
          `which is almost certainly wrong for this page. Set SITE_OVERRIDE and re-run.`,
      );
    }
    const j = await r.json();
    return j.FormDigestValue;
  };

  // Every library this system's documents can be filed into. Mirrors the candidate lists in
  // shared/naming.ts (resolved by TITLE here since we search by it, not the URL segment).
  const LIBRARY_CANDIDATES = [
    "Approval for Document",
    "Approval Document",
    "Staging",
    "Approval for Highly Confidential Document",
    "HC Approval Document",
    "Restricted & Confidential Document",
    "Documents",
    "Shared Documents",
    "Highly Confidential Document",
    "HC Documents",
    "Archive Restricted & Confidential Document",
    "Archive",
    "Archive Highly Confidential Document",
    "HC Archive",
  ];

  const TAX_FIELDS = ["Document_x0020_Type", "Year", "Confidentiality_x0020_Level"];

  const findItem = async (submissionFileId) => {
    for (const title of LIBRARY_CANDIDATES) {
      const hit = await tryGet(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/items` +
          `?$select=Id&$filter=SubmissionFileId eq '${submissionFileId.replace(/'/g, "''")}'&$top=1`,
      );
      if (hit && !hit.__error && hit.value && hit.value.length > 0) {
        return { library: title, id: hit.value[0].Id };
      }
    }
    return undefined;
  };

  const readRawTax = async (library, id) => {
    const sel = TAX_FIELDS.join(",");
    return tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(library)}')/items(${id})?$select=${sel}`,
    );
  };

  const readTextTax = async (library, id) => {
    const ft = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(library)}')/items(${id})/FieldValuesAsText`,
    );
    if (!ft || ft.__error) return ft;
    // OData double-encodes the underscore in FieldValuesAsText's own keys.
    const norm = {};
    for (const k of Object.keys(ft)) norm[k.replace(/_x005f_/g, "_")] = ft[k];
    return norm;
  };

  const describeRaw = (raw) => {
    const out = {};
    for (const f of TAX_FIELDS) {
      const v = raw[f];
      out[f] = v && typeof v === "object" ? `Label="${v.Label}" WssId=${v.WssId}` : `(${v})`;
    }
    return out;
  };

  for (const item of ITEMS) {
    console.log(`\n=== ${item.submissionFileId} ===`);
    const found = await findItem(item.submissionFileId);
    if (!found) {
      console.log("  ✗ Could not find a live item carrying this SubmissionFileId in any known library.");
      continue;
    }
    console.log(`  Found in "${found.library}", item #${found.id}`);

    const before = await readRawTax(found.library, found.id);
    if (before.__error) {
      console.log("  ✗ Could not read raw taxonomy fields:", before.__error);
      continue;
    }
    const beforeText = await readTextTax(found.library, found.id);

    console.log("  BEFORE — raw field (what search compares against):");
    const beforeDesc = describeRaw(before);
    for (const f of TAX_FIELDS) console.log(`    ${f}: ${beforeDesc[f]}`);
    console.log("  BEFORE — FieldValuesAsText (what the app's detail panels show):");
    for (const f of TAX_FIELDS) console.log(`    ${f}: ${beforeText.__error ? "(could not read)" : beforeText[f]}`);

    // Corrupted = the raw Label equals the WssId as a string (the exact symptom this script
    // exists to test a repair for). Correct means the raw Label already reads real text.
    const corrupted = TAX_FIELDS.filter((f) => {
      const v = before[f];
      return v && typeof v === "object" && String(v.Label) === String(v.WssId);
    });
    if (corrupted.length === 0) {
      console.log("  ✓ Every raw field's Label already reads real text — nothing corrupted here right now.");
      continue;
    }
    console.log(`  ⚠ Corrupted (Label === WssId): ${corrupted.join(", ")}`);

    const formValues = corrupted.map((f) => {
      const v = before[f];
      // Re-send exactly the label|guid shape the field already carries once resolved via
      // FieldValuesAsText for the same field — this is the "correct" value, re-asserted.
      const label = beforeText.__error ? undefined : beforeText[f];
      return {
        FieldName: f,
        FieldValue: `${label ?? v.Label}|${v.TermGuid}`,
      };
    });

    console.log("  Would write (the exact re-assertion, via validateUpdateListItem):");
    for (const fv of formValues) console.log(`    ${fv.FieldName} = "${fv.FieldValue}"`);

    if (!DO_WRITE) {
      console.log("  (DO_WRITE is false — nothing written. Set DO_WRITE = true and re-run to test the repair.)");
      continue;
    }

    const digestValue = await getDigest();
    if (!digestValue) {
      console.log("  ✗ Could not get a request digest — cannot write. Nothing changed.");
      continue;
    }

    const writeRes = await fetch(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(found.library)}')/items(${found.id})/validateUpdateListItem`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json",
          "X-RequestDigest": digestValue,
        },
        body: JSON.stringify({ formValues, bNewDocumentUpdate: false }),
      },
    );
    if (!writeRes.ok) {
      console.log(`  ✗ Write failed: HTTP ${writeRes.status}`);
      continue;
    }
    const writeBody = await writeRes.json();
    const exceptions = (writeBody.value || []).filter((v) => v.HasException);
    if (exceptions.length > 0) {
      console.log("  ✗ Write reported field errors:", JSON.stringify(exceptions));
      continue;
    }
    console.log("  ✓ Write succeeded, no field exceptions reported. Re-reading...");

    const after = await readRawTax(found.library, found.id);
    if (after.__error) {
      console.log("  ✗ Could not re-read raw taxonomy fields after the write:", after.__error);
      continue;
    }
    const afterDesc = describeRaw(after);
    console.log("  AFTER — raw field:");
    for (const f of TAX_FIELDS) console.log(`    ${f}: ${afterDesc[f]}`);

    const stillCorrupted = corrupted.filter((f) => {
      const v = after[f];
      return v && typeof v === "object" && String(v.Label) === String(v.WssId);
    });
    if (stillCorrupted.length === 0) {
      console.log("  ✅ REPAIRED — every previously-corrupted field now shows a real Label.");
    } else {
      console.log(`  ✗ STILL CORRUPTED after a second identical write: ${stillCorrupted.join(", ")}`);
      console.log("    A second write does not fix this — the cache issue needs a different mitigation.");
    }
  }
})();

/*
 * diagnose-search-filter-gap.js — READ ONLY. Browser console, run AS THE ACCOUNT SEEING THE BUG
 * (clarence@trinergydigital.com), NOT as a site admin — the whole point is to see what THIS
 * account's own session gets back, since a different account's permissions would not reproduce it.
 *
 * WHY THIS EXISTS
 * ----------------
 * CRS Search's Advanced Filters were fixed 2026-09-21 (client-side narrowing over a follow-up REST
 * read, replacing exact KQL clauses against managed properties that are always empty on this
 * tenant — see CLAUDE.md, "CRS SEARCH'S METADATA FILTERS ARE SOLVED"). Deployed, confirmed present
 * in the shipped .sppkg, hard-refreshed — and a search for "fruit" + Document type=Working File +
 * Year=2024 + Confidentiality=Restricted + Segment=Group Head Office still answers "No Result
 * Found", even though the exact document (Fruit Juice-Tealive-Orange-01092026.pdf, in the Archive
 * library) demonstrably carries all four of those values and free text alone finds it.
 *
 * Two things could be true, and only a live run tells them apart:
 *   1. The Search API itself does not return the hit once free text runs alongside the OTHER
 *      criteria present in the request (a scope/word-clause interaction nobody has checked), or
 *   2. Search DOES return the hit, and the follow-up REST metadata read (the new narrowing step)
 *      either fails outright, comes back with no matching row, or comes back with values that do
 *      not compare equal to what is on screen.
 *
 * This script reproduces BOTH steps exactly as the shipped code does them and prints the raw
 * result of each, so the failure point is read off the console rather than guessed.
 *
 * HOW TO RUN
 *   1. Sign in to the site AS clarence@trinergydigital.com (the account reporting the bug).
 *   2. Open CRS Search's own page (or any page on the site) so the session cookie is live.
 *   3. Edit the CONFIG block below if you want to try a different word/file.
 *   4. F12 -> Console -> paste this whole file -> Enter.
 *   5. Read the printed JSON against the checklist at the bottom of this file.
 *
 * ⚠ EVERY REQUEST IS A GET. It creates nothing, changes nothing, deletes nothing.
 */
(async () => {
  /* ── Configuration — edit if testing a different document ─────────────────────────────── */
  const CONFIG = {
    word: "fruit",
    documentType: "Working File",
    year: "2024",
    confidentiality: "Restricted",
    segment: "Group Head Office",
  };

  /* ── Resolve the site — never rely on `_spPageContextInfo` alone, it is undefined on several
     modern pages and in any non-top frame. Same three-route fallback every script here uses. */
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const resolveWeb = () => {
    try {
      if (
        typeof _spPageContextInfo !== "undefined" &&
        _spPageContextInfo.webAbsoluteUrl
      ) {
        return _spPageContextInfo.webAbsoluteUrl.replace(/\/+$/, "");
      }
    } catch (e) {
      /* fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    return (location.origin + (m ? m[1] : "")).replace(/\/+$/, "");
  };
  const web = SITE_OVERRIDE || resolveWeb();
  console.log("Site:", web);

  const bust = () => `_=${Date.now()}${Math.random().toString(36).slice(2)}`;

  const get = async (url, extraHeaders) => {
    const sep = url.indexOf("?") > -1 ? "&" : "?";
    try {
      const r = await fetch(`${url}${sep}${bust()}`, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Cache-Control": "no-cache",
          Pragma: "no-cache",
          ...(extraHeaders || {}),
        },
        cache: "no-store",
      });
      const status = r.status;
      let body;
      try {
        body = await r.json();
      } catch (e) {
        body = { __unparseable: true };
      }
      return { ok: r.ok, status, body };
    } catch (e) {
      return { ok: false, status: 0, body: { __error: String((e && e.message) || e) } };
    }
  };

  const textOf = (v) => {
    if (v === undefined || v === null) return "";
    if (typeof v === "string") return v.trim();
    if (typeof v === "number" || typeof v === "boolean") return String(v);
    if (typeof v === "object") {
      if (typeof v.Label === "string") return v.Label.trim();
      if (typeof v.Title === "string") return v.Title.trim();
    }
    return "";
  };

  /* ── Resolve every approved-side library by URL SEGMENT (never title — a rename never touches
     the URL, and this client renames libraries routinely). Same authoritative read this project's
     own diagnose-blank-libraries.js already uses. ─────────────────────────────────────────────── */
  const libsRead = await get(
    `${web}/_api/web/lists?$select=Title,ItemCount&$expand=RootFolder&$filter=BaseTemplate eq 101`,
  );
  const libs = libsRead.ok ? libsRead.body.value || [] : [];
  const bySeg = (seg) => {
    const found = libs.find((l) => {
      const url = ((l.RootFolder || {}).ServerRelativeUrl || "").toLowerCase();
      const parts = url.split("/").filter((p) => p.length > 0);
      const tail = parts.length > 0 ? parts[parts.length - 1] : "";
      return tail === seg.toLowerCase();
    });
    return found ? found.Title : undefined;
  };
  const archiveTitle = bySeg("Archive");
  const archiveHcTitle = bySeg("HCArchive");
  const docsTitle = bySeg("Shared Documents") || bySeg("Documents");
  const docsHcTitle = bySeg("HCDocuments");
  console.log("Resolved libraries:", {
    archiveTitle,
    archiveHcTitle,
    docsTitle,
    docsHcTitle,
  });
  if (!libsRead.ok) {
    console.error(
      "Could not read the library list at all (HTTP",
      libsRead.status,
      ") — everything below will be empty. This alone would explain the bug: if this account",
      "cannot even enumerate libraries, it certainly cannot run the follow-up REST narrowing.",
    );
  }

  /* ── STEP -1: WHERE DID THE BAD LABEL COME FROM — Form.tsx's own payload, or the flow that
     applies it? Added after confirming (a) the field really is TaxonomyFieldType, (b) $expand is
     not valid syntax here, and (c) ORDINARY recent documents on this site ALSO carry this exact
     defect (Label literally equals the numeric WssId as a string) — so this is not one bad test
     file, it is systemic and recent. Since 2026-09-18, an upload does not write these fields
     directly — it stages a `TagPayload` JSON on `CRS Submissions`, and a separate Power Automate
     flow (`CRS — Apply pending tags`) applies it asynchronously, running as the service account.
     `Form.tsx`'s own payload-building code (`toTaxValue`, `loadTermSet`) reads correctly on
     inspection — so the two live possibilities are: the PAYLOAD itself already has the wrong label
     (a code bug, fixable here), or the payload is correct and the FLOW mis-applies it (Power
     Automate, not fixable from this codebase). Reading the stored payload directly settles which. */
  {
    console.log("\n── STEP -1: checking the staged TagPayload behind this document's write ──");
    const subCandidates = ["CRS Submissions", "DMS Submissions", "GDC Submissions"];
    let subsTitle;
    for (const cand of subCandidates) {
      const probe = await get(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(cand)}')?$select=Title`,
      );
      if (probe.ok) {
        subsTitle = cand;
        break;
      }
    }
    console.log("  Resolved submissions list:", subsTitle || "(not found under any candidate name)");
    if (subsTitle && archiveTitle) {
      // Read the document's own stamp first — this is what joins it to its CRS Submissions row.
      const stampRead = await get(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(archiveTitle)}')/items` +
          `?$select=Id,SubmissionFileId&$filter=${encodeURIComponent(`FSObjType eq 0 and substringof('${CONFIG.word}',FileLeafRef)`)}&$top=1`,
      );
      const stampRow = stampRead.ok ? (stampRead.body.value || [])[0] : undefined;
      const sfi = stampRow ? textOf(stampRow.SubmissionFileId) : "";
      console.log("  Document's own SubmissionFileId stamp:", sfi || "(none — this document predates the tag/approve-by-proxy migration, or was never stamped)");
      if (sfi) {
        const rec = await get(
          `${web}/_api/web/lists/getbytitle('${encodeURIComponent(subsTitle)}')/items` +
            `?$select=Id,SubmissionFileId,TagStatus,TagError,TagPayload,LibraryTitle` +
            `&$filter=${encodeURIComponent(`SubmissionFileId eq '${sfi}'`)}&$top=1`,
        );
        if (!rec.ok) {
          console.log("  Could not read the submission record — status:", rec.status, rec.body);
        } else {
          const recRow = (rec.body.value || [])[0];
          if (!recRow) {
            console.log("  No CRS Submissions row found for this stamp.");
          } else {
            console.log("  TagStatus:", recRow.TagStatus, " TagError:", recRow.TagError);
            let payload;
            try {
              payload = JSON.parse(recRow.TagPayload || "[]");
            } catch (e) {
              payload = { __unparseable: recRow.TagPayload };
            }
            console.log("  Staged TagPayload (what Form.tsx BUILT, before the flow applied it):", payload);
            const dtEntry = Array.isArray(payload)
              ? payload.find((p) => p && (p.FieldName === "Document_x0020_Type" || p.FieldName === "Year"))
              : undefined;
            if (dtEntry) {
              const looksBad = /^\d+\|/.test(dtEntry.FieldValue || "");
              console.log(
                looksBad
                  ? "  CONCLUSION: the PAYLOAD ITSELF already has a numeric label (e.g. '14|<guid>' instead of " +
                      "'Working File|<guid>') — this is a bug in the code that BUILT the payload " +
                      "(Form.tsx/BulkUpload.tsx), not in the Power Automate flow that applied it."
                  : "  CONCLUSION: the payload looks CORRECT (a real label before the '|') — so the flow that " +
                      "APPLIED it (CRS — Apply pending tags) is what wrote the wrong value. That is Power " +
                      "Automate, not something fixable in this codebase — needs opening that flow directly.",
              );
            } else {
              console.log(
                "  Could not find a Document_x0020_Type/Year entry in the payload to check — read the",
                "printed payload above directly.",
              );
            }
          }
        }
      }
    }
  }

  /* ── STEP 0: SKIP THE SEARCH CRAWL ENTIRELY — added after the first run of this script proved
     Search returns ZERO hits for "fruit" across all four libraries, even unfiltered. That means the
     document was never found via the crawl at all — it was found via the SEPARATE 24-hour "recently
     modified" REST fallback (`runListRead`/`buildRecentFilter` in DocumentSearch.tsx), which exists
     precisely because a brand-new or just-modified document is not yet crawled. THIS script's own
     Step 1/2 tested the wrong mechanism. This step finds the item directly by filename via REST
     (bypassing Search) on the Archive library, then reproduces the recency-top-up's OWN filter —
     including the SERVER-SIDE `Business_x0020_Segment eq '...'` clause that mechanism has used since
     before this session and which this session never touched — to see whether THAT is what breaks. */
  if (archiveTitle) {
    console.log("\n── STEP 0: finding the item directly by filename on", archiveTitle, "──");
    const base = `${web}/_api/web/lists/getbytitle('${encodeURIComponent(archiveTitle)}')/items`;
    const cols = "Id,FileLeafRef,Modified,Document_x0020_Type,Year,Confidentiality_x0020_Level,Business_x0020_Segment";

    const baseline = await get(
      `${base}?$select=${cols}&$filter=${encodeURIComponent(`FSObjType eq 0 and substringof('${CONFIG.word}',FileLeafRef)`)}&$top=5`,
    );
    console.log("  Filename-only lookup status:", baseline.status, baseline.ok ? "(ok)" : "(FAILED)");
    if (!baseline.ok) {
      console.log("  Body:", baseline.body);
    } else {
      const rows = baseline.body.value || [];
      console.log(`  Found ${rows.length} item(s) by filename alone (no other filter):`);
      rows.forEach((row) => {
        console.log("   ", {
          Id: row.Id,
          FileLeafRef: row.FileLeafRef,
          Modified: row.Modified,
          documentType: textOf(row.Document_x0020_Type),
          year: textOf(row.Year),
          confidentiality: textOf(row.Confidentiality_x0020_Level),
          segmentRAW: JSON.stringify(row.Business_x0020_Segment),
        });
      });

      /* ⚠⚠ COMPARISON CHECK, added after `$expand` returned 400 — $expand is not valid syntax for a
         taxonomy field at all (it only applies to genuine Lookup/Person navigable properties), so
         that theory is dead. A bare `$select` is NORMALLY expected to already resolve a taxonomy
         field to `{Label, TermGuid, WssId}` with no $expand needed — so getting a raw number back
         instead points at THIS DOCUMENT's stored value being malformed (e.g. written as a bare WssId
         rather than this project's required `Label|TermGuid` format — gotcha #5), not at the read
         approach being wrong. This pulls a few ORDINARY documents from the live (non-archive) library
         the same way, to see whether the SAME raw-number behaviour happens there too. */
      const compareLib = docsTitle || archiveTitle;
      if (compareLib) {
        console.log(`\n  ── Comparison: same bare $select against ordinary documents in ${compareLib} ──`);
        const cmp = await get(
          `${web}/_api/web/lists/getbytitle('${encodeURIComponent(compareLib)}')/items` +
            `?$select=Id,FileLeafRef,Document_x0020_Type,Year&$filter=${encodeURIComponent("FSObjType eq 0")}&$top=3`,
        );
        if (!cmp.ok) {
          console.log("  Comparison read status:", cmp.status, cmp.body);
        } else {
          const cmpRows = cmp.body.value || [];
          console.log(`  Found ${cmpRows.length} ordinary document(s) to compare against:`);
          cmpRows.forEach((r) => {
            console.log("   ", {
              Id: r.Id,
              FileLeafRef: r.FileLeafRef,
              documentTypeRAW: JSON.stringify(r.Document_x0020_Type),
              yearRAW: JSON.stringify(r.Year),
            });
          });
          const anyResolved = cmpRows.some(
            (r) => typeof r.Document_x0020_Type === "object" && r.Document_x0020_Type,
          );
          const anyBareNumber = cmpRows.some((r) => typeof r.Document_x0020_Type === "number");
          if (cmpRows.length === 0) {
            console.log("  (no ordinary documents found to compare against — inconclusive)");
          } else if (anyResolved && !anyBareNumber) {
            console.log(
              "  CONCLUSION: ordinary documents DO resolve to {Label,...} via the exact same bare",
              "$select this script and the shipped code use. That means the read approach is correct",
              "in general, and the '14'/'3' seen on the Archive test file is specific to THAT",
              "document's own stored value — most likely written directly as a bare number rather than",
              "the required 'Label|GUID' format, bypassing this project's normal upload/tagging path.",
              "Worth checking how that specific test file's metadata was set.",
            );
          } else if (anyBareNumber) {
            console.log(
              "  CONCLUSION: ordinary documents ALSO come back as raw numbers via this exact read —",
              "this is a genuine, systemic gap in the narrowing code's read approach, not one bad",
              "test file. The fix needs to resolve these numbers to labels some other way (e.g. by",
              "cross-referencing the term store, since $expand is not valid here).",
            );
          } else {
            console.log("  (mixed/unclear result — read the raw values printed above directly)");
          }
        }
      }

      /* ⚠ "documentType: '14', year: '3'" — RAW NUMERIC IDs, not the `{Label, TermGuid, WssId}`
         object `textOf` expects for a taxonomy field. Two things could explain it, and only a live
         read tells them apart: (a) a bare `$select` on a taxonomy field genuinely needs `$expand` to
         resolve to the label object — a real SharePoint REST quirk `metadataFilterMatches` may have
         always been exposed to and never hit until now, or (b) `Document_x0020_Type`/`Year` on THIS
         library are not actually TaxonomyFieldType at all despite CLAUDE.md's 2026-09-10 record
         saying otherwise. Both checks below settle it directly. */
      if (rows.length > 0) {
        const oneId = rows[0].Id;
        console.log(`\n  ── Checking the taxonomy read for item ${oneId} ──`);
        const schema = await get(
          `${web}/_api/web/lists/getbytitle('${encodeURIComponent(archiveTitle)}')/fields` +
            `?$select=InternalName,TypeAsString&$filter=${encodeURIComponent(
              "InternalName eq 'Document_x0020_Type' or InternalName eq 'Year' or InternalName eq 'Confidentiality_x0020_Level'",
            )}`,
        );
        console.log(
          "  Field types on this library:",
          schema.ok ? (schema.body.value || []).map((f) => `${f.InternalName}=${f.TypeAsString}`) : schema,
        );

        const expanded = await get(
          `${web}/_api/web/lists/getbytitle('${encodeURIComponent(archiveTitle)}')/items(${oneId})` +
            `?$select=Id,Document_x0020_Type/Label,Document_x0020_Type/TermGuid,Year/Label,Year/TermGuid,` +
            `Confidentiality_x0020_Level/Label,Confidentiality_x0020_Level/TermGuid` +
            `&$expand=Document_x0020_Type,Year,Confidentiality_x0020_Level`,
        );
        console.log(
          "  SAME item read WITH $expand — status:",
          expanded.status,
          expanded.ok ? "(ok)" : "(FAILED)",
        );
        if (expanded.ok) {
          console.log("  Values with $expand:", {
            documentType: textOf(expanded.body.Document_x0020_Type),
            year: textOf(expanded.body.Year),
            confidentiality: textOf(expanded.body.Confidentiality_x0020_Level),
          });
          const nowResolved =
            textOf(expanded.body.Document_x0020_Type) !== "14" ||
            textOf(expanded.body.Year) !== "3";
          if (nowResolved) {
            console.log(
              "  CONCLUSION: `$expand` resolves the real labels where a bare `$select` gave raw",
              "term IDs — this is a genuine SharePoint REST behaviour the narrowing code never",
              "accounted for. The fix is adding `$expand` to the follow-up metadata reads.",
            );
          } else {
            console.log(
              "  CONCLUSION: even $expand still gives a raw id — this points at the FIELD SCHEMA",
              "itself rather than the query shape. Check the field-types line printed above.",
            );
          }
        } else {
          console.log("  Body:", expanded.body);
        }
      }

      if (rows.length > 0) {
        // Now add EXACTLY the clause the recency top-up already adds server-side, unrelated to
        // anything this session changed — a plain, case-sensitive, untrimmed OData `eq`.
        const withSegment = await get(
          `${base}?$select=${cols}&$filter=${encodeURIComponent(
            `FSObjType eq 0 and substringof('${CONFIG.word}',FileLeafRef) and Business_x0020_Segment eq '${CONFIG.segment}'`,
          )}&$top=5`,
        );
        console.log(
          "  SAME lookup WITH the segment eq clause added — status:",
          withSegment.status,
          withSegment.ok ? "(ok)" : "(FAILED)",
        );
        if (withSegment.ok) {
          const rows2 = withSegment.body.value || [];
          console.log(`  Found ${rows2.length} item(s) once 'Business_x0020_Segment eq \"${CONFIG.segment}\"' is added.`);
          if (rows.length > 0 && rows2.length === 0) {
            console.log(
              "  CONCLUSION: the item exists and is found by filename alone, but disappears the",
              "moment the segment `eq` filter is added — that clause does not match this item's",
              "actual stored value. Compare the `segmentRAW` value printed above, character by",
              "character, against the exact text `" + CONFIG.segment + "` — a stray trailing space,",
              "a different label, or a managed-metadata shape instead of plain text would all cause",
              "this silently (OData `eq` is exact and case-sensitive, unlike this app's own",
              "client-side comparisons everywhere else).",
            );
          }
        } else {
          console.log("  Body:", withSegment.body);
        }
      } else {
        console.log(
          "  CONCLUSION: the document cannot even be found by filename alone via REST on this",
          "library for this account — check the Modified date and whether it is really within the",
          "last 24 hours, and whether this account can read this library via REST at all.",
        );
      }
    }
  }

  /* ── STEP 1: run the SAME KQL query the shipped `buildKql` builds — scope + IsDocument:true +
     the word clause — WITHOUT any exact-match clause for the four Advanced Filters, exactly as
     the current code does it (2026-09-21). If this returns zero hits, the bug is upstream of the
     new narrowing code entirely. ───────────────────────────────────────────────────────────── */
  const segs = [docsTitle, docsHcTitle, archiveTitle, archiveHcTitle]
    .filter((t) => t)
    .map((t) => {
      const lib = libs.find((l) => l.Title === t);
      const url = ((lib || {}).RootFolder || {}).ServerRelativeUrl || "";
      return url.split("/").filter((p) => p.length > 0).pop() || "";
    })
    .filter((s) => s.length > 0);
  const pathClauses = segs.map((s) => `Path:"${web}/${s}/*"`);
  const scope = pathClauses.length === 1 ? pathClauses[0] : `(${pathClauses.join(" OR ")})`;
  const query = `${scope} IsDocument:true (${CONFIG.word}* OR ProjectNameOWSTEXT:${CONFIG.word}* OR Vendor_x002f_CustomerNameOWSTEXT:${CONFIG.word}* OR KeywordOWSTEXT:${CONFIG.word}*)`;
  console.log("KQL query text (no exact clauses, matching the shipped code):", query);

  const props = "Path,Filename,Title,Created,CreatedBy,Author,Size,UniqueId,ListItemID";
  const searchResp = await get(
    `${web}/_api/search/query?querytext='${encodeURIComponent(query)}'` +
      `&rowlimit=200&trimduplicates=false&selectproperties='${encodeURIComponent(props)}'`,
    { "odata-version": "" },
  );
  if (!searchResp.ok) {
    console.error("STEP 1 FAILED — the Search API call itself did not succeed:", searchResp.status, searchResp.body);
    console.log("CONCLUSION: the bug is at the Search request itself, before narrowing ever runs.");
    return;
  }
  const table =
    (((searchResp.body.PrimaryQueryResult || {}).RelevantResults || {}).Table || {}).Rows || [];
  const hits = table.map((row) => {
    const map = {};
    (row.Cells || []).forEach((c) => {
      map[textOf(c.Key)] = textOf(c.Value);
    });
    return map;
  });
  console.log(`STEP 1: Search returned ${hits.length} hit(s).`);
  hits.forEach((h, i) =>
    console.log(`  [${i}] Path=${h.Path}  ListItemID=${h.ListItemID}  UniqueId=${h.UniqueId}`),
  );

  if (hits.length === 0) {
    console.log(
      "CONCLUSION: Search itself returns ZERO hits for this word + scope, with NO exact-match",
      "clauses in the query at all. That means the bug is not in the narrowing code this session",
      "added — it is either the scope (these libraries are not actually in it for this account),",
      "or the free-text clause itself is not matching this document for some other reason.",
    );
    return;
  }

  /* ── STEP 2: for each hit, resolve which library it is in (by path segment) and run the SAME
     follow-up REST read the shipped narrowing code runs — print the raw status and body. ────── */
  const classify = (path) => {
    const p = (path || "").toLowerCase();
    const has = (seg) => seg && p.indexOf(`/${seg.toLowerCase()}/`) !== -1;
    if (archiveHcTitle && has("hcarchive")) return { key: "ArchiveHC", title: archiveHcTitle };
    if (archiveTitle && has("archive")) return { key: "Archive", title: archiveTitle };
    if (docsHcTitle && has("hcdocuments")) return { key: "DocumentsHC", title: docsHcTitle };
    return { key: "Documents", title: docsTitle };
  };

  for (const h of hits) {
    const lib = classify(h.Path);
    const itemId = Number(h.ListItemID) || 0;
    console.log(`\nSTEP 2 for ${h.Path}  ->  classified as ${lib.key} (${lib.title})`);
    if (!lib.title) {
      console.log("  Could not resolve a library title for this hit — cannot run the follow-up read.");
      continue;
    }
    if (itemId <= 0) {
      console.log("  ListItemID is missing or zero — the follow-up read cannot target this item.");
      continue;
    }
    const metaResp = await get(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(lib.title)}')/items` +
        `?$select=Id,Document_x0020_Type,Year,Confidentiality_x0020_Level,Business_x0020_Segment` +
        `&$filter=${encodeURIComponent(`Id eq ${itemId}`)}&$top=1`,
    );
    console.log("  Follow-up REST read status:", metaResp.status, metaResp.ok ? "(ok)" : "(FAILED)");
    if (!metaResp.ok) {
      console.log(
        "  Body:",
        metaResp.body,
        "\n  CONCLUSION for this hit: the follow-up read is REFUSED — the shipped code should keep",
        "this hit unnarrowed rather than drop it, so if it is still missing from results, check",
        "whether the 'could not be filtered' banner appeared instead of a flat empty state.",
      );
      continue;
    }
    const rows = metaResp.body.value || [];
    if (rows.length === 0) {
      console.log(
        "  Body: zero rows returned for `Id eq",
        itemId,
        "` against",
        lib.title,
        "\n  CONCLUSION for this hit: the item id Search reported does NOT resolve in this library",
        "via REST — either an id mismatch (crawl lag after a move) or a wrong library resolved.",
      );
      continue;
    }
    const row = rows[0];
    const got = {
      documentType: textOf(row.Document_x0020_Type),
      year: textOf(row.Year),
      confidentiality: textOf(row.Confidentiality_x0020_Level),
      segment: textOf(row.Business_x0020_Segment),
    };
    console.log("  Values read back:", got);
    const same = (want, gotVal) => {
      const w = (want || "").trim().toLowerCase();
      if (w.length === 0) return true;
      return (gotVal || "").trim().toLowerCase() === w;
    };
    const matches =
      same(CONFIG.documentType, got.documentType) &&
      same(CONFIG.year, got.year) &&
      same(CONFIG.confidentiality, got.confidentiality) &&
      same(CONFIG.segment, got.segment);
    console.log("  Matches the configured filters?", matches);
    if (!matches) {
      console.log(
        "  CONCLUSION for this hit: the follow-up read SUCCEEDED and returned real values, but",
        "they do not equal what was typed above — compare each field printed against CONFIG at",
        "the top of this script character by character (whitespace, casing already tolerated).",
      );
    } else {
      console.log(
        "  This hit SHOULD have been kept and shown. If it was not, something between this point",
        "and the render is wrong rather than the metadata check itself — worth a second look.",
      );
    }
  }
})().catch((e) => console.error("Script error:", e));

/*
 * READING THE OUTPUT
 * -------------------
 * - "STEP 1: Search returned 0 hit(s)" -> the bug is upstream of narrowing entirely (the query or
 *   its scope), not the fix this session made. Report the printed query text back.
 * - A hit appears in STEP 1 but its STEP 2 follow-up read status is NOT `ok` -> the account cannot
 *   read that library via a list-level `$filter` query even though item-level reads (e.g. opening
 *   the detail panel) succeeded — a genuinely surprising SharePoint permission split worth reporting
 *   with the exact HTTP status.
 * - STEP 2 succeeds but returns 0 rows -> the item id from Search does not exist in that library via
 *   REST right now — most likely explanation is the document was recently moved (e.g. into Archive)
 *   and the crawled `ListItemID` property has not caught up with the item's current numbering.
 * - STEP 2 succeeds, returns a row, but the printed values do not match CONFIG -> a real data or
 *   comparison mismatch — report the exact printed values.
 */

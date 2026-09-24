/*
 * Is the raw-Label-corruption per DOCUMENT, or per TERM (site-wide cache entry)?
 * -------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written.
 *
 * WHY THIS EXISTS
 * `check-taxonomy-relabel.js` found that a document's raw `Document_x0020_Type`/`Year`/
 * `Confidentiality_x0020_Level` fields have `Label === WssId` (e.g. "8" instead of "Term Sheet"),
 * while `FieldValuesAsText` resolves the same fields correctly — and that RE-WRITING the correct
 * value a second time via `validateUpdateListItem` does NOT fix it. That rules out "the write is
 * flaky" and points at SharePoint's own `TaxonomyHiddenList` (a per-site term cache) holding a
 * corrupted entry for these specific term GUIDs — which every write resolves THROUGH, so what you
 * send in the request may not matter once the cache entry itself is bad.
 *
 * IF THAT IS RIGHT, THE CORRUPTION IS KEYED ON (site, TermGuid) — NOT ON WHICH DOCUMENT OR WHICH
 * FLOW WROTE IT. This script tests that directly: it reads a batch of OTHER, unrelated documents
 * across the approved-side libraries and checks whether any of them ALSO carry one of the three
 * specific term GUIDs already confirmed corrupt on the original test document — and if so, whether
 * THEY show the identical Label===WssId corruption too.
 *
 * TWO POSSIBLE OUTCOMES, and they point at very different next steps:
 *   - Other documents using the SAME term GUIDs are ALSO corrupted → confirms a persistent,
 *     site-wide cache corruption for these specific terms, independent of which document or which
 *     write path touched them. Likely affects every document ever tagged with these three terms,
 *     not just recent bulk-upload/proxy-tagged ones — worth a full scope count once confirmed.
 *   - Other documents using the SAME term GUIDs are CORRECT (raw Label shows real text) → the
 *     corruption is narrower than a permanent per-term cache entry; something else about the
 *     specific write that produced THIS document's corruption is the real cause, and the theory
 *     needs revisiting rather than assuming every document sharing these terms is affected.
 *
 * HOW TO RUN
 *   F12 → Console on the CRS site, signed in as a site administrator. Paste the whole file, Enter.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────────────────────── */
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"

  // The three term GUIDs already confirmed corrupt on SFI-20260924-JNXK (item #1677, Restricted &
  // Confidential Document). Checking every OTHER document that also carries one of these, to see
  // whether the corruption is specific to that one document or shared by every document using
  // these exact terms.
  const WATCH_TERMS = {
    Document_x0020_Type: {
      guid: "ffd9961d-f70d-4044-b5b9-a404a1d50f3d",
      label: "Term Sheet",
    },
    Year: {
      guid: "5d081550-e418-4f3a-9f78-1cb994cc3ca6",
      label: "2024",
    },
    Confidentiality_x0020_Level: {
      guid: "248d979e-73f4-4639-90ca-8f67d6cabf22",
      label: "Confidential",
    },
  };

  // How many recent items to read per library, per taxonomy field, before giving up on that
  // combination. Kept modest — this is a scoping test, not an exhaustive site-wide sweep.
  const PER_FIELD_SCAN = 100;

  const LIBRARY_CANDIDATES = [
    "Restricted & Confidential Document",
    "Documents",
    "Highly Confidential Document",
    "HC Documents",
  ];

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
  console.log("Site:", web);

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

  /* ── Which of the LIBRARY_CANDIDATES actually resolve on this site ──────────────────────── */
  const libraries = [];
  for (const title of LIBRARY_CANDIDATES) {
    const probe = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Title,ItemCount`,
    );
    if (probe && probe.Title) libraries.push(probe.Title);
  }
  console.log("Checking libraries:", libraries.join(", ") || "(none resolved)");

  /* ── For each taxonomy field, scan each library for other documents carrying the watched
     term GUID, and report whether their raw Label is corrupted too ───────────────────────── */
  let totalChecked = 0;
  let totalMatchingTerm = 0;
  let totalCorrupted = 0;
  let totalCorrect = 0;

  for (const [field, watch] of Object.entries(WATCH_TERMS)) {
    console.log(`\n=== ${field} — watching term "${watch.label}" (${watch.guid}) ===`);
    for (const library of libraries) {
      const res = await tryGet(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(library)}')/items` +
          `?$select=Id,FileLeafRef,${field}&$top=${PER_FIELD_SCAN}&$orderby=Id desc`,
      );
      if (res.__error) {
        console.log(`  ✗ [${library}] could not read: ${res.__error}`);
        continue;
      }
      const rows = res.value || [];
      totalChecked += rows.length;
      let libMatching = 0;
      let libCorrupted = 0;
      let libCorrect = 0;
      for (const r of rows) {
        const v = r[field];
        if (!v || typeof v !== "object" || v.TermGuid !== watch.guid) continue;
        libMatching += 1;
        totalMatchingTerm += 1;
        const isCorrupted = String(v.Label) === String(v.WssId);
        if (isCorrupted) {
          libCorrupted += 1;
          totalCorrupted += 1;
        } else {
          libCorrect += 1;
          totalCorrect += 1;
        }
        console.log(
          `    #${r.Id} "${r.FileLeafRef}" — Label="${v.Label}" WssId=${v.WssId} ` +
            (isCorrupted ? "⚠ CORRUPTED" : "✓ correct"),
        );
      }
      if (libMatching === 0) {
        console.log(
          `  [${library}] ${rows.length} item(s) scanned, none carry this exact term GUID in the top ${PER_FIELD_SCAN}.`,
        );
      } else {
        console.log(
          `  [${library}] ${libMatching} document(s) carry this term: ${libCorrupted} corrupted, ${libCorrect} correct.`,
        );
      }
    }
  }

  /* ── Summary ──────────────────────────────────────────────────────────────────────────── */
  console.log(`\n=== SUMMARY ===`);
  console.log(`Items scanned across all libraries/fields: ${totalChecked}`);
  console.log(`Documents found carrying one of the three watched terms: ${totalMatchingTerm}`);
  console.log(`  Corrupted (Label === WssId): ${totalCorrupted}`);
  console.log(`  Correct (Label is real text): ${totalCorrect}`);
  if (totalMatchingTerm === 0) {
    console.log(
      "\nNo other documents in the scanned window carry these exact term GUIDs — cannot confirm " +
        "or rule out the per-term theory from this sample. Try raising PER_FIELD_SCAN, or this may " +
        "genuinely be the only document using these particular terms so far.",
    );
  } else if (totalCorrupted > 0 && totalCorrect > 0) {
    console.log(
      "\n⚠ MIXED — some documents sharing these term GUIDs are corrupted, others are not. This is " +
        "NOT a simple 'every document with this term is broken' cache corruption — something else " +
        "distinguishes the corrupted ones (worth checking: upload date, which flow/path wrote them).",
    );
  } else if (totalCorrupted > 0) {
    console.log(
      "\n⚠⚠ EVERY document found carrying these term GUIDs is corrupted — consistent with a " +
        "persistent, site-wide TaxonomyHiddenList cache corruption for these specific terms, not " +
        "something specific to one document or one write. Worth a full scope count next.",
    );
  } else {
    console.log(
      "\n✓ Every OTHER document found carrying these term GUIDs is correct — the corruption does " +
        "NOT appear to be a blanket per-term cache issue. Something specific to the original test " +
        "document's write is the more likely cause; the per-term cache theory needs revisiting.",
    );
  }
})();

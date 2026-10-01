/**
 * Seeds the Term Abbreviation list for a segment whose tiers are NOT
 * Department/Unit — written for `12-upstream-operations-abbreviations.csv`
 * (Region → Estate/Mill), but generic over the tier columns.
 *
 * ── Why not seed-term-abbreviations.js ───────────────────────────────────────
 * That script branches on the LITERAL level names: `r.Level === "Department"`
 * takes two path parts, anything else takes three. A Region row therefore falls
 * into the three-part branch, builds "segment|region|" with an empty last part
 * and resolves nothing — while its Estate/Mill rows happen to work. Half the
 * rows would silently go missing.
 *
 * This one builds the label path from whichever TIER_COLUMNS are filled, so the
 * `Level` value is written through as data rather than steering the logic. A
 * future segment of any shape works by editing TIER_COLUMNS.
 *
 * JavaScript rather than PowerShell for the same reason as the original: PS does
 * not run reliably in this environment.
 *
 * ── How to run ───────────────────────────────────────────────────────────────
 * 1. Import `11-upstream-operations-malaysia.csv` into the term store FIRST.
 *    Nothing here can resolve a term that does not exist yet.
 * 2. Copy the new term set's ID and paste it into SETS below.
 * 3. Paste the contents of `12-upstream-operations-abbreviations.csv` into
 *    CSV_TEXT below, between the backticks.
 * 4. Open any page of the site in a browser, signed in as someone who can edit
 *    the list, and paste this whole file into the DevTools console.
 *
 * Re-runnable: a row already holding the same TermGuid is UPDATED, never
 * duplicated — so a corrected abbreviation is a re-run, not a cleanup job.
 *
 * The `Segment` level row is skipped deliberately: a segment container has no
 * term, so its folder name comes from `StagingFolder` on the DMS Config mode row
 * (the "Top folder name" field on the New segment tab) instead.
 */
const SITE = "/sites/ClarenceDMSTesting";

/** Tried in order — the list is `CRS Term Abbreviation` on a renamed site. */
const LIST_CANDIDATES = ["CRS Term Abbreviation", "DMS Term Abbreviation"];

/** Segment name (must match the CSV's Segment column) -> its term set ID. */
const SETS = {
  "Upstream Operations Malaysia": "PASTE-THE-TERM-SET-ID-HERE",
};

/**
 * The CSV columns that make up a term's path, shallowest first. The segment name
 * is the root of the walk, then one column per tier.
 */
const TIER_COLUMNS = ["Region", "EstateOrMill"];

const CSV_TEXT = `PASTE THE CONTENTS OF 12-upstream-operations-abbreviations.csv HERE`;

/**
 * Folds the differences that are presentation, not identity, so a CSV label
 * matches its live term.
 *
 * The fullwidth ＆ is the one that actually bit on the head-office run: the term
 * store requires it, spreadsheets use plain &, and the first attempt matched only
 * 102 of 175 rows. A parent mismatch takes its children down with it.
 */
function normalizeLabel(s) {
  return (s || "")
    .replace(/＆/g, "&")
    .replace(/[‘’ʼ]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[‐‑‒–—―]/g, "-")
    .replace(/ /g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

const getJson = async (url) => {
  const r = await fetch(url, { headers: { Accept: "application/json;odata=nometadata" } });
  if (!r.ok) throw new Error(`HTTP ${r.status} on ${url}: ${await r.text()}`);
  return r.json();
};

/** A naive split on "," is wrong: several labels contain commas and are quoted. */
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') inQuotes = false;
      else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (c !== "\r") field += c;
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  const head = rows.shift();
  return rows
    .filter((r) => r.length === head.length)
    .map((r) => {
      const o = {};
      head.forEach((h, i) => (o[h.trim()] = r[i]));
      return o;
    });
}

(async () => {
  if (Object.keys(SETS).some((k) => SETS[k].indexOf("PASTE") === 0)) {
    console.error("Paste the term set ID into SETS first.");
    return;
  }
  if (CSV_TEXT.indexOf("PASTE") === 0) {
    console.error("Paste the CSV contents into CSV_TEXT first.");
    return;
  }

  // Resolve the list name before anything else: a 404 here is the difference
  // between "nothing was written" and "wrote 0 rows", and only one of those sends
  // you to look at the right thing.
  let LIST = "";
  for (const candidate of LIST_CANDIDATES) {
    try {
      await getJson(
        `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(candidate)}')?$select=Title`,
      );
      LIST = candidate;
      break;
    } catch (e) {
      /* try the next name */
    }
  }
  if (!LIST) {
    console.error(`None of these lists exist on ${SITE}: ${LIST_CANDIDATES.join(", ")}`);
    return;
  }
  console.log(`using list "${LIST}"`);

  // 1. Walk every configured term set, building "segment|tier1[|tier2…]" -> term.
  const guidByPath = new Map();
  for (const segName of Object.keys(SETS)) {
    const setId = SETS[segName];
    const walk = async (url, prefix) => {
      const d = await getJson(url);
      for (const t of d.value || []) {
        const label = t.labels[0].name;
        const path = `${prefix}|${normalizeLabel(label)}`;
        // The LIVE label, not the CSV's, becomes Title — and Title is what
        // reconciliation writes to each folder's Full Name column. The client must
        // see the same text there as in the upload dropdown.
        guidByPath.set(path, { id: t.id, label: label });
        await walk(`${SITE}/_api/v2.1/termStore/sets/${setId}/terms/${t.id}/children`, path);
      }
    };
    await walk(`${SITE}/_api/v2.1/termStore/sets/${setId}/children`, normalizeLabel(segName));
  }
  console.log(`resolved ${guidByPath.size} terms from the term store`);

  // 2. Existing rows, so a re-run updates instead of duplicating.
  const existing = new Map();
  const cur = await getJson(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items?$select=Id,TermGuid&$top=5000`,
  );
  (cur.value || []).forEach((r) => existing.set((r.TermGuid || "").toLowerCase(), r.Id));
  console.log(`list already holds ${existing.size} rows`);

  const ctx = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" },
  });
  const digest = (await ctx.json()).FormDigestValue;

  // 3. Write.
  let written = 0;
  let skippedBlank = 0;
  const missing = [];
  for (const r of parseCsv(CSV_TEXT)) {
    if ((r.Level || "").trim() === "Segment") continue; // named via StagingFolder

    // Path comes from the FILLED tier columns, not from the Level literal — see
    // the header. A tier column left blank simply ends the path.
    const parts = [r.Segment];
    for (const col of TIER_COLUMNS) {
      const v = (r[col] || "").trim();
      if (!v) break;
      parts.push(v);
    }
    const path = parts.map(normalizeLabel).join("|");
    const term = guidByPath.get(path);
    if (!term) {
      missing.push(path);
      continue;
    }

    // An empty Abbreviation is a deliberate state, not an error: reconciliation
    // SKIPS a term without one and reports it. Writing a blank row would look like
    // an abbreviation exists, so leave the row out entirely.
    if (!(r.Abbreviation || "").trim()) {
      skippedBlank++;
      continue;
    }

    const id = existing.get(term.id.toLowerCase());
    const base = `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items`;
    const res = await fetch(id ? `${base}(${id})` : base, {
      method: "POST",
      headers: Object.assign(
        {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest,
        },
        id ? { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" } : {},
      ),
      body: JSON.stringify({
        Title: term.label,
        TermGuid: term.id,
        Abbreviation: (r.Abbreviation || "").trim(),
        Level: (r.Level || "").trim(),
      }),
    });
    if (!res.ok) {
      console.error(`FAILED ${path}: HTTP ${res.status} ${await res.text()}`);
      continue;
    }
    written++;
  }

  console.log(`wrote ${written} rows`);
  if (skippedBlank) {
    console.log(`${skippedBlank} row(s) had no abbreviation and were left out deliberately`);
  }
  if (missing.length) {
    // A label path that does not resolve means the CSV text and the live term label
    // differ. Check the term store rather than editing the CSV to match. If a
    // missing row's Level says Region, suspect TIER_COLUMNS.
    console.warn(`could not resolve ${missing.length} label path(s):`, missing);
  }
})();

/**
 * Diagnoses why a CSV label path fails to match a live term. READ-ONLY.
 *
 * The seeder keys the CSV to the term store by label path, because the client
 * supplied names rather than GUIDs. Any character-level difference between the
 * two — a fullwidth ampersand, a curly apostrophe, a doubled space — silently
 * drops that row, and every row beneath it. The seeder reports only a count,
 * which is not enough to tell a fixable encoding difference from a real data
 * difference; this script separates the two.
 *
 * Reports four buckets:
 *   exact      — matched as-is
 *   normalized — matched only after normalising, so the seeder can be fixed
 *   unmatched  — a real difference in the data, needing a human decision
 *   orphanLive — live terms no CSV row claims (folders that would be skipped)
 *
 * Run: paste into DevTools on the site, with CSV_TEXT filled in as per
 * seed-term-abbreviations.js. Writes nothing — no POST, no digest.
 */
const SITE = "/sites/ClarenceDMSTesting";

const SETS = {
  "Group Head Office": "08dd94cb-f76c-431c-9b37-e9c98f739ffc",
  "Minamas Head Office": "9ad00b00-a43c-4a8b-a39a-d0efa89ba706",
  "NBPOL Head Office": "77c3993b-0c3c-4a18-89d9-d69209886322",
};

const CSV_TEXT = `PASTE THE CONTENTS OF docs/term-store-import/10-per-level-abbreviations.csv HERE`;

/**
 * Folds the differences that are presentation, not identity.
 *
 * Fullwidth ＆ is the known one: the term store requires it (see CLAUDE.md),
 * while the client's CSV uses plain &. The quote, dash and space rules are here
 * because Word and Excel rewrite those silently and the client's data came
 * through both.
 */
function normalizeLabel(s) {
  return (s || "")
    .replace(/＆/g, "&") // fullwidth ampersand
    .replace(/[‘’ʼ]/g, "'") // curly / modifier apostrophes
    .replace(/[“”]/g, '"')
    .replace(/[‐‑‒–—―]/g, "-") // en/em dashes and friends
    .replace(/ /g, " ") // non-breaking space
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

const getJson = async (url) => {
  const r = await fetch(url, {
    headers: { Accept: "application/json;odata=nometadata" },
  });
  if (!r.ok) throw new Error(`HTTP ${r.status} on ${url}: ${await r.text()}`);
  return r.json();
};

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

/** Segment rows are named via StagingFolder, not this list — see the seeder. */
function pathParts(r) {
  return r.Level === "Department"
    ? [r.BusinessSegment, r.Department]
    : [r.BusinessSegment, r.Department, r.Unit];
}

(async () => {
  const exactByPath = new Map();
  const normByPath = new Map();

  for (const segName of Object.keys(SETS)) {
    const setId = SETS[segName];
    const walk = async (url, prefix, normPrefix) => {
      const d = await getJson(url);
      for (const t of d.value || []) {
        const label = t.labels[0].name;
        const path = `${prefix}|${label}`;
        const nPath = `${normPrefix}|${normalizeLabel(label)}`;
        exactByPath.set(path, t.id);
        normByPath.set(nPath, { id: t.id, path: path });
        await walk(
          `${SITE}/_api/v2.1/termStore/sets/${setId}/terms/${t.id}/children`,
          path,
          nPath,
        );
      }
    };
    await walk(
      `${SITE}/_api/v2.1/termStore/sets/${setId}/children`,
      segName,
      normalizeLabel(segName),
    );
  }
  console.log(`live terms: ${exactByPath.size}`);

  const csvRows = parseCsv(CSV_TEXT).filter((r) => r.Level !== "Segment");
  const exact = [];
  const normalized = [];
  const unmatched = [];

  csvRows.forEach((r) => {
    const parts = pathParts(r);
    const path = parts.join("|");
    const nPath = parts.map(normalizeLabel).join("|");
    if (exactByPath.has(path)) exact.push(path);
    else if (normByPath.has(nPath))
      normalized.push({ csv: path, live: normByPath.get(nPath).path });
    else unmatched.push(path);
  });

  console.log(
    `csv rows: ${csvRows.length} | exact: ${exact.length} | fixed by normalising: ${normalized.length} | still unmatched: ${unmatched.length}`,
  );
  if (normalized.length) console.table(normalized.slice(0, 20));
  if (unmatched.length) {
    console.warn("UNMATCHED — these need a human decision:");
    unmatched.forEach((p) => console.warn("  " + p));
  }

  // Live terms no CSV row claims. Each would become a folder with no
  // abbreviation, which reconciliation skips and reports.
  const claimed = new Set();
  csvRows.forEach((r) => claimed.add(pathParts(r).map(normalizeLabel).join("|")));
  const orphanLive = [];
  normByPath.forEach((v, k) => {
    if (!claimed.has(k)) orphanLive.push(v.path);
  });
  console.log(`live terms with no CSV row: ${orphanLive.length}`);
  orphanLive.forEach((p) => console.log("  " + p));
})();

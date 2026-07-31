/**
 * Seeds the `DMS Term Abbreviation` list from
 * docs/term-store-import/10-per-level-abbreviations.csv.
 *
 * ── Why a console script rather than a web part ──────────────────────────────
 * Phase 1 of the abbreviation work needs the list populated once, with 175 rows.
 * The admin UI that maintains it afterwards is Phase 2. This unblocks testing
 * without waiting for that UI, the same way the term store and DMS Config have
 * been inspected throughout this project. JavaScript rather than PowerShell
 * because PS scripts do not run reliably in this environment.
 *
 * ── How to run ───────────────────────────────────────────────────────────────
 * 1. Create the list first: Custom List named `DMS Term Abbreviation` with Text
 *    columns TermGuid and Abbreviation, and a Choice column Level
 *    (Segment | Department | Unit). Title holds the full label.
 * 2. Open any page of the site in a browser, signed in.
 * 3. Paste the CSV contents into CSV_TEXT below.
 * 4. Paste this whole file into the DevTools console and press Enter.
 *
 * Re-runnable: an existing row for the same TermGuid is updated, not duplicated.
 *
 * Segments are skipped deliberately — a segment container has no term, so its
 * folder name comes from `StagingFolder` on the DMS Config mode row instead.
 * Set those by hand: mode_gho -> GHO, mode_minamas_ho -> MHO, mode_nbpol_ho -> NBPOLHO.
 */
const SITE = "/sites/ClarenceDMSTesting";
const LIST = "DMS Term Abbreviation";

// Per-site term set GUIDs. These are ClarenceDMSTesting's — a different site has
// its own, because the term store is site-collection local (see CLAUDE.md).
const SETS = {
  "Group Head Office": "08dd94cb-f76c-431c-9b37-e9c98f739ffc",
  "Minamas Head Office": "9ad00b00-a43c-4a8b-a39a-d0efa89ba706",
  "NBPOL Head Office": "77c3993b-0c3c-4a18-89d9-d69209886322",
};

const CSV_TEXT = `PASTE THE CONTENTS OF docs/term-store-import/10-per-level-abbreviations.csv HERE`;

const getJson = async (url) => {
  const r = await fetch(url, {
    headers: { Accept: "application/json;odata=nometadata" },
  });
  if (!r.ok) throw new Error(`HTTP ${r.status} on ${url}: ${await r.text()}`);
  return r.json();
};

/**
 * Minimal CSV reader. A naive split on "," is wrong here: the department
 * "Group Legal, Risk & Compliance" contains a comma and is therefore quoted.
 */
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
  // 1. Walk every term set, building "Segment|Department[|Unit]" -> term GUID.
  //    The CSV is keyed by label path because the client supplied names, not GUIDs.
  const guidByPath = new Map();
  for (const segName of Object.keys(SETS)) {
    const setId = SETS[segName];
    const walk = async (url, prefix) => {
      const d = await getJson(url);
      for (const t of d.value || []) {
        const label = t.labels[0].name;
        const path = `${prefix}|${label}`;
        guidByPath.set(path, t.id);
        await walk(
          `${SITE}/_api/v2.1/termStore/sets/${setId}/terms/${t.id}/children`,
          path,
        );
      }
    };
    await walk(`${SITE}/_api/v2.1/termStore/sets/${setId}/children`, segName);
  }
  console.log(`resolved ${guidByPath.size} terms from the term store`);

  // 2. Read existing rows so re-running updates rather than duplicates.
  const existing = new Map();
  const cur = await getJson(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(
      LIST,
    )}')/items?$select=Id,TermGuid&$top=5000`,
  );
  (cur.value || []).forEach((r) =>
    existing.set((r.TermGuid || "").toLowerCase(), r.Id),
  );
  console.log(`list already holds ${existing.size} rows`);

  // 3. Write.
  const ctx = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" },
  });
  const digest = (await ctx.json()).FormDigestValue;

  let written = 0;
  const missing = [];
  for (const r of parseCsv(CSV_TEXT)) {
    if (r.Level === "Segment") continue; // named via StagingFolder, not this list
    const path =
      r.Level === "Department"
        ? `${r.BusinessSegment}|${r.Department}`
        : `${r.BusinessSegment}|${r.Department}|${r.Unit}`;
    const guid = guidByPath.get(path);
    if (!guid) {
      missing.push(path);
      continue;
    }
    const id = existing.get(guid.toLowerCase());
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
        Title: r.FullName,
        TermGuid: guid,
        Abbreviation: r.Abbreviation,
        Level: r.Level,
      }),
    });
    if (!res.ok) {
      console.error(`FAILED ${path}: HTTP ${res.status} ${await res.text()}`);
      continue;
    }
    written++;
  }

  console.log(`wrote ${written} rows`);
  if (missing.length) {
    // A label path that does not resolve means the CSV text and the live term
    // label differ. Investigate the term store rather than editing the CSV to
    // match — the CSV was generated from the client's own data.
    console.warn(`could not resolve ${missing.length} label paths:`, missing);
  }
})();

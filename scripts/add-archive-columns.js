/*
 * Add the metadata columns the archive libraries are missing, so a MoveTo stops
 * silently dropping them.
 *
 * WHY THIS EXISTS
 * ---------------
 * SharePoint's MoveTo carries over ONLY the columns that already EXIST at the
 * destination. Everything else vanishes with no error and a green run. The two
 * archive libraries are missing several of `Documents`' columns - including
 * `LegallyPrivileged`, a LEGAL MARKER - so every archived document loses them
 * permanently. Same class of gap as the 2026-08-10 Remark/LegallyPrivileged
 * defect, at a new library.
 *
 * ⚠ RUN THIS BEFORE THE ARCHIVE MOVERS ARE EVER TURNED BACK ON. Nothing on the
 * test site is genuinely seven years old (the first file is eligible in 2033),
 * and the movers there were switched off after 2026-09-02 - so this is a
 * fix-before-enabling item, not a live data loss in progress. It does NOT
 * backfill: documents already archived carry nothing to restore.
 *
 * HOW TO RUN
 * ----------
 * Open the site in a browser, signed in as someone holding Manage Lists on
 * both archive libraries, press F12 -> Console, paste the whole file, Enter.
 * Read-modify only: it CREATES columns and never deletes or writes item data.
 *
 * ⚠ THE TECHNIQUE IS `src/shared/spColumns.ts`'s `ensureColumn`, DELIBERATELY -
 * do not "simplify" any of the four things it does:
 *   1. The existence check is `$filter`ed on InternalName, NEVER a `$top` page.
 *      A capped read is indistinguishable from the column being ABSENT, so it
 *      would try to CREATE a column that exists, fail on the duplicate internal
 *      name, and report a column it cannot create on a library that has it.
 *   2. `CreateFieldAsXml` is the ONLY form that sets the internal name
 *      (Name/StaticName) independently of the display name. Creating by display
 *      name would give `Vendor_x002f_CustomerName` a different internal name -
 *      and a matching display name over a DIFFERENT internal name fails exactly
 *      like an absent column, and looks right.
 *   3. `Options: 8` is AddToAllContentTypes, NOT 12. Twelve would add every
 *      column to every VIEW and reshape the views the client arranged.
 *   4. The Title MERGE afterwards is COSMETIC and must never fail the run - the
 *      column is already correct without it.
 *
 * ⚠ TAXONOMY COLUMNS ARE NOT CREATED HERE AND MUST NOT BE ADDED TO THIS LIST.
 * A taxonomy field needs its hidden note field and a term-set binding, and a
 * wrong binding yields a column that looks right and silently tags nothing -
 * which is why `ensureColumn` refuses to create one either. `Document Type`,
 * `Year` and `Confidentiality Level` were fixed BY HAND on 2026-09-02; if a
 * future site is missing them, do the same.
 */
(async () => {
  // ── Configure ────────────────────────────────────────────────────────────
  // Titles, not URL segments: `getbytitle` takes the title, and a rename never
  // touches the URL (gotcha #12). These are the LIVE titles as of 2026-09-02.
  // ⚠⚠ THESE TWO LISTS ARE PER-SITE AND MUST BE MEASURED, NEVER CARRIED OVER.
  // The values below are SDG's (`/sites/CRS`), read off its own `Restricted &
  // Confidential Document` library by `check-archive-column-gap.js` on 2026-09-10.
  // ClarenceDMSTesting needed FOUR columns and different library TITLES; SDG needs
  // TEN. **Run the gap check first on any site and paste its output here** - a
  // wrong TYPE yields a column that looks right and silently accepts nothing from
  // a MoveTo, which is the exact failure this script exists to prevent.
  const LIBRARIES = [
    "Archive Restricted & Confidential Document",
    "Archive Highly Confidential Document",
  ];

  // MEASURED, not guessed. Note `Remark` is single-line TEXT here, not Note, and
  // `Business_x0020_Segment`/`SubUnit` are PLAIN TEXT with a text `<Base>Tid` twin
  // rather than taxonomy - which is why they are script-safe at all.
  //
  // ⚠ SDG's archives ALREADY carry the three genuine taxonomy columns
  // (`Document_x0020_Type`, `Year`, `Confidentiality_x0020_Level`) with matching
  // types, so nothing has to be done by hand there. Do not assume that of a third
  // site: on ClarenceDMSTesting all three were wrong-type or absent and had to be
  // deleted and rebound to their term sets manually on 2026-09-02.
  const COLUMNS = [
    { name: "Business_x0020_Segment", title: "Business Segment", kind: "Text" },
    { name: "BusinessSegmentTid", title: "BusinessSegmentTid", kind: "Text" },
    { name: "SubUnit", title: "SubUnit", kind: "Text" },
    { name: "SubUnitTid", title: "SubUnitTid", kind: "Text" },
    { name: "DocumentDate", title: "Document Date", kind: "DateTime" },
    { name: "Vendor_x002f_CustomerName", title: "Vendor/CustomerName", kind: "Text" },
    { name: "ProjectName", title: "ProjectName", kind: "Text" },
    { name: "Remark", title: "Remark", kind: "Text" },
    { name: "LegallyPrivileged", title: "LegallyPrivileged", kind: "Boolean" },
    { name: "Full_x0020_Name", title: "Full Name", kind: "Text" },
  ];
  // ─────────────────────────────────────────────────────────────────────────

  // ⚠⚠ RESOLVED, NEVER PARSED OUT OF THE PAGE PATH — the first version parsed it and
  // broke the moment it was run from a library page on SDG
  // (`/sites/CRS/SiteAssets/Forms/AllItems.aspx`): it stripped only a few known
  // segments, kept the rest of the page path, and every REST call then fetched the
  // PAGE'S OWN HTML. It worked on ClarenceDMSTesting purely by which page it was run
  // from. `_spPageContextInfo` is tried first but threw a live `ReferenceError` in
  // `dump-site-inventory.js`, so it cannot be the only route.
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const resolveWeb = () => {
    if (SITE_OVERRIDE) return SITE_OVERRIDE.replace(/\/$/, "");
    try {
      const ctx = window._spPageContextInfo;
      if (ctx && ctx.webAbsoluteUrl) return String(ctx.webAbsoluteUrl).replace(/\/$/, "");
    } catch (e) {
      /* not defined on some modern pages — fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
    return location.origin + (m ? m[1] : "");
  };
  const web = resolveWeb();
  console.log("%cSite: " + web, "font-weight:bold");

  // ⚠ THE DIGEST FETCH IS WHERE A WRONG WEB URL FIRST SHOWS, AND IT MUST STOP THE RUN.
  // Unchecked, a page's HTML parses as no JSON, `FormDigestValue` is undefined, and
  // every create below then fails with a 403 — reading as a permissions problem on a
  // site the admin plainly administers.
  const ctxRes = await fetch(web + "/_api/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" },
  });
  const ctxType = ctxRes.headers.get("content-type") || "";
  if (!ctxRes.ok || !/json/i.test(ctxType)) {
    console.error(
      `Could not get a form digest from ${web} (HTTP ${ctxRes.status}, "${ctxType || "no content-type"}").\n` +
        `NOTHING was created. If that site URL is wrong, set SITE_OVERRIDE at the top and run again.`
    );
    return;
  }
  const digest = (await ctxRes.json()).FormDigestValue;
  if (!digest) {
    console.error(`No FormDigestValue came back from ${web}. NOTHING was created.`);
    return;
  }

  // The XML per type, verbatim from `ensureColumn`. DisplayName is set to the
  // INTERNAL name on purpose - SharePoint derives the internal name from
  // DisplayName when Name is absent, and setting all three identically means
  // the internal name cannot come out encoded. The Title is corrected after.
  const xmlFor = (n, kind) => {
    if (kind === "Boolean") {
      return `<Field Type="Boolean" DisplayName="${n}" Name="${n}" StaticName="${n}"><Default>0</Default></Field>`;
    }
    if (kind === "Note") {
      return `<Field Type="Note" DisplayName="${n}" Name="${n}" StaticName="${n}" NumLines="6" RichText="FALSE" />`;
    }
    if (kind === "DateTime") {
      return `<Field Type="DateTime" DisplayName="${n}" Name="${n}" StaticName="${n}" Format="DateTime" />`;
    }
    return `<Field Type="Text" DisplayName="${n}" Name="${n}" StaticName="${n}" MaxLength="255" />`;
  };

  const base = (lib) => `${web}/_api/web/lists/getbytitle('${encodeURIComponent(lib)}')`;

  let created = 0;
  let skipped = 0;
  let failed = 0;

  for (const lib of LIBRARIES) {
    console.log("%c\n" + lib, "font-weight:bold;font-size:13px");

    for (const col of COLUMNS) {
      // 1. Existence check - server-side filter, at most one row, cannot truncate.
      const probe = await fetch(
        `${base(lib)}/fields?$select=InternalName,Title,TypeAsString&$filter=InternalName eq '${col.name}'`,
        { headers: { Accept: "application/json;odata=nometadata" } }
      );
      if (!probe.ok) {
        console.warn(`  ✗ ${col.name}: could not read fields (HTTP ${probe.status}) - NOTHING was created for it`);
        failed++;
        continue;
      }
      const rows = (await probe.json()).value || [];
      if (rows.length > 0) {
        const r = rows[0];
        const warn = r.TypeAsString !== col.kind ? `  ⚠ type is ${r.TypeAsString}, expected ${col.kind}` : "";
        console.log(`  = ${col.name} already exists (${r.TypeAsString}, "${r.Title}")${warn}`);
        skipped++;
        continue;
      }

      // 2. Create. `Options: 8` = AddToAllContentTypes, joins no view.
      const mk = await fetch(`${base(lib)}/fields/createfieldasxml`, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest,
        },
        body: JSON.stringify({
          parameters: { SchemaXml: xmlFor(col.name, col.kind), Options: 8 },
        }),
      });
      if (!mk.ok) {
        console.warn(`  ✗ ${col.name}: create failed (HTTP ${mk.status}) ${await mk.text()}`);
        failed++;
        continue;
      }
      console.log(`  + ${col.name} created (${col.kind})`);
      created++;

      // 3. Rename to the display title. COSMETIC - a failure here leaves a
      //    correct column with an ugly label, so it must never fail the run.
      if (col.title !== col.name) {
        const ren = await fetch(
          `${base(lib)}/fields/getbyinternalnameortitle('${col.name}')`,
          {
            method: "POST",
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "X-RequestDigest": digest,
              "X-HTTP-Method": "MERGE",
              "IF-MATCH": "*",
            },
            body: JSON.stringify({ Title: col.title }),
          }
        );
        console.log(
          ren.ok
            ? `    renamed to "${col.title}"`
            : `    ⚠ could not rename to "${col.title}" (HTTP ${ren.status}) - the column is still correct`
        );
      }
    }
  }

  console.log(
    `%c\n${created} created · ${skipped} already there · ${failed} failed`,
    "font-weight:bold;font-size:13px"
  );
  if (failed > 0) {
    console.warn("Re-run it - creating a column that exists is skipped, so a second run only fills the gaps.");
  }
  console.log(
    "%cNothing is backfilled. Documents already in the archive carry no values for these columns.",
    "color:#a4262c"
  );
})();

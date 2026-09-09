/*
 * Diagnose the archive metadata gap on ANY site, before creating anything.
 *
 * Answers three questions in one run, in the order they have to be answered:
 *   1. What are this site's six libraries actually CALLED, and where do they sit?
 *   2. Which of `Documents`' columns are MISSING from each archive library?
 *   3. For each missing one, what TYPE is it on `Documents`?
 *
 * ⚠ RUN THIS BEFORE `add-archive-columns.js` ON A NEW SITE. That script carries a
 * hardcoded library list and column list measured on ClarenceDMSTesting, and both
 * are wrong on SDG: SDG has never been renamed, so its archive pair is probably
 * titled `Archive` / `HC Archive`, and its column types must be read rather than
 * assumed. A wrong TYPE gives you a column that looks right and silently accepts
 * nothing from a MoveTo — which is the exact failure this whole exercise is about.
 *
 * HOW TO RUN
 * ----------
 * Open the site, F12 -> Console, paste, Enter. **READ ONLY** - every request is a
 * GET and nothing is created, renamed or deleted.
 *
 * ⚠ THE LIBRARIES ARE FOUND BY URL SEGMENT, NEVER BY TITLE. A rename never touches
 * the URL (gotcha #12), and this client renames libraries as a matter of course -
 * three times in two days at one point. So `Archive` and `HCArchive` are the stable
 * handles, and the TITLE is only reported so the creation script can be pointed at
 * whatever they are called today.
 *
 * ⚠ TAXONOMY COLUMNS ARE REPORTED AND MUST BE CREATED BY HAND. A taxonomy field
 * needs its hidden note field and a term-set binding, and a wrong binding yields a
 * column that looks right and tags nothing - so `ensureColumn` refuses to create
 * one and so does the creation script. On ClarenceDMSTesting the three
 * (`Document Type`, `Year`, `Confidentiality Level`) were fixed by hand on
 * 2026-09-02; expect to do the same here.
 */
(async () => {
  // ── Configure ────────────────────────────────────────────────────────────
  // URL segments, not titles. `Shared Documents` is the normal Documents library
  // (gotcha #12) - it has never been at /Documents on any site.
  const SOURCE_SEGMENT = "Shared Documents";
  const ARCHIVE_SEGMENTS = ["Archive", "HCArchive"];

  // ⚠⚠ AN EXACT-MATCH LIST, NEVER A PREFIX PATTERN — AND THE FIRST VERSION OF THIS
  // SCRIPT PROVED WHY. It began as a regex with a `Doc` alternative, which swallowed
  // `DocumentDate` AND `Document_x0020_Type` — the two columns this whole exercise is
  // about — so the report silently omitted the very gaps it exists to find. Found by
  // running it against a synthetic site, not by reading it.
  //
  // Same lesson as the `Forms` folder filter (2026-09-07): **a name-based filter
  // cannot tell a system object from user data**, and a client column may legitimately
  // begin with any of these words. Only the two genuinely reserved PREFIXES stay
  // patterned; everything else is named in full, so adding a name is deliberate.
  const SKIP_PREFIX = /^(_|ows|xd_|Tax(CatchAll|Keyword)|AppAuthor|AppEditor|ComplianceAsset)/;
  const SKIP_EXACT = new Set([
    "ID", "GUID", "Title", "Created", "Modified", "Author", "Editor", "ContentType",
    "ContentTypeId", "Attachments", "Order", "MetaInfo", "InstanceID", "owshiddenversion",
    "FileRef", "FileDirRef", "FileLeafRef", "FSObjType", "FileSizeDisplay", "File_x0020_Type",
    "File_x0020_Size", "DocIcon", "LinkFilename", "LinkFilenameNoMenu", "LinkFilename2",
    "LinkTitle", "LinkTitleNoMenu", "LinkTitle2", "SelectTitle", "SelectFilename",
    "EditHTML", "Edit", "ServerUrl", "EncodedAbsUrl", "BaseName", "ServerRedirectedEmbedUrl",
    "ServerRedirectedEmbedUri", "Combine", "RepairDocument", "TemplateUrl", "ProgId",
    "ScopeId", "UniqueId", "SyncClientId", "CheckoutUser", "CheckedOutTitle",
    "CheckedOutUserId", "IsCheckedoutToLocal", "VirusStatus", "PermMask", "SortBehavior",
    "ItemChildCount", "FolderChildCount", "Restricted", "OriginatorId", "NoExecute",
    "ContentVersion", "ParentVersionString", "ParentLeafName", "Last_x0020_Modified",
    "Created_x0020_Date", "Modified_x0020_By", "Created_x0020_By", "HTML_x0020_File_x0020_Type",
    "Predecessors", "WorkflowVersion", "WorkflowInstanceID", "UIVersion", "UIVersionString",
    "Level", "Modified_By", "Created_By", "ThumbnailOnForm", "AccessPolicy", "StreamHash",
    "MediaServiceAutoTags", "MediaServiceOCR", "MediaServiceLocation", "MediaServiceDateTaken",
  ]);
  const skip = (n) => SKIP_PREFIX.test(n) || SKIP_EXACT.has(n);
  // ─────────────────────────────────────────────────────────────────────────

  // ⚠⚠ THE WEB URL IS RESOLVED, NEVER PARSED OUT OF THE PAGE PATH — AND THE FIRST
  // VERSION PARSED IT AND BROKE ON SDG. It stripped only `/SitePages` and `/Lists`,
  // so run from `/sites/CRS/SiteAssets/Forms/AllItems.aspx` it kept the whole page
  // path, and every REST call fetched the PAGE'S OWN HTML — surfacing as
  // `Unexpected token '<', "<!DOCTYPE "... is not valid JSON`, which names neither
  // the cause nor the URL. **A library page is exactly where somebody runs this**,
  // so the working directory can be any of a dozen shapes.
  //
  // ⚠ `_spPageContextInfo` IS TRIED FIRST BUT CANNOT BE RELIED ON — it threw a live
  // `ReferenceError` in `dump-site-inventory.js` on a modern page, which is why that
  // script grew the same fallback. So: the context object, then the `/sites/<name>`
  // or `/teams/<name>` prefix, then the origin. Set SITE_OVERRIDE to settle it by
  // hand if a site is ever shaped differently.
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
  const get = async (u) => {
    const r = await fetch(u, { headers: { Accept: "application/json;odata=nometadata" } });
    if (!r.ok) throw new Error("HTTP " + r.status + " on " + u);
    // ⚠ A WRONG WEB URL ANSWERS 200 WITH A PAGE, NOT AN ERROR — so `r.ok` proves
    // nothing here. `r.json()` then throws `Unexpected token '<'`, which names
    // neither the cause nor the URL and reads as a code fault. Checked explicitly
    // so the message says which address was asked and what came back.
    const ct = r.headers.get("content-type") || "";
    if (!/json/i.test(ct)) {
      throw new Error(
        `expected JSON and got "${ct || "no content-type"}" from ${u}\n` +
          `    The site URL resolved to ${web} — if that is not this site's web, ` +
          `set SITE_OVERRIDE at the top of this script and run it again.`
      );
    }
    return r.json();
  };

  console.log("%cSite: " + web, "font-weight:bold;font-size:13px");

  // 1. Every document library, with its title AND its URL segment. Both are
  //    needed: the title is what `getbytitle` wants, the segment is what
  //    identifies the library through a rename.
  const lists = (await get(
    `${web}/_api/web/lists?$select=Title,Id&$expand=RootFolder&$filter=BaseTemplate eq 101`
  )).value;
  const bySegment = {};
  const table = lists.map((l) => {
    const seg = decodeURIComponent((l.RootFolder?.ServerRelativeUrl || "").split("/").pop() || "");
    bySegment[seg] = l;
    return { Title: l.Title, "URL segment": seg };
  });
  console.log("%c\nLIBRARIES ON THIS SITE", "font-weight:bold");
  console.table(table);

  const source = bySegment[SOURCE_SEGMENT];
  if (!source) {
    console.error(
      `Could not find the source library at /${SOURCE_SEGMENT}. Check the table above and set SOURCE_SEGMENT.`
    );
    return;
  }

  // 2. The source library's real columns, with their types.
  const fieldsOf = async (title) =>
    (await get(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')` +
        `/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField&$top=5000`
    )).value;

  const srcFields = (await fieldsOf(source.Title)).filter(
    (f) => !f.Hidden && !f.ReadOnlyField && !skip(f.InternalName)
  );
  console.log(
    `%c\nSOURCE: "${source.Title}" (/${SOURCE_SEGMENT}) - ${srcFields.length} candidate column(s)`,
    "font-weight:bold"
  );

  // 3. Diff each archive library against it.
  for (const seg of ARCHIVE_SEGMENTS) {
    const lib = bySegment[seg];
    console.log("%c\n" + "=".repeat(60), "color:#888");
    if (!lib) {
      console.warn(`No library at /${seg} on this site - nothing to check.`);
      continue;
    }
    console.log(`%c"${lib.Title}"  (/${seg})`, "font-weight:bold;font-size:13px");

    let have;
    try {
      have = await fieldsOf(lib.Title);
    } catch (e) {
      // ⚠ An unreadable field list is NOT "no columns". Reporting every column as
      // missing here would send somebody to create columns that already exist,
      // which SharePoint refuses on the duplicate internal name.
      console.warn(`  ✗ could not read its fields (${e.message}) - NO conclusion drawn for this library`);
      continue;
    }
    const haveNames = {};
    for (const f of have) haveNames[f.InternalName.toLowerCase()] = f;

    const missing = [];
    const mismatched = [];
    for (const f of srcFields) {
      const there = haveNames[f.InternalName.toLowerCase()];
      if (!there) missing.push(f);
      else if (there.TypeAsString !== f.TypeAsString) mismatched.push({ f, there });
    }

    if (missing.length === 0 && mismatched.length === 0) {
      console.log("  ✓ every source column is present with a matching type");
      continue;
    }

    if (missing.length > 0) {
      // Split by whether the creation script can make it. The taxonomy ones are
      // the reason this is a report rather than a fix.
      const taxonomy = missing.filter((f) => /^Taxonomy/.test(f.TypeAsString));
      const plain = missing.filter((f) => !/^Taxonomy/.test(f.TypeAsString));

      if (plain.length > 0) {
        console.log(`%c  ${plain.length} MISSING, safe to create by script:`, "font-weight:bold");
        console.table(
          plain.map((f) => ({
            "internal name": f.InternalName,
            "display title": f.Title,
            type: f.TypeAsString,
          }))
        );
        // Paste-ready, so nobody has to retype a `_x002f_` by hand.
        console.log(
          "%c  COLUMNS list for add-archive-columns.js:",
          "font-weight:bold"
        );
        console.log(
          plain
            .map(
              (f) =>
                `    { name: "${f.InternalName}", title: "${f.Title}", kind: "${
                  f.TypeAsString === "Note" ? "Note" : f.TypeAsString === "DateTime" ? "DateTime" : f.TypeAsString === "Boolean" ? "Boolean" : "Text"
                }" },`
            )
            .join("\n")
        );
      }

      if (taxonomy.length > 0) {
        console.log(
          `%c  ${taxonomy.length} MISSING and MUST be added BY HAND (managed metadata):`,
          "font-weight:bold;color:#a4262c"
        );
        console.table(
          taxonomy.map((f) => ({
            "internal name": f.InternalName,
            "display title": f.Title,
            type: f.TypeAsString,
          }))
        );
        console.log(
          "%c  A taxonomy column needs its term-set binding. A wrong binding looks correct and tags nothing.",
          "color:#a4262c"
        );
      }
    }

    if (mismatched.length > 0) {
      console.log(
        `%c  ⚠ ${mismatched.length} present with the WRONG TYPE - a MoveTo will not carry these:`,
        "font-weight:bold;color:#a4262c"
      );
      console.table(
        mismatched.map(({ f, there }) => ({
          "internal name": f.InternalName,
          "source type": f.TypeAsString,
          "archive type": there.TypeAsString,
        }))
      );
      console.log(
        "%c  Fixing one means DELETING and recreating it - and a deleted column takes its data with it and does NOT reach the recycle bin.",
        "color:#a4262c"
      );
    }
  }

  console.log(
    "%c\nThe LIBRARY TITLES above are what add-archive-columns.js needs in its LIBRARIES list.",
    "font-weight:bold"
  );
})();

/**
 * Creates the metadata columns on the Highly Confidential library pair.
 *
 * Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md
 *
 * ── Why a console script ─────────────────────────────────────────────────────
 * 18 columns x 2 libraries = 36 operations, three of which carry a PERMANENT trap:
 * a column's internal name is frozen at creation and never changes on rename, so
 * creating "Document Date" yields `Document_x0020_Date` and every upload into that
 * library then fails its WHOLE metadata write — not just the date, because one
 * unknown field name fails the entire validateUpdateListItem call, and that call
 * returns HTTP 200 either way.
 *
 * JavaScript rather than PowerShell because PS does not run reliably in this
 * environment, and because this has to work on SDG's tenant too.
 *
 * ── How to run ───────────────────────────────────────────────────────────────
 * 1. Create the two libraries FIRST, without spaces, then retitle them:
 *      HCApprovalDocument -> "HC Approval Document"
 *      HCDocuments        -> "HC Documents"
 *    (A list's URL is fixed at creation and never moves on rename — gotcha #12.)
 * 2. Open any page of the site in a browser, signed in as someone with Manage Lists.
 * 3. Paste this whole file into the DevTools console and press Enter.
 *
 * Re-runnable: a column already present under the same INTERNAL name is skipped and
 * never converted — a type change on a populated column loses its data.
 *
 * ── What it does NOT do ──────────────────────────────────────────────────────
 * The three MANAGED METADATA columns are deliberately left out:
 *      Document Type, Year, Confidentiality Level
 * Creating a taxonomy field over REST means hand-building its hidden note field and
 * then stitching SspId / TermSetId / TextField together. Get one wrong and the result
 * is a column that looks right, saves nothing, and has to be deleted and rebuilt. The
 * list settings UI does it correctly in four clicks; this script prints the term-set
 * GUIDs at the end. Six manual columns instead of thirty-six.
 */

(async () => {
  "use strict";

  const SITE = _spPageContextInfo.webAbsoluteUrl;

  /**
   * The two HC libraries, as CANDIDATE titles in probe order.
   *
   * The retitled name first, then the name the library is created under. Same order and
   * same reason as `HC_APPROVAL_CANDIDATES` in shared/naming.ts: a list's URL is fixed at
   * creation and never moves on rename, so the two halves diverge on purpose — and until
   * someone performs the retitle, the library answers only to the space-free name.
   *
   * An earlier version of this script demanded the retitled name and answered 404 for a
   * library sitting right there, which reads as "you created it wrong" rather than "you
   * have not renamed it yet".
   */
  const LIBRARIES = [
    ["HC Approval Document", "HCApprovalDocument"],
    ["HC Documents", "HCDocuments"],
  ];

  /**
   * Every non-taxonomy column, as (internalName, displayName, type).
   *
   * `internal` is what the code writes to and is frozen forever; `display` is cosmetic
   * and applied by a rename afterwards. Where the two differ, the difference IS the
   * point — see the file header.
   *
   * Taken from a live read of `Approval Document` on 2026-08-16, with the test debris
   * deliberately excluded: Testing, Testing11, CreditCard, Stage, Estate, Company and
   * Function are leftovers from Structure Manager experiments and are referenced by no
   * live `Levels` chain. Add them here only if a DMS Config mode row proves otherwise.
   */
  const COLUMNS = [
    // ── Tier columns (label + Tid pairs) ──
    ["Business_x0020_Segment", "Business Segment", "Text"],
    ["BusinessSegmentTid", "Business Segment Tid", "Text"],
    ["Department", "Department", "Text"],
    ["DepartmentTid", "DepartmentTid", "Text"],
    ["Unit", "Unit", "Text"],
    ["UnitTid", "UnitTid", "Text"],
    ["SubUnit", "SubUnit", "Text"],
    ["SubUnitTid", "SubUnitTid", "Text"],
    ["Region", "Region", "Text"],
    ["RegionTid", "Region ID", "Text"],
    // TRAP: created as "Estate/Mill" this becomes Estate_x002f_Mill, which is not what
    // Upstream Operations Malaysia writes to.
    ["EstateMill", "Estate/Mill", "Text"],
    ["EstateMillTid", "Estate/Mill ID", "Text"],
    // ── Document metadata ──
    // TRAP: created as "Document Date" this becomes Document_x0020_Date.
    ["DocumentDate", "Document Date", "DateTime"],
    ["Vendor_x002f_CustomerName", "Vendor/CustomerName", "Text"],
    ["ProjectName", "ProjectName", "Text"],
    ["Remark", "Remark", "Text"],
    ["LegallyPrivileged", "LegallyPrivileged", "Boolean"],
    // On FOLDERS, not files: the term's real label behind the abbreviated folder name.
    ["Full_x0020_Name", "Full Name", "Text"],
  ];

  /** Managed metadata — created by hand, listed here so the reminder is exact. */
  const TAXONOMY = [
    ["Document Type", "866c5754-258e-401f-8685-03d20ae59b1d"],
    ["Year", "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf"],
    ["Confidentiality Level", "0d6d1da8-27e5-477f-8684-e8cf169f8fb9"],
  ];

  /* ── Plumbing ─────────────────────────────────────────────────────────────── */

  const digest = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" },
  })
    .then((r) => r.json())
    .then((d) => d.FormDigestValue);

  const listBase = (title) =>
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')`;

  const post = (url, body, extra = {}) =>
    fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        ...extra,
      },
      body: JSON.stringify(body),
    });

  /** The field XML. Only these three types; anything else must be created by hand. */
  const schemaFor = (internal, type) => {
    const common = `DisplayName="${internal}" Name="${internal}" StaticName="${internal}"`;
    if (type === "DateTime") {
      // Format="DateTime" keeps the TIME component. The default is date-only, which
      // silently collapses every document filed on one day to the same instant.
      return `<Field Type="DateTime" ${common} Format="DateTime" />`;
    }
    if (type === "Boolean") return `<Field Type="Boolean" ${common} />`;
    return `<Field Type="Text" ${common} MaxLength="255" />`;
  };

  /* ── Run ──────────────────────────────────────────────────────────────────── */

  for (const candidates of LIBRARIES) {
    // Existence is matched on INTERNAL name, never the display title: a matching title
    // over a different internal name fails exactly like an absent column, and looks
    // completely correct in the UI.
    let library;
    let res;
    for (const candidate of candidates) {
      const attempt = await fetch(
        `${listBase(candidate)}/fields?$select=InternalName&$top=500`,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (attempt.ok) {
        library = candidate;
        res = attempt;
        break;
      }
    }
    if (!library) {
      console.error(
        `%c\nx no library found under ${candidates.map((c) => `"${c}"`).join(" or ")}`,
        "font-weight:bold;color:#a4262c",
      );
      console.error(
        "  Create it, or run the snippet below to see what the document libraries are called:\n" +
          "  fetch(`${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists?$select=Title&$filter=BaseTemplate eq 101`,\n" +
          "    {headers:{Accept:'application/json;odata=nometadata'}}).then(r=>r.json()).then(d=>console.table(d.value));",
      );
      continue;
    }

    console.log(`%c\n${library}`, "font-weight:bold;font-size:14px");
    if (library !== candidates[0]) {
      // Not an error — the library exists and the columns will be created correctly. But the
      // web parts resolve BOTH candidates too, so this only matters for how it reads on screen.
      console.log(
        `  (still titled "${library}" — retitle it to "${candidates[0]}" when convenient; ` +
          "the URL stays as it is, which is the point)",
      );
    }
    const present = new Set(
      ((await res.json()).value || []).map((f) => (f.InternalName || "").toLowerCase()),
    );

    let made = 0;
    let skipped = 0;
    for (const [internal, display, type] of COLUMNS) {
      if (present.has(internal.toLowerCase())) {
        skipped++;
        console.log(`  = ${internal} already there`);
        continue;
      }

      // Options: 8 = AddToAllContentTypes. NOT 12, which also adds every column to the
      // default view and would reshape it eighteen times over.
      const create = await post(`${listBase(library)}/fields/CreateFieldAsXml`, {
        parameters: { SchemaXml: schemaFor(internal, type), Options: 8 },
      });
      if (!create.ok) {
        const body = await create.text().catch(() => "");
        console.error(`  x ${internal} — HTTP ${create.status} ${body.slice(0, 200)}`);
        continue;
      }

      // The internal name is frozen by the create above, so this only changes what is
      // displayed. A failure here is cosmetic and must not fail the column.
      if (display !== internal) {
        const rename = await post(
          `${listBase(library)}/fields/getbyinternalnameortitle('${encodeURIComponent(internal)}')`,
          { Title: display },
          { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
        );
        if (!rename.ok) {
          console.warn(
            `  ! ${internal} created, but could not be renamed to "${display}" ` +
              `(HTTP ${rename.status}). The internal name is correct — rename it in list settings.`,
          );
        }
      }
      made++;
      console.log(`  + ${internal}${display !== internal ? ` -> "${display}"` : ""}`);
    }
    console.log(`  ${made} created, ${skipped} already present`);
  }

  console.log(
    "%c\nSTILL TO DO BY HAND — the three managed metadata columns, in BOTH libraries:",
    "font-weight:bold;color:#a4262c",
  );
  for (const [name, termSet] of TAXONOMY) {
    console.log(`  "${name}"  ->  term set ${termSet}`);
  }
  console.log(
    "Library settings -> Create column -> Managed Metadata -> pick the term set.\n" +
      "Create them with the name EXACTLY as printed: the space is what produces the\n" +
      "_x0020_ in the internal name, and here that is the correct outcome.",
  );
  console.log(
    "%cThen verify — diff InternalName, never Title:",
    "font-weight:bold",
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIBRARIES[0][0])}')` +
      "/fields?$select=Title,InternalName&$filter=Hidden eq false",
  );
})();

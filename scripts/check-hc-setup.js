/**
 * Checks whether the Highly Confidential library pair is ready to use.
 *
 * Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md
 *
 * READ ONLY. It writes nothing and changes nothing, so it is safe to run at any point,
 * and safe to run repeatedly.
 *
 * ── Why this exists ──────────────────────────────────────────────────────────
 * Every failure it looks for is one that LOOKS FINE in the SharePoint UI:
 *
 *   - a column whose display name is right and whose internal name is wrong fails
 *     exactly like a missing column, and takes the whole metadata write with it,
 *     because ONE unknown field name fails the entire validateUpdateListItem call —
 *     and that call answers HTTP 200 either way;
 *   - a "Confidentiality Level" created as plain Text instead of Managed Metadata
 *     accepts a value and stores the wrong thing;
 *   - a taxonomy column bound to the wrong term set offers the wrong dropdown;
 *   - content approval left ON in HC Documents makes every routed file invisible to
 *     the whole unit, which presents as a permissions bug and is not one;
 *   - Draft Item Security left at the default on the approval side lets peers read
 *     each other's pending Highly Confidential documents.
 *
 * ── How to run ───────────────────────────────────────────────────────────────
 * Open any page of the site, signed in, and paste this whole file into the DevTools
 * console.
 */

(async () => {
  "use strict";

  /**
   * The site URL, WITHOUT depending on `_spPageContextInfo`.
   *
   * That global exists only on a rendered SharePoint page. Paste this into the console while looking
   * at an `_api` response — which is exactly what someone checking columns has open — and it throws
   * `_spPageContextInfo is not defined` before doing anything, which reads as a broken script rather
   * than as the wrong tab (hit 2026-08-19).
   *
   * The fallback derives the site from the path, so it works on any page of the site collection,
   * including an `_api` URL. Cookies travel with same-origin fetch either way.
   */
  const SITE =
    (typeof _spPageContextInfo !== "undefined" && _spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) ||
    location.origin + ((location.pathname.match(/^\/sites\/[^/]+/) || [""])[0]);

  /** Candidate titles, retitled name first — a library answers to whichever it has. */
  const LIBRARIES = [
    { role: "approval", candidates: ["HC Approval Document", "HCApprovalDocument"] },
    { role: "documents", candidates: ["HC Documents", "HCDocuments"] },
  ];

  /** internalName -> expected TypeAsString. Must match the normal pair exactly. */
  const EXPECTED = {
    Business_x0020_Segment: "Text",
    BusinessSegmentTid: "Text",
    Department: "Text",
    DepartmentTid: "Text",
    Unit: "Text",
    UnitTid: "Text",
    SubUnit: "Text",
    SubUnitTid: "Text",
    Region: "Text",
    RegionTid: "Text",
    EstateMill: "Text",
    EstateMillTid: "Text",
    DocumentDate: "DateTime",
    Vendor_x002f_CustomerName: "Text",
    ProjectName: "Text",
    Remark: "Text",
    LegallyPrivileged: "Boolean",
    Full_x0020_Name: "Text",
    Document_x0020_Type: "TaxonomyFieldType",
    Year: "TaxonomyFieldType",
    Confidentiality_x0020_Level: "TaxonomyFieldType",
  };

  /** The term set each taxonomy column must be bound to, on THIS site. */
  const TERM_SETS = {
    Document_x0020_Type: "866c5754-258e-401f-8685-03d20ae59b1d",
    Year: "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf",
    Confidentiality_x0020_Level: "0d6d1da8-27e5-477f-8684-e8cf169f8fb9",
  };

  const get = (url) =>
    fetch(url, { headers: { Accept: "application/json;odata=nometadata" } });

  const listBase = (title) =>
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')`;

  const ok = (m) => console.log(`%c  PASS  %c${m}`, "color:#0f6c3f;font-weight:bold", "");
  const warn = (m) => console.log(`%c  ??    %c${m}`, "color:#8a4b00;font-weight:bold", "");

  let failures = 0;
  const fail = (m) => {
    failures++;
    console.log(`%c  FAIL  %c${m}`, "color:#a4262c;font-weight:bold", "");
  };

  /* Read the document libraries once and match in code. Probing titles one at a time
     would log a 404 per miss whether or not the code handles it — which made a
     completely successful run of the sibling script look broken. */
  const lists = await get(
    `${SITE}/_api/web/lists?$select=Title,EnableModeration,DraftVersionVisibility,EnableVersioning&$filter=BaseTemplate eq 101`,
  )
    .then((r) => (r.ok ? r.json() : { value: [] }))
    .then((d) => d.value || []);

  for (const { role, candidates } of LIBRARIES) {
    const list = lists.find((l) =>
      candidates.some((c) => c.toLowerCase() === (l.Title || "").toLowerCase()),
    );

    console.log(`%c\n${candidates[0]}`, "font-weight:bold;font-size:14px");
    if (!list) {
      fail(`no library titled ${candidates.map((c) => `"${c}"`).join(" or ")}`);
      console.log(`        libraries here: ${lists.map((l) => l.Title).join(", ") || "(none)"}`);
      continue;
    }
    const title = list.Title;
    if (title !== candidates[0]) {
      warn(`still titled "${title}" — retitle to "${candidates[0]}" when convenient (the URL stays put)`);
    }

    /* ── Columns ─────────────────────────────────────────────────────────────
       Matched on INTERNAL name, never the display title. A right-looking title over a
       wrong internal name is the failure this whole check exists for. */
    const fieldsRes = await get(
      `${listBase(title)}/fields?$select=InternalName,Title,TypeAsString&$top=500`,
    );
    if (!fieldsRes.ok) {
      fail(`could not read its columns (HTTP ${fieldsRes.status})`);
      continue;
    }
    const fields = ((await fieldsRes.json()).value || []).reduce((m, f) => {
      m[(f.InternalName || "").toLowerCase()] = f;
      return m;
    }, {});

    const missing = [];
    const wrongType = [];
    for (const [internal, type] of Object.entries(EXPECTED)) {
      const f = fields[internal.toLowerCase()];
      if (!f) {
        missing.push(internal);
        continue;
      }
      // TaxonomyFieldTypeMulti is wrong here too: these are single-value fields, and the
      // multi variant writes a different value shape.
      if ((f.TypeAsString || "") !== type) {
        wrongType.push(`${internal} is ${f.TypeAsString}, expected ${type}`);
      }
    }

    if (missing.length === 0) ok(`all ${Object.keys(EXPECTED).length} columns present`);
    else fail(`${missing.length} column(s) MISSING: ${missing.join(", ")}`);

    if (wrongType.length > 0) for (const w of wrongType) fail(w);
    else if (missing.length === 0) ok("every column is the right type");

    /* ── Taxonomy bindings ───────────────────────────────────────────────────
       A column bound to the wrong term set offers the wrong dropdown and stores terms
       nothing else on the site recognises. Read per field WITHOUT a $select: TermSetId
       exists only on TaxonomyField, and naming it against the base type fails the whole
       request rather than omitting the key. */
    for (const [internal, wanted] of Object.entries(TERM_SETS)) {
      if (!fields[internal.toLowerCase()]) continue; // already reported as missing
      const res = await get(
        `${listBase(title)}/fields/getbyinternalnameortitle('${encodeURIComponent(internal)}')`,
      );
      if (!res.ok) {
        warn(`${internal}: could not read its binding (HTTP ${res.status})`);
        continue;
      }
      const got = ((await res.json()).TermSetId || "").toLowerCase();
      if (!got) warn(`${internal}: no term set recorded — is it really Managed Metadata?`);
      else if (got !== wanted.toLowerCase()) fail(`${internal} is bound to ${got}, expected ${wanted}`);
      else ok(`${internal} bound to the right term set`);
    }

    /* ── Moderation ──────────────────────────────────────────────────────────
       Opposite settings on the two halves, and BOTH failures stay invisible until
       somebody cannot find a document. */
    if (role === "approval") {
      if (list.EnableModeration) ok("content approval is ON");
      else fail("content approval is OFF — it must be ON, or nothing is ever approved");
      // 2 = "Only users who can approve items (and the author)". Anything else lets a PIC
      // read a peer's pending Highly Confidential drafts.
      if (list.DraftVersionVisibility === 2) ok("Draft Item Security = approver and author");
      else {
        fail(
          `Draft Item Security is ${list.DraftVersionVisibility}, expected 2 ` +
            "(Only users who can approve items, and the author)",
        );
      }
    } else if (!list.EnableModeration) {
      ok("content approval is OFF");
    } else {
      fail(
        "content approval is ON — every routed file will arrive Pending and be invisible to " +
          "the whole unit. Versioning settings -> Require content approval -> No.",
      );
    }

    /* ── Folder content type ─────────────────────────────────────────────────
       Without it Full Name never reaches the details pane. Soft: folders are still
       created and permissioned correctly, so this is a warning, not a failure. */
    const ctRes = await get(`${listBase(title)}/ContentTypes?$select=Name`);
    if (ctRes.ok) {
      const names = ((await ctRes.json()).value || []).map((c) => (c.Name || "").toLowerCase());
      if (names.indexOf("crs folder") > -1 || names.indexOf("dms folder") > -1) {
        ok("a folder content type is attached");
      } else {
        warn('no "CRS Folder" content type — folders work, but Full Name will not show in the details pane');
      }
    }
  }

  console.log(
    failures === 0
      ? "%c\nReady. Next: deploy 1.0.118.0, create the HC groups, then re-run reconciliation."
      : `%c\n${failures} thing(s) to fix above before this pair is usable.`,
    `font-weight:bold;font-size:13px;color:${failures === 0 ? "#0f6c3f" : "#a4262c"}`,
  );
  console.log(
    "Not checked here, because none of it is readable from a library: the two Power Automate " +
      "flows, the HC groups and their Group Map rows, and whether reconciliation has run.",
  );
})();

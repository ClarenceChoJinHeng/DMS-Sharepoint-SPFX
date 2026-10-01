/**
 * Checks whether the CRS document libraries are ready to use.
 *
 * Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md
 * Referenced by: docs/2026-08-16-full-system-test-plan.md
 *
 * READ ONLY. It writes nothing and changes nothing, so it is safe to run at any point,
 * and safe to run repeatedly.
 *
 * ── Scope widened 2026-08-28, and the reason matters more than the widening ───
 * This checked the HC PAIR against a HARDCODED manifest, and both halves of that were
 * pinned to ClarenceDMSTesting:
 *
 *   - TERM_SETS held that site's three GUIDs. SDG's Document Type is
 *     fd973a45-3c34-483a-bfae-2a5b9dc5a556, so every binding read FAIL there.
 *   - EXPECTED listed Region/RegionTid/EstateMill/EstateMillTid — Upstream Operations
 *     tier columns, which exist on no site that has not onboarded that segment.
 *
 * Seven spurious failures on SDG, on a script whose entire value is being trusted. A
 * false FAIL is worse than no check: it teaches the reader to skim past the line that
 * matters on the day it is real.
 *
 * So nothing is hardcoded now. `Approval Document` is the REFERENCE — its own columns
 * are the expected set, and its taxonomy bindings are the expected bindings. Both are
 * read at runtime, so the check follows the site instead of describing one site.
 *
 * ⚠ THE REFERENCE LIBRARY IS THEREFORE NOT ITSELF VERIFIED. Nothing here can tell you
 *   `Approval Document` is right; it can only tell you the other five agree with it.
 *   That is the correct trade — a column set is only meaningful relative to the library
 *   documents are copied INTO — but it is a real limit, so it is stated in the output.
 *
 * ── Why this exists ──────────────────────────────────────────────────────────
 * Every failure it looks for is one that LOOKS FINE in the SharePoint UI:
 *
 *   - a column whose display name is right and whose internal name is wrong fails
 *     exactly like a missing column, and takes the whole metadata write with it,
 *     because ONE unknown field name fails the entire validateUpdateListItem call —
 *     and that call answers HTTP 200 either way;
 *   - a column absent at the destination is silently dropped by Auto-route's copy, on a
 *     green run, so an approved document arrives with no Remark and no LegallyPrivileged;
 *   - a "Confidentiality Level" created as plain Text instead of Managed Metadata
 *     accepts a value and stores the wrong thing;
 *   - a taxonomy column bound to the wrong term set offers the wrong dropdown;
 *   - content approval left ON in HC Documents makes every routed file invisible to
 *     the whole unit, which presents as a permissions bug and is not one;
 *   - Draft Item Security left at the default on the approval side lets peers read
 *     each other's pending Highly Confidential documents;
 *   - versioning OFF on the approved side means Auto-route's Replace DESTROYS the old
 *     content instead of writing a version.
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

  /**
   * Candidate titles, retitled name first — a library answers to whichever it has.
   *
   * ⚠ This client renames libraries as a matter of course, twice in one afternoon on
   *   2026-08-27, and a title outside these arrays fails SILENTLY in the app: hcAvailable()
   *   goes false and the Highly Confidential level simply stops being offered. Keep these in
   *   step with the HC and archive candidate arrays in src/shared/naming.ts.
   *
   * `role` drives the expectations, because the two halves of a pair want OPPOSITE settings:
   *   approval — moderation ON, Draft Item Security 2
   *   approved — moderation OFF, versioning ON (Auto-route replaces on a name clash)
   *   archive  — moderation OFF; column gaps are the client-deferred 2033 problem, not a fault
   */
  /* ⚠ KEEP IN STEP WITH THE CANDIDATE ARRAYS IN `src/shared/naming.ts`. This client retitles these
     libraries routinely — three times between 2026-08-27 and 2026-08-28 — and a stale list here
     reports a perfectly healthy library as MISSING, which is worse than not checking: it sends
     somebody looking for a fault that is in this script. Read the live titles with
     `/_api/web/lists?$select=Title&$expand=RootFolder&$filter=BaseTemplate eq 101` and add whatever
     is actually there; the URL never changes, so it is what identifies a library across renames. */
  const LIBRARIES = [
    { role: "reference", candidates: ["Approval for Document", "Approval Document", "ApprovalDocument", "Staging"] },
    { role: "approved", candidates: ["Documents", "Restricted & Confidential Document"] },
    { role: "approval", candidates: ["Approval for Highly Confidential Document", "HC Approval Document", "HCApprovalDocument"] },
    { role: "approved", candidates: ["Highly Confidential Document", "HC Documents", "HC Document", "HCDocuments"] },
    { role: "archive", candidates: ["Archive", "Archive Restricted & Confidential Document", "CRSArchive"] },
    { role: "archive", candidates: ["Archive Highly Confidential Document", "HC Archive", "HC Archive Document", "HCArchive"] },
  ];

  /** The three managed-metadata columns. Their term sets are read off the reference, never listed. */
  const TAXONOMY = ["Document_x0020_Type", "Year", "Confidentiality_x0020_Level"];

  /**
   * Excused everywhere.
   *
   * `Full_x0020_Name` is absent from the archives by design — they carry `FullName`, created
   * through the modern Add column button, which strips the space instead of encoding it.
   * pickFullNameField accepts either, and creating the encoded twin would give it TWO
   * candidates and push it into the branch that refuses to write at all. The Full Name section
   * below checks the column actually resolves; this stops the column diff reporting it as a gap.
   */
  const EXCUSED = ["Full_x0020_Name"];

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
    `${SITE}/_api/web/lists?$select=Title,EnableModeration,DraftVersionVisibility,EnableVersioning,MajorVersionLimit,ItemCount&$filter=BaseTemplate eq 101`,
  )
    .then((r) => (r.ok ? r.json() : { value: [] }))
    .then((d) => d.value || []);

  const match = (candidates) =>
    lists.filter((l) => candidates.some((c) => c.toLowerCase() === (l.Title || "").toLowerCase()))[0];

  /**
   * A taxonomy column's companion note field is named from a GUID and is created per LIST, so
   * it legitimately differs between libraries. Comparing them yields three false gaps per
   * library — exactly the noise this rewrite exists to remove.
   */
  const isTaxNote = (n) => /^[a-z][0-9a-f]{31}$/i.test(n);

  /**
   * ⚠ `$top=5000`, NOT 500. This read was capped at 500 and searched in code, which is the
   *   trap memory `sp-capped-read-reads-as-absent` records: a document library carries
   *   hundreds of fields, the newest sit at the END of the collection, and a truncated read is
   *   indistinguishable from an absent column. That would report a correct library as missing
   *   whichever columns were added most recently — which is always the ones under suspicion.
   */
  const readFields = async (title) => {
    const res = await get(
      `${listBase(title)}/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField&$top=5000`,
    );
    if (!res.ok) return { __err: res.status };
    return (await res.json()).value || [];
  };

  /* ── The reference ─────────────────────────────────────────────────────────── */
  const refDef = LIBRARIES[0];
  const refList = match(refDef.candidates);
  if (!refList) {
    fail(`no reference library titled ${refDef.candidates.map((c) => `"${c}"`).join(" or ")} — nothing can be checked`);
    console.log(`        libraries here: ${lists.map((l) => l.Title).join(", ") || "(none)"}`);
    return;
  }
  const refFields = await readFields(refList.Title);
  if (refFields.__err) {
    fail(`could not read ${refList.Title}'s columns (HTTP ${refFields.__err}) — nothing can be checked`);
    return;
  }
  const EXPECTED = refFields.filter((f) => !f.ReadOnlyField && !isTaxNote(f.InternalName));

  /* Bindings expected on every other library, read off the reference rather than listed. A
     field's TermSetId exists only on TaxonomyField, so this is read WITHOUT a $select:
     naming it against the base type fails the whole request instead of omitting the key. */
  const refBindings = {};
  for (const n of TAXONOMY) {
    const res = await get(`${listBase(refList.Title)}/fields/getbyinternalnameortitle('${encodeURIComponent(n)}')`);
    if (!res.ok) { refBindings[n] = undefined; continue; }
    const f = await res.json();
    refBindings[n] = {
      termSet: (f.TermSetId || "").toLowerCase(),
      type: f.TypeAsString,
      multi: f.AllowMultipleValues,
    };
  }

  console.log(
    `%cReference: ${refList.Title} — ${EXPECTED.length} columns. The other five are checked AGAINST it,\n` +
      `so nothing here can tell you the reference itself is correct.`,
    "font-weight:bold;font-size:13px",
  );

  for (const { role, candidates } of LIBRARIES) {
    const list = match(candidates);

    console.log(
      `%c\n${candidates[0]}${role === "reference" ? "  (reference)" : ""}`,
      "font-weight:bold;font-size:14px",
    );
    if (!list) {
      /* An absent ARCHIVE is a valid state — the pair mirrors the site's HC state, and a site
         with no archive libraries simply has the feature inert. An absent HC library is not. */
      if (role === "archive") {
        warn(`no library titled ${candidates.map((c) => `"${c}"`).join(" or ")} — archiving is inert here, which is a valid state`);
      } else {
        fail(`no library titled ${candidates.map((c) => `"${c}"`).join(" or ")}`);
      }
      continue;
    }
    const title = list.Title;
    if (title !== candidates[0]) {
      warn(`titled "${title}" — a title outside the candidate arrays in naming.ts fails SILENTLY, so keep them in step`);
    }

    /* ── Columns ─────────────────────────────────────────────────────────────
       Matched on INTERNAL name, never the display title. A right-looking title over a
       wrong internal name is the failure this whole check exists for. */
    if (role !== "reference") {
      const fields = await readFields(title);
      if (fields.__err) {
        fail(`could not read its columns (HTTP ${fields.__err})`);
        continue;
      }
      const byName = {};
      fields.forEach((f) => { byName[(f.InternalName || "").toLowerCase()] = f; });

      const missing = [];
      const wrongType = [];
      for (const want of EXPECTED) {
        if (EXCUSED.indexOf(want.InternalName) > -1) continue;
        const f = byName[want.InternalName.toLowerCase()];
        if (!f) { missing.push(`${want.InternalName} (${want.TypeAsString})`); continue; }
        // TaxonomyFieldTypeMulti is wrong here too: these are single-value fields, and the
        // multi variant writes a different value shape.
        if ((f.TypeAsString || "") !== want.TypeAsString) {
          wrongType.push(`${want.InternalName} is ${f.TypeAsString}, expected ${want.TypeAsString}`);
        }
      }

      if (missing.length === 0) ok(`all ${EXPECTED.length - EXCUSED.length} columns present`);
      else if (role === "archive") {
        /* The archive column gap is the client's deferred 2033 problem (2026-08-23), not a
           provisioning fault — nothing archives for seven years. Reported so it is not
           forgotten, at a severity that does not compete with a live defect. */
        warn(`${missing.length} column(s) missing — the DEFERRED archive gap, harmless until anything actually archives: ${missing.join(", ")}`);
      } else fail(`${missing.length} column(s) MISSING: ${missing.join(", ")}`);

      if (wrongType.length > 0) for (const w of wrongType) fail(w);
      else if (missing.length === 0) ok("every column is the right type");

      /* ── Taxonomy bindings ─────────────────────────────────────────────────
         A column bound to the wrong term set offers the wrong dropdown and stores terms
         nothing else on the site recognises — and it looks completely correct. */
      for (const n of TAXONOMY) {
        const wanted = refBindings[n];
        if (!wanted) { warn(`${n}: not readable on the reference, so nothing to compare against`); continue; }
        if (!byName[n.toLowerCase()]) continue; // already reported as missing
        const res = await get(`${listBase(title)}/fields/getbyinternalnameortitle('${encodeURIComponent(n)}')`);
        if (!res.ok) { warn(`${n}: could not read its binding (HTTP ${res.status})`); continue; }
        const f = await res.json();
        const got = (f.TermSetId || "").toLowerCase();
        if (!got) warn(`${n}: no term set recorded — is it really Managed Metadata?`);
        else if (got !== wanted.termSet) fail(`${n} is bound to ${got}, but ${refList.Title} uses ${wanted.termSet}`);
        else if (f.AllowMultipleValues !== wanted.multi) {
          fail(`${n} multi=${f.AllowMultipleValues}, reference has ${wanted.multi} — the multi variant writes a different value shape`);
        } else ok(`${n} bound to the same term set as ${refList.Title}`);
      }
    }

    /* ── Full Name resolves ──────────────────────────────────────────────────
       Mirrors pickFullNameField, including that one candidate wins outright and several
       need an exact "Full Name" title. Unresolved means folders keep only their
       abbreviation, with no way to see the real term label — which matters most on a site
       whose folders ARE abbreviations, as SDG's are. */
    const ff = await readFields(title);
    if (!ff.__err) {
      const key = (s) => (s || "").replace(/_x0020_/gi, "").replace(/[^a-z0-9]/gi, "").toLowerCase();
      const cands = ff.filter(
        (f) =>
          (key(f.Title) === "fullname" || key(f.InternalName) === "fullname") &&
          !f.ReadOnlyField &&
          (f.TypeAsString === "Text" || f.TypeAsString === "Note"),
      );
      const exact = cands.filter((f) => (f.Title || "").trim().toLowerCase() === "full name")[0];
      const resolved = cands.length === 1 ? cands[0].InternalName : exact && exact.InternalName;
      if (resolved) ok(`Full Name resolves to ${resolved}`);
      else fail(`Full Name UNRESOLVED (${cands.length} candidate(s)) — folders will show only their abbreviation`);
    }

    /* ── Moderation ──────────────────────────────────────────────────────────
       Opposite settings on the two halves, and BOTH failures stay invisible until
       somebody cannot find a document. */
    const wantsModeration = role === "approval" || role === "reference";
    if (wantsModeration) {
      if (list.EnableModeration) ok("content approval is ON");
      else fail("content approval is OFF — it must be ON, or nothing is ever approved");
      /* 2 = "Only users who can approve items (and the author)". Anything else lets a PIC read
         a peer's pending drafts. 0 was a deliberate client decision on 2026-08-19 that was
         then never actually applied, so this is REPORTED rather than failed — but it must
         never be a surprise, which is the whole lesson of that episode. */
      if (list.DraftVersionVisibility === 2) ok("Draft Item Security = approver and author");
      else {
        warn(
          `Draft Item Security is ${list.DraftVersionVisibility}, not 2 — every uploader in a unit ` +
            "can see every pending file. Deliberate on some sites; confirm it is deliberate here",
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

    /* ── Versioning ──────────────────────────────────────────────────────────
       Auto-route's Copy file uses nameConflictBehavior 1 (Replace) since 2026-08-27, at the
       client's explicit decision. Versioning is the ONLY thing that then makes a replacement
       recoverable: the old content becomes a version instead of being destroyed. */
    if (role === "approved") {
      if (list.EnableVersioning) {
        ok(`version history ON (keeps ${list.MajorVersionLimit || "unlimited"}) — a Replace writes a version`);
      } else {
        fail("version history OFF — Auto-route replaces on a name clash, so the previous content would be DESTROYED with no trail");
      }
    }

    console.log(`        ${list.ItemCount} item(s)`);
  }

  console.log(
    failures === 0
      ? "%c\nAll six libraries agree with the reference."
      : `%c\n${failures} thing(s) to fix above.`,
    `font-weight:bold;font-size:13px;color:${failures === 0 ? "#0f6c3f" : "#a4262c"}`,
  );
  console.log(
    "Not checked here, because none of it is readable from a library: the Power Automate flows, " +
      "the groups and their Group Map rows, whether reconciliation has run, the CRS lists and their " +
      "permissions, the config rows, and the segment Levels chains.",
  );
})();

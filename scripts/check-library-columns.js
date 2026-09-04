/**
 * Compares the columns across all six CRS libraries, and prints a fingerprint for cross-site diffing.
 *
 * READ ONLY. One GET per library. Paste into the browser console on any page of a CRS site.
 *
 * ── The invariant it checks ──────────────────────────────────────────────────
 * A SharePoint move or copy carries over ONLY the columns that EXIST at the destination. The rest
 * vanish with no error and a green run. That is the most expensive class of bug in this project's
 * history:
 *
 *   - `Documents` was missing Remark, LegallyPrivileged, ProjectName and Vendor/CustomerName for
 *     MONTHS. Auto-route dropped all four off every approved document — including a legal marker —
 *     and separately bulk upload tagged NOTHING at all, because one unknown field name fails the
 *     WHOLE validateUpdateListItem call, losing every column and not only the missing ones. That
 *     call answers HTTP 200 either way.
 *   - `SubmissionId` / `BatchId` had to be added to all four libraries or a submission would lose
 *     its files one at a time AS THEY WERE APPROVED.
 *
 * So the rule is: whatever `Approval Document` has, every downstream library must have too, under the
 * SAME internal name and type.
 *
 * ── Why nothing is hardcoded ─────────────────────────────────────────────────
 * `Approval Document` is the reference, and the expected set is READ FROM IT. A literal list of
 * column names would be right today and silently wrong after the next segment is onboarded, because
 * each new segment creates its own tier columns (Region/RegionTid, EstateMill/EstateMillTid, …). That
 * exact mistake — a two-element literal where a derived list belonged — cost four days on the HC pair.
 *
 * ── Known, legitimate differences (not faults) ───────────────────────────────
 *   - `Full Name` exists as `Full_x0020_Name` on libraries created recently and as `FullName0` on
 *     older ones. `pickFullNameField` resolves either at runtime, by display title OR internal name,
 *     so a mismatch here is expected and is not reported as a problem.
 *   - `Archive` / `HC Archive` are KNOWN to be missing most metadata columns. Mirroring `Documents`
 *     into them is deferred until after the SDG migration (nothing archives until 2033), so a long
 *     MISSING list there is the documented state rather than a new discovery.
 *   - The three TAXONOMY columns (Document Type, Year, Confidentiality Level) cannot be created by
 *     script safely: a taxonomy field needs its hidden note field and a term-set binding, and a wrong
 *     binding yields a column that looks right and silently tags nothing. They are reported
 *     separately, for creating BY HAND.
 */
(async () => {
  /* THE SITE URL COMES FROM SHAREPOINT, NOT FROM THE ADDRESS BAR.
     This was originally derived by splitting `location.href` on "/_layouts" and "/SitePages" — which
     silently produces a WRONG base on any other kind of page (a /Lists/... or /Forms/... URL contains
     neither), and the symptom is not a clean 404 but an HTML sign-in/error page parsed as JSON:
     `Unexpected token '<', "<!DOCTYPE"...`. `_spPageContextInfo` is set by SharePoint on every page,
     so it is right wherever this is pasted. The split stays only as a last resort. */
  const spx = window._spPageContextInfo || {};
  const site = (spx.webAbsoluteUrl
    || (spx.webServerRelativeUrl ? location.origin + spx.webServerRelativeUrl : "")
    || location.href.split("/_layouts")[0].split("/SitePages")[0]).replace(/\/$/, "");
  console.log("site: " + site);

  /* The approval library FIRST — it is the reference every other library is compared against.
     ⚠ TITLES, NOT URLS, AND THIS CLIENT RENAMES THEM. Listed per library as alternatives, newest
     first; the script takes whichever resolves. A name missing here reads as a missing LIBRARY. */
  const LIB_CANDIDATES = [
    ["Approval for Document", "Approval Document"],
    ["Restricted & Confidential Document", "Documents"],
    ["Approval for Highly Confidential Document", "HC Approval Document"],
    ["Highly Confidential Document", "HC Documents", "HC Document"],
    ["Archive Restricted & Confidential Document", "Archive"],
    ["Archive Highly Confidential Document", "HC Archive"],
  ];

  // Deferred by design; a long MISSING list here is expected, not news. Keyed on every spelling,
  // because the resolved title is whichever candidate answered.
  const DEFERRED = {
    "Archive": true,
    "Archive Restricted & Confidential Document": true,
    "HC Archive": true,
    "Archive Highly Confidential Document": true,
  };

  // Managed metadata: report, never offer to script.
  const TAXONOMY = { "document_x0020_type": true, "year": true, "confidentiality_x0020_level": true };

  // Resolved at runtime by pickFullNameField, so either spelling is correct.
  const FULLNAME = { "full_x0020_name": true, "fullname0": true, "fullname": true };

  const BUILT_IN = {};
  ("title,id,contenttype,modified,created,author,editor,_uiversionstring,attachments,edit,"
    + "linktitlenomenu,linktitle,linkfilename,linkfilenamenomenu,docicon,fileleafref,fileref,"
    + "filedirref,filesizedisplay,itemchildcount,folderchildcount,_complianceflags,_compliancetag,"
    + "_compliancetaguserid,_compliancetagwrittentime,_islatest,appauthor,appeditor,_colortag,"
    + "complianceassetid,_shortcuturl,_shortcutsiteid,_shortcutwebid,_shortcutuniqueid,"
    + "_activitystream,_commentcount,_commentflags,_displayname,_extendeddescription,_sourceurl,"
    + "_sharedfileindex,templateurl,xd_progid,xd_signature,sortbehavior,parentversionstring,"
    + "parentleafname,_hascopydestinations,_copysource,_level,owshiddenversion,_uiversion,"
    + "syncclientid,progid,streamhash,checkedouttitle,checkedoutuserid,checkinComment,virusstatus,"
    + "_virusstatus,_virusvendorid,_virusinfo,_stubfile,_hasencryptedcontent,_parcelsize,"
    + "originatorid,noexecute,contentversion,_listschemaversion,_dirty,_parsable,"
    + "_ip_unifiedcompliancepolicyproperties,_ip_unifiedcompliancepolicyuitags,"
    + "_modernaudiencetargetusersfield,accesspolicy").split(",").forEach((n) => { BUILT_IN[n] = true; });

  const KIND = { 2: "Text", 3: "Note", 4: "DateTime", 8: "Boolean", 9: "Number", 20: "User" };

  const read = async (title) => {
    const url = site + "/_api/web/lists/getbytitle('" + encodeURIComponent(title) + "')"
      + "/fields?$select=Title,InternalName,FieldTypeKind,TypeAsString,Hidden,ReadOnlyField&$top=500";
    const res = await fetch(url, { headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) return { status: res.status, fields: undefined };
    return { status: 200, fields: (await res.json()).value || [] };
  };

  /** Author-created columns only, keyed by lower-cased internal name. */
  const authored = (fields) => {
    const out = {};
    for (const f of fields) {
      const n = (f.InternalName || "").toLowerCase();
      if (!n || f.Hidden || f.ReadOnlyField || BUILT_IN[n]) continue;
      // A taxonomy column brings a hidden note field of its own; skip those explicitly.
      if (/taxcatchall/i.test(n) || /_0$/.test(n)) continue;
      out[n] = f;
    }
    return out;
  };

  /* Resolve each library to whichever of its candidate titles actually answers.
     ⚠ A 404 ON ONE CANDIDATE IS NOT AN ABSENT LIBRARY — it is the wrong name. Only when EVERY
     candidate 404s is the library genuinely missing, and the last attempt's status is what gets
     reported so the message names a real read rather than a guess. */
  const results = {};
  const LIBS = [];
  for (const candidates of LIB_CANDIDATES) {
    let picked = null;
    let last = null;
    for (const title of candidates) {
      const out = await read(title);
      last = { title: title, out: out };
      if (out.fields !== undefined) { picked = { title: title, out: out }; break; }
    }
    const chosen = picked || last;
    LIBS.push(chosen.title);
    results[chosen.title] = chosen.out;
    if (picked && picked.title !== candidates[0]) {
      console.log("%c" + candidates[0] + " -> resolved as '" + picked.title + "'", "color:#605e5c");
    }
  }

  const refTitle = LIBS[0];
  const ref = results[refTitle];
  if (ref.fields === undefined) {
    console.log("%cCannot continue: the approval library could not be read (HTTP " + ref.status
      + "). Tried: " + LIB_CANDIDATES[0].join(", "), "color:#a4262c;font-weight:bold");
    console.log("It is the reference every other library is compared against. A 404 on every"
      + " candidate means it is titled something else again — read the live titles with"
      + " /_api/web/lists?$select=Title&$filter=BaseTemplate eq 101 and add it to LIB_CANDIDATES.");
    return;
  }

  const expected = authored(ref.fields);
  const expectedNames = Object.keys(expected);
  console.log("%cReference: " + refTitle + " — " + expectedNames.length + " author-created columns",
    "font-weight:bold;font-size:13px");

  const fingerprints = {};

  for (const title of LIBS) {
    const out = results[title];
    console.log("%c" + title, "font-weight:bold;font-size:13px");

    /* A FAILED read is not an empty library, and the two lead somewhere different. 404 means it does
       not exist on this site — normal for Archive / HC Archive, and for the HC pair on a site with no
       HC. Anything else means we learned nothing about it, and reporting every column as missing
       would be a false alarm on the largest possible scale. */
    if (out.fields === undefined) {
      console.log(out.status === 404
        ? "  does not exist on this site" + (DEFERRED[title] ? " (fine — deferred by design)" : "")
        : "  could not be read (HTTP " + out.status + ") — nothing below is a statement about it");
      continue;
    }

    const have = authored(out.fields);
    fingerprints[title] = Object.keys(have).sort()
      .map((n) => n + ":" + (KIND[have[n].FieldTypeKind] || have[n].TypeAsString)).join(",");

    if (title === "Approval Document") { console.log("  (the reference)"); continue; }

    const missing = [];
    const wrongType = [];
    for (const n of expectedNames) {
      if (FULLNAME[n]) continue;                  // either spelling is correct
      const f = have[n];
      if (!f) { missing.push(expected[n].InternalName); continue; }
      if (f.FieldTypeKind !== expected[n].FieldTypeKind) {
        wrongType.push(expected[n].InternalName + " is "
          + (KIND[f.FieldTypeKind] || f.TypeAsString) + ", expected "
          + (KIND[expected[n].FieldTypeKind] || expected[n].TypeAsString));
      }
    }

    // Full Name counts as present under EITHER spelling, or genuinely absent.
    const refHasFullName = expectedNames.some((n) => FULLNAME[n]);
    const hasFullName = Object.keys(have).some((n) => FULLNAME[n]);
    if (refHasFullName && !hasFullName) missing.push("Full Name (either spelling)");

    if (missing.length === 0 && wrongType.length === 0) {
      console.log("  %cmatches the reference", "color:#0f6c3f;font-weight:bold");
    }
    if (missing.length) {
      const tax = missing.filter((m) => TAXONOMY[m.toLowerCase()]);
      const plain = missing.filter((m) => !TAXONOMY[m.toLowerCase()]);
      const label = DEFERRED[title] ? "MISSING (deferred by design)" : "MISSING";
      const colour = DEFERRED[title] ? "color:#8a4b00" : "color:#a4262c;font-weight:bold";
      if (plain.length) {
        console.log("  %c" + label + " (" + plain.length + "): " + plain.join(", "), colour);
      }
      /* Reported apart because the fix differs: a plain column can be created from its name, a
         taxonomy one needs a term-set binding, and a wrong binding yields a column that looks
         correct and silently tags nothing. */
      if (tax.length) {
        console.log("  %cMISSING TAXONOMY — create BY HAND with the right term set: " + tax.join(", "),
          "color:#a4262c;font-weight:bold");
      }
      if (!DEFERRED[title] && plain.length) {
        console.log("  why it matters: a move/copy carries over only columns that exist here, and one"
          + " unknown field name fails the WHOLE metadata write — losing every column, not just these.");
      }
    }
    if (wrongType.length) {
      console.log("  %cWRONG TYPE: " + wrongType.join(" | "), "color:#a4262c;font-weight:bold");
    }
  }

  console.log("%c\nFINGERPRINTS — run on both sites and diff these lines",
    "font-weight:bold;font-size:13px");
  for (const title of Object.keys(fingerprints)) {
    console.log(title + " (" + fingerprints[title].split(",").length + "): " + fingerprints[title]);
  }
})();

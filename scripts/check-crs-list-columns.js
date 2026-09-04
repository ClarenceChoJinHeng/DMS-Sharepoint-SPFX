/**
 * Checks that CRS Requests and CRS Audit Log have the columns the code writes.
 *
 * Sibling of scripts/check-hc-setup.js, which does the same job for the HC LIBRARY pair.
 *
 * READ ONLY. One GET per list. It writes nothing, so it is safe to run at any time and safe to run
 * repeatedly. Paste it into the browser console on any page of a CRS site.
 *
 * ── Read this before hand-creating anything ──────────────────────────────────
 * BOTH LISTS PROVISION THEMSELVES. Opening the `CRS Requests` or `CRS Audit Log` web part creates
 * the list and its columns, and the Requests page carries an "add missing columns" repair button for
 * a list created before a column was added. So the normal fix for anything reported below is NOT to
 * add the column by hand — it is to open the page and let it build.
 *
 * Hand-creating invites the one failure that cannot be repaired afterwards: SharePoint fixes a
 * column's INTERNAL NAME at creation, permanently, from the title it was created with. A column whose
 * display name looks right over a wrong internal name fails exactly like an absent column — and
 * because one unknown field name fails an entire validateUpdateListItem call, it takes every other
 * column's value with it, on a call that answers HTTP 200 either way.
 *
 * ── Keeping this honest ──────────────────────────────────────────────────────
 * The expected sets are COPIED from source and must stay in step with it:
 *   CRS Audit Log  -> AUDIT_COLUMNS in src/shared/spAuditLog.ts
 *   CRS Requests   -> COLUMNS in src/webparts/requests/components/Requests.tsx
 * A second definition of a schema is a thing that drifts, and the drifting copy is always the one
 * nobody reads. If a column is added to either list in code, add it here in the same change.
 *
 * FieldTypeKind: 2 = Text, 3 = Note (multi-line), 4 = DateTime.
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

  const EXPECTED = {
    "CRS Audit Log": {
      // name -> [FieldTypeKind, shouldBeIndexed]
      EventTime: [4, true],
      EventType: [2, true],
      Outcome: [2, false],
      ActorName: [2, false],
      ActorEmail: [2, true],
      Source: [2, false],
      LibraryName: [2, false],
      ItemUniqueId: [2, true],
      ItemName: [2, false],
      ItemPath: [3, false],
      Segment: [2, false],
      UnitPath: [2, false],
      Details: [3, false],
    },
    "CRS Requests": {
      RequestType: [2, false],
      Stage: [2, false],
      Status: [2, false],
      ItemUniqueId: [2, false],
      ItemName: [2, false],
      ItemUrl: [2, false],
      Segment: [2, false],
      Unit: [2, false],
      UnitTermGuid: [2, false],
      RequestedBy: [2, false],
      RequestedAt: [4, false],
      Reason: [3, false],
      ShareWith: [2, false],
      SharePermission: [2, false],
      ExpiresAt: [4, false],
      DecidedBy: [2, false],
      DecidedAt: [4, false],
      DecisionNote: [3, false],
    },
  };

  const KIND = { 2: "Text", 3: "Note", 4: "DateTime" };

  const BUILT_IN = new Set(["title", "id", "contenttype", "modified", "created", "author", "editor",
    "_uiversionstring", "attachments", "edit", "linktitlenomenu", "linktitle", "docicon",
    "itemchildcount", "folderchildcount", "_complianceflags", "_compliancetag",
    "_compliancetaguserid", "_compliancetagwrittentime", "_islatest", "appauthor", "appeditor",
    "_colortag", "complianceassetid", "_extendeddescription"]);

  const read = async (title) => {
    const url = site + "/_api/web/lists/getbytitle('" + encodeURIComponent(title) + "')"
      + "/fields?$select=Title,InternalName,FieldTypeKind,Hidden,ReadOnlyField,Indexed&$top=500";
    const res = await fetch(url, { headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) return { status: res.status, fields: undefined };
    return { status: 200, fields: (await res.json()).value || [] };
  };

  for (const title of Object.keys(EXPECTED)) {
    const expected = EXPECTED[title];
    const out = await read(title);
    console.log("%c" + title, "font-weight:bold;font-size:13px");

    /* A FAILED read is not an empty list, and the difference decides what someone does next.
       404 means the list does not exist yet — the normal state until its page has been opened once.
       Reporting "18 columns missing" there would send an admin to hand-create exactly what the page
       creates correctly. Anything else means we learned nothing about the list at all. */
    if (out.fields === undefined) {
      console.log(out.status === 404
        ? "  list does not exist yet — open its web part once and it provisions itself"
        : "  could not read it (HTTP " + out.status + ") — nothing below is a statement about it");
      continue;
    }

    // Author-created columns only; hidden, read-only and system fields are noise here.
    const have = {};
    for (const f of out.fields) {
      if (f.Hidden || f.ReadOnlyField) continue;
      have[(f.InternalName || "").toLowerCase()] = f;
    }

    const missing = [];
    const wrongType = [];
    const notIndexed = [];

    for (const name of Object.keys(expected)) {
      const kind = expected[name][0];
      const wantIndex = expected[name][1];
      const f = have[name.toLowerCase()];
      if (!f) { missing.push(name); continue; }
      if (f.FieldTypeKind !== kind) {
        wrongType.push(name + " is " + (KIND[f.FieldTypeKind] || f.FieldTypeKind)
          + ", expected " + KIND[kind]);
      }
      if (wantIndex && f.Indexed === false) notIndexed.push(name);
    }

    const known = {};
    for (const n of Object.keys(expected)) known[n.toLowerCase()] = true;
    const extra = Object.keys(have).filter((n) => !known[n] && !BUILT_IN.has(n));

    if (missing.length === 0 && wrongType.length === 0) {
      console.log("  %call " + Object.keys(expected).length + " columns present and correctly typed",
        "color:#0f6c3f;font-weight:bold");
    }
    if (missing.length) {
      console.log("  %cMISSING (" + missing.length + "): " + missing.join(", "),
        "color:#a4262c;font-weight:bold");
      console.log("  fix: open this list's web part and use its provision / add-missing-columns "
        + "button. Do NOT create these by hand — the internal name is fixed at creation.");
    }
    /* ⚠ A WRONG TYPE IS THE ONE THING THE PAGE CANNOT REPAIR. `ensureColumns` skips any column that
       already exists, WHATEVER its type — so a Text column where a Note is expected stays wrong for
       ever and silently truncates long values. Repair means delete and re-create, losing its data,
       which is why this is reported separately from "missing" rather than lumped in with it. */
    if (wrongType.length) {
      console.log("  %cWRONG TYPE: " + wrongType.join(" | "), "color:#a4262c;font-weight:bold");
      console.log("  fix: the page will NOT correct this — it skips columns that already exist. "
        + "Delete and re-create it, accepting the loss of its data.");
    }
    if (notIndexed.length) {
      // Performance only. It starts to matter past the 5,000-item list view threshold, which an
      // append-only audit log reaches and a requests list probably never does.
      console.log("  not indexed (performance only, past ~5,000 rows): " + notIndexed.join(", "));
    }
    if (extra.length) console.log("  extra, usually harmless: " + extra.join(", "));
  }

  console.log("%cRun on BOTH sites and compare the summary lines.",
    "color:#666;font-style:italic");
})();

/**
 * Creates the missing columns on CRS Requests and CRS Audit Log.
 *
 * Paste into the browser console on any page of a CRS site. WRITES — it POSTs field definitions. It
 * creates no list and no list item, and it never touches a column that already exists.
 *
 * ── Prefer the web part to this ──────────────────────────────────────────────
 * Both lists provision themselves: the `CRS Requests` page has an "add missing columns" button and
 * the `CRS Audit Log` page a provision button, each reusing the same definition the code writes
 * against. THAT is the correct route, because it cannot drift from the code. This script is a SECOND
 * copy of the same schema, and a second copy goes stale — use it only when the page has failed, and
 * update it in the same change as the code if a column is ever added.
 *
 * Source of truth:
 *   CRS Audit Log  -> AUDIT_COLUMNS in src/shared/spAuditLog.ts
 *   CRS Requests   -> COLUMNS in src/webparts/requests/components/Requests.tsx
 *
 * ── Traps this script exists to avoid ────────────────────────────────────────
 *  1. NO `NumberOfLines` ON A NOTE COLUMN. It belongs to `SP.FieldMultiLineText`, not `SP.Field`, and
 *     with `odata=nometadata` SharePoint infers the entity type from the ENDPOINT — so the extra
 *     property is rejected outright with HTTP 400. `Reason` is the first Note column in the Requests
 *     list and was the exact point EVERY provisioning run died on 2026-08-20. FieldTypeKind 3 already
 *     yields a multi-line column; the line count is display-only and defaults sensibly.
 *  2. AN INTERNAL NAME IS FIXED AT CREATION, PERMANENTLY, from the Title it is created with. Every
 *     name here is space-free by design, so Title and internal name coincide — do not "improve" a
 *     Title into something with spaces or punctuation. A column whose display name looks right over a
 *     wrong internal name fails exactly like an absent column, and because one unknown field name
 *     fails an entire validateUpdateListItem call, it takes every other column's value with it, on a
 *     call that answers HTTP 200 either way.
 *  3. TEXT, NEVER CHOICE, for RequestType / Stage / Status / EventType. A value absent from a Choice
 *     column's `Choices` fails the WHOLE write — so the day a new status is added in code, every row
 *     carrying it would be lost silently.
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

  const PLAN = {
    "CRS Audit Log": [
      // [name, FieldTypeKind, indexed]
      ["EventTime", 4, true],
      ["EventType", 2, true],
      ["Outcome", 2, false],
      ["ActorName", 2, false],
      ["ActorEmail", 2, true],
      ["Source", 2, false],
      ["LibraryName", 2, false],
      ["ItemUniqueId", 2, true],
      ["ItemName", 2, false],
      ["ItemPath", 3, false],
      ["Segment", 2, false],
      ["UnitPath", 2, false],
      ["Details", 3, false],
    ],
    "CRS Requests": [
      ["RequestType", 2, false],
      ["Stage", 2, false],
      ["Status", 2, false],
      ["ItemUniqueId", 2, false],
      ["ItemName", 2, false],
      ["ItemUrl", 2, false],
      ["Segment", 2, false],
      ["Unit", 2, false],
      ["UnitTermGuid", 2, false],
      ["RequestedBy", 2, false],
      ["RequestedAt", 4, false],
      ["Reason", 3, false],
      ["ShareWith", 2, false],
      ["SharePermission", 2, false],
      ["ExpiresAt", 4, false],
      ["DecidedBy", 2, false],
      ["DecidedAt", 4, false],
      ["DecisionNote", 3, false],
    ],
  };

  /* The digest call is checked BEFORE parsing, because SharePoint answers a wrong site URL or an
     expired session with an HTML page, not JSON — and `r.json()` on that throws
     `Unexpected token '<', "<!DOCTYPE"...`, which names neither the URL nor the real problem. */
  const dRes = await fetch(site + "/_api/contextinfo", {
    method: "POST", headers: { Accept: "application/json;odata=nometadata" },
  });
  if (!dRes.ok || (dRes.headers.get("content-type") || "").indexOf("json") === -1) {
    console.log("%cCould not get a form digest from " + site + " (HTTP " + dRes.status + ").",
      "color:#a4262c;font-weight:bold");
    console.log("The response was not JSON, which usually means the site URL is wrong or the session "
      + "has expired. Reload a page on the site and run this again. Nothing was created.");
    return;
  }
  const digest = (await dRes.json()).FormDigestValue;

  const listUrl = (title) =>
    site + "/_api/web/lists/getbytitle('" + encodeURIComponent(title) + "')";

  /** Existing field internal names, lower-cased. `names === undefined` means the read FAILED. */
  const existing = async (title) => {
    const res = await fetch(listUrl(title) + "/fields?$select=InternalName&$top=500",
      { headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) return { status: res.status, names: undefined };
    const v = (await res.json()).value || [];
    const set = {};
    for (const f of v) set[(f.InternalName || "").toLowerCase()] = true;
    return { status: 200, names: set };
  };

  const post = (url, body, extraHeaders) =>
    fetch(url, {
      method: "POST",
      headers: Object.assign({
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
      }, extraHeaders || {}),
      body: JSON.stringify(body),
    });

  for (const title of Object.keys(PLAN)) {
    console.log("%c" + title, "font-weight:bold;font-size:13px");
    const have = await existing(title);

    /* A FAILED READ MUST NOT LICENSE A WRITE. If we cannot see what the list already has, creating
       columns blind risks a duplicate-name 400 at best. 404 means the list does not exist — and this
       script deliberately does NOT create it, because the web part creates it with the right template
       and settings, and a list made here would differ in ways nobody would think to check. */
    if (have.names === undefined) {
      console.log(have.status === 404
        ? "  %clist does not exist. Open its web part and press provision — that creates the list AND "
          + "its columns. This script only fills in columns on a list that already exists."
        : "  %ccould not read its existing columns (HTTP " + have.status + "). Nothing was created — "
          + "creating blind risks duplicates.", "color:#a4262c;font-weight:bold");
      continue;
    }

    const created = [];
    const skipped = [];
    const failed = [];
    const toIndex = [];

    for (const row of PLAN[title]) {
      const name = row[0];
      const kind = row[1];
      const indexed = row[2];
      if (have.names[name.toLowerCase()]) {
        skipped.push(name);
        /* Already there — but possibly with the WRONG TYPE, which this script cannot fix and must not
           pretend to have checked. check-crs-list-columns.js is what reports that case. */
        if (indexed) toIndex.push(name);
        continue;
      }
      // MINIMAL BODY. See trap 1 — any extra property is rejected for the inferred entity type.
      const res = await post(listUrl(title) + "/fields", { Title: name, FieldTypeKind: kind });
      if (res.ok) {
        created.push(name);
        if (indexed) toIndex.push(name);
      } else {
        let detail = "";
        try { detail = (await res.text()).slice(0, 300); } catch (e) { detail = ""; }
        failed.push(name + " (HTTP " + res.status + ") " + detail);
      }
    }

    /* Indexing is a SEPARATE MERGE, after creation, and is performance only — it starts to matter
       past the 5,000-item list view threshold, which an append-only audit log reaches and a requests
       list probably never does. So a failure here is reported and never treated as fatal: the column
       itself is correct and the list works. */
    const indexFailed = [];
    for (const name of toIndex) {
      const res = await post(listUrl(title) + "/fields/getbyinternalnameortitle('" + name + "')",
        { Indexed: true }, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" });
      if (!res.ok) indexFailed.push(name + " (HTTP " + res.status + ")");
    }

    if (created.length) {
      console.log("  %ccreated (" + created.length + "): " + created.join(", "),
        "color:#0f6c3f;font-weight:bold");
    }
    if (skipped.length) console.log("  already there (" + skipped.length + "): " + skipped.join(", "));
    if (failed.length) {
      console.log("  %cFAILED (" + failed.length + "): " + failed.join(" | "),
        "color:#a4262c;font-weight:bold");
      console.log("  If a Note column (ItemPath, Details, Reason, DecisionNote) failed with 400, the "
        + "response body above names the rejected property — that is trap 1 in this file's header.");
    }
    if (indexFailed.length) {
      console.log("  could not index (performance only, not a fault): " + indexFailed.join(", "));
    }
  }

  console.log("%cNow re-run scripts/check-crs-list-columns.js to confirm.",
    "color:#666;font-style:italic");
})();

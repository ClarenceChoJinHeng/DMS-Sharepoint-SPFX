/**
 * Read-only browser-console script. Prints the list GUID (Id) for any list/library whose title
 * matches (case-insensitive, partial match) — needed for building Power Automate URIs like
 * `_api/web/lists(guid'...')/items(...)`, which use the GUID rather than the title.
 *
 * Usage: open any CRS page in the browser, open DevTools console, paste this whole file, press
 * Enter. Nothing is written — one GET.
 *
 * Edit MATCH below to whatever you're looking for, e.g. "Archive" or "Submissions".
 */
(async () => {
  const MATCH = "Archive"; // <-- change this, then re-paste and run again for the next one

  const spx = window._spPageContextInfo || {};
  const site = (
    spx.webAbsoluteUrl
    || (spx.webServerRelativeUrl ? location.origin + spx.webServerRelativeUrl : "")
    || location.href.split("/_layouts")[0].split("/SitePages")[0]
  ).replace(/\/$/, "");
  console.log("site: " + site);

  const url = site + "/_api/web/lists?$select=Title,Id&$filter=BaseTemplate eq 101 or BaseTemplate eq 100";
  const res = await fetch(url, { headers: { Accept: "application/json;odata=nometadata" } });
  if (!res.ok) {
    console.error("Could not list lists/libraries — HTTP " + res.status);
    return;
  }
  const all = (await res.json()).value || [];
  const hits = all.filter((l) => (l.Title || "").toLowerCase().indexOf(MATCH.toLowerCase()) > -1);

  if (hits.length === 0) {
    console.log("No list/library title contains \"" + MATCH + "\". Every title on this site:");
    all.forEach((l) => console.log("  " + l.Title));
    return;
  }
  console.log("Matches for \"" + MATCH + "\":");
  hits.forEach((l) => console.log("  " + l.Title + "  ->  " + l.Id));
})();

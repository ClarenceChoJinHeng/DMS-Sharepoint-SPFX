/*
 * delete-submission-duplicate-columns.js — Browser console, run as a SITE ADMINISTRATOR.
 *
 * WHY THIS EXISTS
 * ----------------
 * A now-fixed bug in reconciliation used to create fresh duplicate columns on `CRS Submissions`
 * every run (e.g. `SubmissionRef0`, `SubmissionRef1`, `BatchRef0`, `ArchivedAt2`) because
 * SharePoint doesn't reject a duplicate DISPLAY name — it silently invents a unique INTERNAL name
 * by appending a digit. The fix stops NEW ones being created; it does not clean up the ones
 * already sitting there. This script finds and (optionally) deletes exactly those leftovers.
 *
 * ⚠⚠ DELETING A COLUMN IS PERMANENT. It does NOT go to the recycle bin, and any value ever stored
 * in it is destroyed with it. This script defaults to a DRY RUN for that reason — read the list it
 * prints, and only set DRY_RUN = false once you're satisfied every flagged field is genuinely a
 * duplicate.
 *
 * HOW A FIELD IS FLAGGED (deliberately strict, no hardcoded column names to trust from memory):
 *   1. Its internal name ends in one or more digits, e.g. "SubmissionRef0".
 *   2. Stripping those trailing digits gives a name that ALSO EXISTS on this same list, EXACTLY,
 *      with no digits — e.g. "SubmissionRef" is a real field here too.
 *   Only fields satisfying BOTH are flagged. A field whose "base" name does not independently
 *   exist is left alone — that is real client data, not a duplicate, whatever its name looks like.
 *   SharePoint's own built-in/system fields are excluded outright (see byInternalName filter below).
 *
 * HOW TO RUN
 *   1. Set SITE_OVERRIDE below if needed, and LIST_TITLE if this client has renamed the list.
 *   2. Open the site as a SITE ADMINISTRATOR, F12 -> Console, paste this whole file, Enter.
 *   3. Read the printed table. If it looks right, come back, change DRY_RUN to false, and run it
 *      again — nothing is deleted on the dry run, whatever the count says.
 */
(async () => {
  const SITE_OVERRIDE = ""; // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const LIST_TITLE = "CRS Submissions"; // change if this client has renamed the list
  const DRY_RUN = true; // ⚠ set to false ONLY after reviewing the dry-run output below

  const derived = (() => {
    const m = location.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    return location.origin + (m ? m[1] : "");
  })();
  const web = (SITE_OVERRIDE || derived).replace(/\/+$/, "");
  console.log("reading site:", web || "(root)", "| list:", LIST_TITLE, "| DRY_RUN:", DRY_RUN);

  const get = async (url) => {
    const r = await fetch(web + url, { headers: { Accept: "application/json;odata=nometadata" } });
    if (!r.ok) throw new Error(`GET ${url} -> HTTP ${r.status}`);
    return r.json();
  };

  const listUrl = `/_api/web/lists/getbytitle('${encodeURIComponent(LIST_TITLE)}')`;

  // ── Read every field, including hidden ones, so the base-name check is complete. ────────────
  let fields;
  try {
    const res = await get(
      `${listUrl}/fields?$select=Id,Title,InternalName,Hidden,FromBaseType,ReadOnlyField&$top=5000`,
    );
    fields = res.value || [];
  } catch (e) {
    console.log("Could not read fields on", LIST_TITLE, "-", String(e));
    return;
  }
  console.log(`${fields.length} field(s) total on ${LIST_TITLE}.`);

  const byInternalName = new Map(fields.map((f) => [f.InternalName, f]));

  // SharePoint's own built-in fields are never candidates — only ours can have this shape.
  const candidates = fields.filter((f) => !f.Hidden && !f.FromBaseType);

  const duplicates = [];
  for (const f of candidates) {
    const m = /^(.+?)(\d+)$/.exec(f.InternalName);
    if (!m) continue;
    const base = m[1];
    const baseField = byInternalName.get(base);
    if (baseField && !baseField.Hidden) {
      duplicates.push({ Id: f.Id, Title: f.Title, InternalName: f.InternalName, base });
    }
  }

  console.log(`\n${duplicates.length} duplicate column(s) found:\n`);
  console.table(
    duplicates.map((d) => ({ InternalName: d.InternalName, DisplayTitle: d.Title, RealFieldIs: d.base })),
  );

  if (duplicates.length === 0) {
    console.log("Nothing to do.");
    return;
  }

  if (DRY_RUN) {
    console.log(
      "\nDRY RUN — nothing was deleted. Review the table above. If it looks right, set " +
        "DRY_RUN = false at the top of this script and run it again to actually delete these.",
    );
    return;
  }

  // ── Actual deletion pass. Requires a request digest for the write. ──────────────────────────
  const digestRes = await fetch(web + "/_api/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" },
  });
  if (!digestRes.ok) {
    console.log("Could not get a request digest — nothing was deleted. HTTP", digestRes.status);
    return;
  }
  const digestJson = await digestRes.json();
  const digest = digestJson.FormDigestValue;

  let ok = 0;
  let failed = 0;
  for (const d of duplicates) {
    try {
      const r = await fetch(`${web}${listUrl}/fields('${d.Id}')`, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          "X-RequestDigest": digest,
          "X-HTTP-Method": "DELETE",
        },
      });
      if (r.ok || r.status === 204) {
        console.log(`✓ deleted ${d.InternalName}`);
        ok++;
      } else {
        console.log(`✗ ${d.InternalName} — HTTP ${r.status}`);
        failed++;
      }
    } catch (e) {
      console.log(`✗ ${d.InternalName} — ${String(e)}`);
      failed++;
    }
  }
  console.log(`\nDone: ${ok} deleted, ${failed} failed.`);
})();

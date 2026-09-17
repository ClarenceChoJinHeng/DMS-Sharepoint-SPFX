/*
 * Why does an APPROVER still get the reason-required request dialog?
 * ------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, deleted or granted.
 *
 * Run it in the browser console, ON THE CRS SITE, SIGNED IN AS THE ACCOUNT THAT IS SEEING THE
 * PROBLEM (the approver — not an administrator; an administrator bypasses the whole check and will
 * always look fine).
 *
 * It answers the one question the code cannot answer from here: when My Submissions decides whether
 * to offer "Delete" (act now) or "Request deletion" (ask an approver), it matches the DOCUMENT'S OWN
 * tier term GUIDs against the tier GUIDs on THIS VIEWER'S Group Map rows. Either side can be the
 * reason the match fails, and they fail in completely different places:
 *
 *   - no groups read             -> the viewer's own membership could not be listed
 *   - no Group Map rows          -> the list could not be read, or holds no row for their groups
 *   - rows but no DEL / SHARE    -> their groups are mapped, but not with the roles that grant it
 *   - roles but no GUID overlap  -> both sides are populated and simply name different terms
 *   - document has no tier GUIDs -> the file was filed without its `<Base>Tid` values
 *
 * The verdict at the bottom names which one it is.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  /* Leave blank to auto-detect. Set it if the auto-detected site below is wrong, e.g.
     "https://sdguthrie.sharepoint.com/sites/CRS" */
  const SITE_OVERRIDE = "";

  /* Optional but STRONGLY recommended: the exact file name of one APPROVED document whose Delete
     button is wrongly asking for a reason, e.g. "s-ds-sd-15092026.pdf". Without it this checks only
     the viewer's own side and cannot compare the two. */
  const FILE_NAME = "";

  /* ── Resolve the site ──────────────────────────────────────────────────────
     Three routes, in order. `_spPageContextInfo` is missing on some modern pages (hit live on this
     project once already), and parsing the page path breaks the moment the script is run from a
     library or a settings page — which is how an earlier script in this repo ended up fetching the
     page's own HTML and failing with "Unexpected token '<'". */
  const resolveWeb = () => {
    if (SITE_OVERRIDE) return SITE_OVERRIDE.replace(/\/+$/, "");
    try {
      if (
        typeof _spPageContextInfo !== "undefined" &&
        _spPageContextInfo.webAbsoluteUrl
      ) {
        return _spPageContextInfo.webAbsoluteUrl.replace(/\/+$/, "");
      }
    } catch (e) {
      /* not defined on this page — fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
    return (location.origin + (m ? m[1] : "")).replace(/\/+$/, "");
  };
  const web = resolveWeb();
  console.log("Site:", web);

  const get = async (url) => {
    const r = await fetch(url, {
      headers: { Accept: "application/json;odata=nometadata" },
    });
    const type = r.headers.get("content-type") || "";
    if (type.indexOf("json") === -1) {
      /* A 200 carrying HTML means the ADDRESS is wrong, not the query — set SITE_OVERRIDE. */
      throw new Error(
        `Expected JSON and got "${type}" from ${url}\n` +
          `The site resolved to ${web}. If that is wrong, set SITE_OVERRIDE at the top and re-run.`,
      );
    }
    if (!r.ok) throw new Error(`HTTP ${r.status} from ${url}`);
    return r.json();
  };
  /* Soft form: returns undefined instead of throwing, so one unreadable thing does not end the run
     and hide every other answer. */
  const tryGet = async (url) => {
    try {
      return await get(url);
    } catch (e) {
      console.warn("  (could not read)", e.message);
      return undefined;
    }
  };

  const norm = (v) => (v ?? "").replace(/[{}]/g, "").trim().toLowerCase();

  /* Mirrors the app's `normalizeRoleValue`. A Group Map Role is often the LONG form, and a raw
     compare silently skips such a row — which is one of the candidate causes being tested here. */
  const SHORT = ["MEMBER","UPL","APR","DEL","DELS","SEGVIEW","UPLHC","APRHC","DELSHC","DELHC","SHAREHC","MEMBERHC","DEPTVIEW","SHARE","GLOBAL","ENTRY"];
  const ALIASES = {
    UPLOADER: "UPL",
    APPROVER: "APR",
    DELETER_DOCUMENTS: "DEL",
    DELETER_STAGING: "DELS",
    MEMBER: "MEMBER",
    VIEWER: "MEMBER",
    SEGMENTVIEW: "SEGVIEW",
    DEPARTMENTVIEW: "DEPTVIEW",
    VIEWER_HC: "MEMBERHC",
  };
  const role = (raw) => {
    const v = (raw ?? "").trim().toUpperCase().replace(/[\s-]+/g, "_");
    if (!v) return "";
    if (SHORT.indexOf(v) !== -1) return v;
    return ALIASES[v] ?? v;
  };

  /* ── 1. Who am I, and which groups am I in? ───────────────────────────────── */
  console.log("\n=== 1. This viewer ===");
  const me = await tryGet(
    `${web}/_api/web/currentuser?$select=Id,Title,Email,LoginName,IsSiteAdmin`,
  );
  if (me) {
    console.log(`  ${me.Title} · ${me.Email || "(no email)"}`);
    console.log(`  IsSiteAdmin: ${me.IsSiteAdmin}`);
    if (me.IsSiteAdmin) {
      console.log(
        "  NOTE: a site collection administrator BYPASSES this whole check in the app, so the " +
          "Delete button would work for them regardless of everything below. Re-run as the approver.",
      );
    }
  }
  const groupsRes = await tryGet(
    `${web}/_api/web/currentuser/groups?$select=Id,Title&$top=500`,
  );
  const myGroups = (groupsRes && groupsRes.value) || [];
  console.log(`  In ${myGroups.length} group(s):`);
  for (const g of myGroups) console.log(`    ${g.Id}  ${g.Title}`);
  const myIds = myGroups.map((g) => g.Id);

  /* ── 2. The Group Map, resolved the way the app resolves it ───────────────── */
  console.log("\n=== 2. Group Map rows for those groups ===");
  let mapTitle;
  for (const candidate of ["CRS Group Map", "DMS Group Map", "Group Map"]) {
    const probe = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(candidate)}')?$select=Title,ItemCount`,
    );
    if (probe && probe.Title) {
      mapTitle = probe.Title;
      console.log(`  List: "${mapTitle}" (${probe.ItemCount} items)`);
      break;
    }
  }
  if (!mapTitle) {
    console.log("  ✗ NO GROUP MAP LIST COULD BE READ under any known name, as this account.");
    console.log("    That alone explains it: with no rows, the viewer holds no recorded rights and");
    console.log("    every action falls back to raising a request.");
    return;
  }

  const rowsRes = await tryGet(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(mapTitle)}')/items` +
      `?$select=GroupId,GroupName,Role,UnitTermGuid&$top=5000`,
  );
  if (!rowsRes) {
    console.log("  ✗ The list exists but its ITEMS could not be read as this account.");
    console.log("    Same consequence: no recorded rights, so every action becomes a request.");
    return;
  }
  const all = rowsRes.value || [];
  const mine = all.filter((r) => myIds.indexOf(r.GroupId) !== -1);
  console.log(`  ${all.length} row(s) total, ${mine.length} for this viewer's groups.`);

  const directDelete = [];
  const directDeleteStaging = [];
  const directShare = [];
  const add = (list, guid) => {
    if (guid && list.indexOf(guid) === -1) list.push(guid);
  };
  for (const r of mine) {
    const code = role(r.Role);
    const guid = (r.UnitTermGuid || "").trim();
    console.log(
      `    ${r.GroupName || r.GroupId}  Role="${r.Role}" -> ${code || "(unrecognised)"}  tier=${guid || "(blank)"}`,
    );
    if (!guid) continue;
    if (code === "DEL" || code === "DELHC") add(directDelete, guid);
    if (code === "DELS" || code === "DELSHC") add(directDeleteStaging, guid);
    if (code === "SHARE" || code === "SHAREHC") add(directShare, guid);
  }
  console.log(
    `\n  Delete APPROVED documents, at tier(s): ${directDelete.length ? directDelete.join(", ") : "(none)"}`,
  );
  console.log(
    `  Delete PENDING files,      at tier(s): ${directDeleteStaging.length ? directDeleteStaging.join(", ") : "(none)"}`,
  );
  console.log(
    `  Share APPROVED documents,  at tier(s): ${directShare.length ? directShare.join(", ") : "(none)"}`,
  );

  /* ── 3. The document's own tier GUIDs ─────────────────────────────────────── */
  if (!FILE_NAME) {
    console.log("\n=== 3. Document ===");
    console.log("  Skipped — set FILE_NAME at the top to compare both sides. Without it this can");
    console.log("  only show what the viewer holds, not whether it matches the file in question.");
    return;
  }

  console.log(`\n=== 3. Document "${FILE_NAME}" ===`);
  /* The approved side only: this is about an APPROVED document. Titles are probed because this
     client renames libraries routinely, so a hardcoded title is the least reliable thing here. */
  const libCandidates = [
    "Restricted & Confidential Document",
    "Documents",
    "Highly Confidential Document",
    "HC Documents",
    "HC Document",
  ];
  let found;
  for (const title of libCandidates) {
    const hit = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/items` +
        `?$select=Id,FileLeafRef&$filter=FileLeafRef eq '${FILE_NAME.replace(/'/g, "''")}'&$top=5`,
    );
    if (hit && hit.value && hit.value.length > 0) {
      found = { title, id: hit.value[0].Id };
      console.log(`  Found in "${title}", item ${found.id}.`);
      break;
    }
  }
  if (!found) {
    console.log("  ✗ Not found in any approved-side library under that exact name.");
    console.log("    Check the spelling, or that it really is Approved rather than still pending.");
    return;
  }

  const ft = await tryGet(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(found.title)}')/items(${found.id})/FieldValuesAsText`,
  );
  if (!ft) {
    console.log("  ✗ Its field values could not be read as this account.");
    return;
  }
  /* The app derives a tier from any field that has a `<Base>Tid` twin carrying a value — exactly
     what is reproduced here, including the double-encoded `_x005f_` spelling OData returns. */
  const plain = {};
  for (const k of Object.keys(ft)) plain[k.replace(/_x005f_/g, "_")] = ft[k];
  const chain = [];
  for (const k of Object.keys(plain)) {
    if (!/Tid$/.test(k)) continue;
    const base = k.slice(0, -3);
    const guid = (plain[k] || "").trim();
    if (base === "BusinessSegment") {
      console.log(
        `    (segment) ${base} = "${plain[base] || ""}" tid=${guid || "(blank)"}  — excluded from matching by design`,
      );
      continue;
    }
    console.log(`    ${base} = "${plain[base] || ""}"  tid=${guid || "(blank)"}`);
    if (guid) chain.push(guid);
  }
  if (chain.length === 0) {
    console.log("  ✗ This document carries NO tier term GUIDs at all.");
    console.log("    Nothing can match, so every viewer falls back to raising a request. The file");
    console.log("    was filed without its tier `Tid` values — a metadata problem on the document,");
    console.log("    not a permissions one.");
    return;
  }

  /* ── 4. Verdict ───────────────────────────────────────────────────────────── */
  console.log("\n=== 4. Verdict ===");
  const overlaps = (held) =>
    chain.some((c) => held.some((h) => norm(h) === norm(c)));
  const delOk = overlaps(directDelete);
  const shareOk = overlaps(directShare);
  console.log(`  Delete this approved document directly?  ${delOk ? "YES" : "NO"}`);
  console.log(`  Share  this approved document directly?  ${shareOk ? "YES" : "NO"}`);

  /* ⚠ THE TWO WAYS OF ANSWERING "NO" NEED DIFFERENT FIXES, and saying only "no overlap" names the
     wrong one half the time — caught by the synthetic run before this was ever used on a site.
     EMPTY means the viewer is not mapped with the role at all (fix the Group Map / group
     membership); POPULATED-BUT-DIFFERENT means they are mapped, just at another term (compare the
     GUIDs). */
  const explain = (label, held, ok) => {
    if (ok) return;
    if (held.length === 0) {
      console.log(
        `\n  ${label}: this viewer's groups carry NO row granting it, at any tier. They are in the` +
          ` groups listed in section 1, but none of those groups is mapped with the role that grants` +
          ` this — so the app can only offer a request. Fix it on the Group Map, not on the document.`,
      );
      return;
    }
    console.log(
      `\n  ${label}: the viewer IS granted it, but at tier(s) ${held.join(", ")} — and this document` +
        ` is filed under ${chain.join(", ")}. Both sides are populated and simply name different` +
        ` terms: the group is mapped at one tier and the file sits under another.`,
    );
  };
  explain("Delete", directDelete, delOk);
  explain("Share", directShare, shareOk);

  if (delOk && shareOk) {
    console.log("\n  Both sides DO overlap, so the app should be offering the direct buttons here.");
    console.log("  If it is not, the problem is in the page rather than the data — say so and");
    console.log("  include this whole output.");
  }
})();

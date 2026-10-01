/*
 * Why does an APPROVER (or a system admin) still get the reason-required request dialog?
 * ----------------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, deleted or granted.
 *
 * Run it in the browser console, ON THE CRS SITE, SIGNED IN AS THE ACCOUNT THAT IS SEEING THE
 * PROBLEM.
 *
 * ⚠⚠ v2 (2026-09-22) — v1 checked ONLY the Group Map route and `IsSiteAdmin`. It missed TWO of the
 * app's three eligibility routes, so a clean "NO" from v1 could not actually explain why the app
 * offered a request — the account could still be an admin via the site's OWNERS GROUP (which
 * `IsSiteAdmin` alone does not catch), or the live folder ACL could grant delete/share directly even
 * where the Group Map says nothing. `MySubmissions.tsx`'s own rule is explicit: "EITHER saying yes is
 * enough; NEITHER is a veto" — so all three routes must be checked before concluding anything is
 * actually wrong. This version checks all three, in the SAME order and with the SAME bit arithmetic
 * the app itself uses (`isSystemAdmin` in spGroups.ts, `probeFileRights` in dmsFolderMap.ts).
 *
 * Three independent routes to "yes", checked in the app's own order:
 *   1. System admin — `IsSiteAdmin` OR membership of the site's own Owners group.
 *   2. The live folder ACL, asked directly (`EffectiveBasePermissions` on the file itself) — this is
 *      the route that "cannot be wrong the way the Group Map can", per the app's own comment.
 *   3. The Group Map — the viewer's own groups matched against the document's tier GUIDs.
 *
 * If ALL THREE say no, the request dialog is the CORRECT, intended behaviour for this account on
 * this document — not a bug. The verdict at the bottom says which of the three (if any) actually
 * grants it, and if none do, names exactly what would need to change.
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
  }
  /* ⚠ `isSystemAdmin` (spGroups.ts) checks TWO things, not one — `IsSiteAdmin` alone is only half of
     it. The app's other half is membership of the site's own OWNERS group, which is what Group
     Management actually puts an "administrator" into (client, 2026-08-27: "we did not include a
     group as the System Admin group? they should have all the power" — the answer was that Owners
     always WAS that group). v1 of this script missed this entirely, which is exactly the gap that
     made a real admin account look like it had no rights at all. */
  let ownersId;
  let inOwners = false;
  if (me && typeof me.Id === "number") {
    const ownersGroup = await tryGet(
      `${web}/_api/web/AssociatedOwnerGroup?$select=Id,Title`,
    );
    if (ownersGroup) {
      ownersId = ownersGroup.Id;
      const ownRes = await tryGet(
        `${web}/_api/web/AssociatedOwnerGroup/Users?$filter=Id eq ${me.Id}&$select=Id`,
      );
      inOwners = !!(ownRes && (ownRes.value || []).length > 0);
      console.log(
        `  In the Owners group ("${ownersGroup.Title}")?  ${inOwners ? "YES" : "no"}`,
      );
    }
  }
  const isSystemAdmin = !!(me && me.IsSiteAdmin) || inOwners;
  if (isSystemAdmin) {
    console.log(
      "  NOTE: this account IS a system administrator (site admin or Owners-group member).",
    );
    console.log(
      "  The app should offer the DIRECT Delete/Share buttons regardless of everything below.",
    );
    console.log(
      "  If it is still asking for a reason, the problem is in the PAGE, not the data — say so and",
    );
    console.log("  include this whole output.");
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

  /* ⚠⚠ v3 — DEL/DELHC/DELS/DELSHC WERE REMOVED FROM EVERY PERSONA ON 2026-09-17 (the delete-by-proxy
     redesign: "they can delete folders which is dangerous"). Confirmed by reading `groupMapModel.ts`
     directly — `hou`, `hou_hc` AND `hod` all lost these roles the same day. So `directDelete`/
     `directDeleteStaging` below are now STRUCTURALLY EMPTY FOR EVERY PERSONA ON THE SITE — that is
     expected, not a bug, and is NOT what decides whether a genuine approver gets treated correctly.

     The route that actually matters now is `decidesDeletion` (APR/APRHC, scoped to the viewer's own
     groups) — `MySubmissions.tsx`'s row-level Delete handler uses this EXACT match to decide between
     three outcomes: instant self-delete (`canDelete`, still checks the now-dead DEL route plus the
     ACL probe plus systemAdmin), a small inline notice ("You decide deletion requests for this unit
     in Approval & Request, so none is offered here" — for someone who IS the approver), or the full
     "Request deletion" modal (for someone who is neither). A genuine approver landing in the THIRD
     bucket instead of the second is the exact symptom being chased here. */
  const directDelete = [];
  const directDeleteStaging = [];
  const directShare = [];
  const decidesDeletion = [];
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
    if (code === "APR" || code === "APRHC") add(decidesDeletion, guid);
  }
  console.log(
    `\n  Delete APPROVED documents, at tier(s): ${directDelete.length ? directDelete.join(", ") : "(none — expected, see note above)"}`,
  );
  console.log(
    `  Delete PENDING files,      at tier(s): ${directDeleteStaging.length ? directDeleteStaging.join(", ") : "(none — expected, see note above)"}`,
  );
  console.log(
    `  Share APPROVED documents,  at tier(s): ${directShare.length ? directShare.join(", ") : "(none)"}`,
  );
  console.log(
    `  DECIDES deletion for (APR/APRHC), at tier(s): ${decidesDeletion.length ? decidesDeletion.join(", ") : "(none)"}`,
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
        `?$select=Id,FileLeafRef,File/UniqueId&$expand=File` +
        `&$filter=FileLeafRef eq '${FILE_NAME.replace(/'/g, "''")}'&$top=5`,
    );
    if (hit && hit.value && hit.value.length > 0) {
      const row = hit.value[0];
      found = { title, id: row.Id, uniqueId: (row.File && row.File.UniqueId) || "" };
      console.log(`  Found in "${title}", item ${found.id}.`);
      break;
    }
  }
  if (!found) {
    console.log("  ✗ Not found in any approved-side library under that exact name.");
    console.log("    Check the spelling, or that it really is Approved rather than still pending.");
    return;
  }

  /* ── 3b. The live folder ACL, asked directly — this is the route `MySubmissions.tsx` trusts most,
     because it "cannot be wrong the way the Group Map can": it answers about THIS viewer on THIS
     exact item, with no reconstruction from mapping rows in between. Same endpoint, same two bits,
     same arithmetic as `probeFileRights` in dmsFolderMap.ts — never `&`, because Full Control returns
     `Low = "4294967295"`, which JS bitwise coercion reads as a signed 32-bit int and gets wrong. */
  let aclRemove = "unknown";
  let aclShare = "unknown";
  if (found.uniqueId) {
    const perm = await tryGet(
      `${web}/_api/web/GetFileById(guid'${encodeURIComponent(found.uniqueId)}')` +
        `/ListItemAllFields/EffectiveBasePermissions`,
    );
    if (perm && perm.Low !== undefined) {
      const low = Number(perm.Low);
      const hasBit = (bitIndex) =>
        isFinite(low) && low >= 0 && Math.floor(low / Math.pow(2, bitIndex)) % 2 === 1;
      aclRemove = hasBit(3) ? "granted" : "denied"; // deleteListItems = kind 4, bit index 3
      aclShare = hasBit(25) ? "granted" : "denied"; // managePermissions = kind 26, bit index 25
    }
  }
  console.log(`\n  Live folder ACL on this exact file (EffectiveBasePermissions):`);
  console.log(`    Delete this file directly?  ${aclRemove}`);
  console.log(`    Share  this file directly?  ${aclShare}`);
  if (aclRemove === "unknown") {
    console.log(
      "    (unreadable, or the item's UniqueId could not be resolved — this route is inconclusive," +
        " not a NO)",
    );
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

  /* ── 4. Verdict — ALL THREE ROUTES COMBINED, exactly as `canDeleteSelf`/`canShareSelf` compute
     it: `systemAdmin || rights?.remove === "granted" || canActDirectly(chain, policy.directDelete)`.
     A clean "NO" here requires all three to independently say no — one YES anywhere is enough. */
  console.log("\n=== 4. Verdict ===");
  const overlaps = (held) =>
    chain.some((c) => held.some((h) => norm(h) === norm(c)));
  const groupMapDelOk = overlaps(directDelete);
  const groupMapShareOk = overlaps(directShare);
  const delOk = isSystemAdmin || aclRemove === "granted" || groupMapDelOk;
  const shareOk = isSystemAdmin || aclShare === "granted" || groupMapShareOk;
  console.log(`  Delete this approved document directly?  ${delOk ? "YES" : "NO"}`);
  console.log(`    admin=${isSystemAdmin}  ACL=${aclRemove}  GroupMap(DEL)=${groupMapDelOk}`);
  console.log(`  Share  this approved document directly?  ${shareOk ? "YES" : "NO"}`);
  console.log(`    admin=${isSystemAdmin}  ACL=${aclShare}  GroupMap(SHARE)=${groupMapShareOk}`);

  /* ⚠⚠ THE THIRD, SEPARATE QUESTION — and the one that actually decides what the ROW-level Delete
     button in the My Submissions TABLE shows, per `MySubmissions.tsx`'s own three-way branch:
       canDelete (above)         -> instant self-delete
       else decidesDeletionOk    -> small inline notice: "You decide deletion requests for this
                                     unit in Approval & Request, so none is offered here."
       else                      -> the FULL "Request deletion" modal, asking someone else
     If `delOk` is NO but this is YES, the app should have shown the inline notice, not the modal —
     that combination is the exact symptom this script was extended to catch. */
  const decidesDeletionOk = overlaps(decidesDeletion);
  console.log(
    `\n  Does this viewer DECIDE deletion requests for this unit (APR/APRHC)?  ${decidesDeletionOk ? "YES" : "NO"}`,
  );
  console.log(`    GroupMap(APR/APRHC)=${decidesDeletionOk}`);
  if (!delOk && decidesDeletionOk) {
    console.log(
      "\n  ⚠⚠ delOk=NO but decidesDeletionOk=YES: the app should have shown the SMALL INLINE NOTICE",
    );
    console.log(
      '  ("You decide deletion requests for this unit in Approval & Request, so none is offered',
    );
    console.log(
      '  here.") — NOT the full "Request deletion" modal with a Reason box. If the modal appeared',
    );
    console.log(
      "  instead, the mismatch is in the PAGE (most likely a stale/not-yet-loaded `policy` at the",
    );
    console.log(
      "  moment of the click, or the tier chain above disagreeing with what Group Management shows),",
    );
    console.log("  not a missing permission. Say so and include this whole output.");
  }

  if (delOk && shareOk) {
    console.log(
      "\n  At least one route says YES for both. The app SHOULD be offering the direct buttons —",
    );
    console.log(
      "  if it is still showing the request dialog, the problem is in the PAGE (a timing/caching",
    );
    console.log(
      "  issue, or the probe simply had not landed yet when the button was clicked), not the data.",
    );
    console.log("  Say so and include this whole output.");
    return;
  }

  /* Only reached when the ACL and admin routes BOTH said no for the relevant action — the Group Map
     detail below explains the one route that is genuinely about DATA, not about the page. THE TWO
     WAYS OF ANSWERING "NO" HERE NEED DIFFERENT FIXES, and saying only "no overlap" names the wrong
     one half the time. EMPTY means the viewer is not mapped with the role at all (fix the Group Map
     / group membership); POPULATED-BUT-DIFFERENT means they are mapped, just at another term
     (compare the GUIDs). */
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
  explain("Delete", directDelete, groupMapDelOk);
  explain("Share", directShare, groupMapShareOk);
})();

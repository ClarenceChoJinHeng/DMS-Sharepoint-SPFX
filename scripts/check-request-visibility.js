/*
 * Why can this account see (or not see) a particular request on the "File Permission" /
 * "Document Deletion & Sharing Approval" page (Requests.tsx)?
 * -----------------------------------------------------------------------------------------
 * READ ONLY. Every request is a GET. Nothing is written, deleted or granted.
 *
 * Run it in the browser console, ON THE CRS SITE, SIGNED IN AS THE ACCOUNT THAT SEES THE PAGE
 * DIFFERENTLY (e.g. crs@sdguthrie.com), then run it AGAIN signed in as the account that sees more
 * (e.g. the system administrator) and compare the two printouts side by side.
 *
 * Built 2026-09-23 after a live report: "crs@sdguthrie.com is an approver but cannot see the file
 * in file permission, but system admin can see." A system admin is EXPECTED to see everything
 * (`ViewerScope.systemAdmin`, by design, since 2026-08-27) — that alone does not mean anything is
 * wrong. What actually decides whether an ORDINARY approver sees one specific request is:
 *   1. Did they raise it themselves? (`isVisibleTo` always shows a viewer their own requests.)
 *   2. Do they hold APR/APRHC for that request's UNIT (matched on `UnitTermGuid`, never the label)?
 *   3. Only for an APPROVED-stage SHARE request: do they hold DEPTVIEW for that unit (a Head of
 *      Department decides Share only, never Deletion, since 2026-09-17)?
 * This script prints exactly those three facts for THIS signed-in account against every request
 * currently on the list, so "should this be visible" can be read off directly instead of guessed.
 *
 * ⚠ THE HEADER COUNT ("Requests N") AND THE ACCORDION BADGES ("N pending") ARE DELIBERATELY
 * DIFFERENT NUMBERS, AND THAT ON ITS OWN IS NOT A BUG. The header counts only requests this viewer
 * can actually DECIDE (`queueFor` → `canDecide`, which does NOT include "you raised this one
 * yourself"); each accordion's "N pending" badge counts everything VISIBLE to this viewer of that
 * type (`ofType`, which DOES include your own raised requests). A viewer who raised requests they
 * cannot approve will always see the accordion badges add up to MORE than the header number. This
 * script's own tallies at the bottom mirror that split so the two numbers can be reasoned about
 * rather than treated as inconsistent.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  /* Leave blank to auto-detect. Set it if the auto-detected site below is wrong, e.g.
     "https://sdguthrie.sharepoint.com/sites/CRS" */
  const SITE_OVERRIDE = "";

  /* Optional: exact file names to call out specifically in a dedicated section at the end, e.g.
     ["test5-test5-test5-19092026.xlsx", "test3-test3-test3-21092026.xlsx", "Test NBPOL Doc-14092026.pdf"].
     Leave empty to just see every request. */
  const WATCH_ITEM_NAMES = [];

  /* How many CRS Requests rows to read. 500 covers this site comfortably today; raise it if the
     list has grown past that and rows are missing from the printout below. */
  const REQUEST_TOP = 500;

  /* ── Resolve the site ──────────────────────────────────────────────────────
     `_spPageContextInfo` is missing on some modern pages, and parsing the page path breaks the
     moment this is run from a library or a settings page — the exact trap that once made a
     different script here fetch the page's own HTML and fail with "Unexpected token '<'". */
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

  /* Mirrors the app's `normalizeRoleValue` — a Group Map Role is often the LONG form. */
  const SHORT = ["MEMBER","UPL","APR","DEL","DELS","SEGVIEW","UPLHC","APRHC","DELSHC","DELHC","SHAREHC","MEMBERHC","DEPTVIEW","SHARE","GLOBAL","ENTRY"];
  const ALIASES = {
    UPLOADER: "UPL",
    APPROVER: "APR",
    DELETER_DOCUMENTS: "DEL",
    DELETER_STAGING: "DELS",
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
  const norm = (v) => (v ?? "").replace(/[{}]/g, "").trim().toLowerCase();

  /* ── 1. Who am I? ──────────────────────────────────────────────────────── */
  console.log("\n=== 1. This viewer ===");
  const me = await tryGet(
    `${web}/_api/web/currentuser?$select=Id,Title,Email,LoginName,IsSiteAdmin`,
  );
  if (!me) {
    console.log("  ✗ Could not read currentuser — stopping.");
    return;
  }
  const myEmail = norm(me.Email);
  console.log(`  ${me.Title} · ${me.Email || "(no email)"}`);
  console.log(`  IsSiteAdmin: ${me.IsSiteAdmin}`);

  let inOwners = false;
  const ownersGroup = await tryGet(
    `${web}/_api/web/AssociatedOwnerGroup?$select=Id,Title`,
  );
  if (ownersGroup) {
    const ownRes = await tryGet(
      `${web}/_api/web/AssociatedOwnerGroup/Users?$filter=Id eq ${me.Id}&$select=Id`,
    );
    inOwners = !!(ownRes && (ownRes.value || []).length > 0);
    console.log(
      `  In the Owners group ("${ownersGroup.Title}")?  ${inOwners ? "YES" : "no"}`,
    );
  }
  const systemAdmin = !!me.IsSiteAdmin || inOwners;
  if (systemAdmin) {
    console.log(
      "  This account IS a system administrator — every request should be VISIBLE to it, and",
    );
    console.log(
      "  every PENDING request DECIDABLE, regardless of everything below. That is by design.",
    );
  }

  const groupsRes = await tryGet(
    `${web}/_api/web/currentuser/groups?$select=Id,Title&$top=500`,
  );
  const myGroups = (groupsRes && groupsRes.value) || [];
  console.log(`  In ${myGroups.length} group(s):`);
  for (const g of myGroups) console.log(`    ${g.Id}  ${g.Title}`);
  const myGroupIds = myGroups.map((g) => g.Id);

  /* ── 2. Group Map rows → aprUnits / hodUnits, exactly as Requests.tsx builds them ────────── */
  console.log("\n=== 2. Group Map — which units this account approves ===");
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
  const aprUnits = [];
  const hodUnits = [];
  if (!mapTitle) {
    console.log("  ✗ NO GROUP MAP LIST COULD BE READ under any known name, as this account.");
    console.log("    With no rows, this account holds no recorded approving units at all.");
  } else {
    const rowsRes = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(mapTitle)}')/items` +
        `?$select=GroupId,GroupName,Role,UnitTermGuid&$top=5000`,
    );
    if (!rowsRes) {
      console.log("  ✗ The list exists but its ITEMS could not be read as this account.");
    } else {
      const all = rowsRes.value || [];
      const mine = all.filter((r) => myGroupIds.indexOf(r.GroupId) !== -1);
      console.log(`  ${all.length} row(s) total, ${mine.length} for this viewer's groups.`);
      for (const r of mine) {
        const code = role(r.Role);
        const guid = norm(r.UnitTermGuid);
        console.log(
          `    ${r.GroupName || r.GroupId}  Role="${r.Role}" -> ${code || "(unrecognised)"}  unit=${guid || "(blank)"}`,
        );
        if (!guid) continue;
        if ((code === "APR" || code === "APRHC") && aprUnits.indexOf(guid) === -1)
          aprUnits.push(guid);
        if (code === "DEPTVIEW" && hodUnits.indexOf(guid) === -1) hodUnits.push(guid);
      }
    }
  }
  console.log(`\n  Approves (APR/APRHC), unit(s): ${aprUnits.length ? aprUnits.join(", ") : "(none)"}`);
  console.log(`  Head of Department (DEPTVIEW), unit(s): ${hodUnits.length ? hodUnits.join(", ") : "(none)"}`);

  /* ── 3. Every request on the list, judged the way `isVisibleTo`/`canDecide` judge it ──────── */
  console.log("\n=== 3. CRS Requests — one line per row ===");
  let reqTitle;
  for (const candidate of ["CRS Requests", "DMS Requests"]) {
    const probe = await tryGet(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(candidate)}')?$select=Title,ItemCount`,
    );
    if (probe && probe.Title) {
      reqTitle = probe.Title;
      console.log(`  List: "${reqTitle}" (${probe.ItemCount} items)`);
      break;
    }
  }
  if (!reqTitle) {
    console.log("  ✗ NO REQUESTS LIST COULD BE READ under any known name, as this account.");
    console.log("    That alone would explain seeing nothing at all on the page.");
    return;
  }

  const reqRes = await tryGet(
    `${web}/_api/web/lists/getbytitle('${encodeURIComponent(reqTitle)}')/items` +
      `?$select=Id,RequestType,Status,ItemName,Segment,Unit,UnitTermGuid,RequestedBy,RequestedAt,Stage` +
      `&$orderby=RequestedAt desc&$top=${REQUEST_TOP}`,
  );
  if (!reqRes) {
    console.log("  ✗ The list exists but its ITEMS could not be read as this account.");
    console.log("    Same consequence as above: this would explain seeing nothing.");
    return;
  }
  const rows = reqRes.value || [];
  console.log(`  ${rows.length} row(s) read.\n`);

  const visibleRows = [];
  const notVisibleRows = [];
  const decidableRows = [];

  for (const r of rows) {
    const unitKey = norm(r.UnitTermGuid) || norm(r.Unit);
    const isMine = myEmail.length > 0 && norm(r.RequestedBy) === myEmail;
    const aprMatch = unitKey.length > 0 && aprUnits.indexOf(unitKey) !== -1;
    // Stage defaults to "approved" when the column is absent/blank — same rule as `stageOf` in
    // shared/requests.ts, for every row written before the Stage column existed.
    const stage = (r.Stage || "approved").trim().toLowerCase();
    const hodMatch =
      unitKey.length > 0 &&
      hodUnits.indexOf(unitKey) !== -1 &&
      stage === "approved" &&
      r.RequestType !== "Deletion";

    let why;
    if (systemAdmin) why = "system administrator";
    else if (isMine) why = "you raised this request";
    else if (aprMatch) why = "you approve this unit (APR/APRHC)";
    else if (hodMatch) why = "you are Head of Department for this unit (Share, approved-stage)";

    const visible = why !== undefined;
    const decidable = visible && !isMine && r.Status === "Pending"; // canDecide = inScope only

    const line =
      `  [${visible ? "VISIBLE" : "hidden "}]` +
      `${decidable ? " [DECIDABLE]" : ""}` +
      ` #${r.Id} ${r.RequestType} · ${r.Status} · "${r.ItemName}"` +
      ` · unit=${unitKey || "(blank)"} · by ${r.RequestedBy}` +
      (visible ? ` — ${why}` : "");

    const watched =
      WATCH_ITEM_NAMES.length === 0 ||
      WATCH_ITEM_NAMES.some(
        (n) => (r.ItemName || "").toLowerCase() === n.toLowerCase(),
      );
    if (watched) console.log(line);

    (visible ? visibleRows : notVisibleRows).push(r);
    if (decidable) decidableRows.push(r);
  }

  if (WATCH_ITEM_NAMES.length > 0) {
    console.log("\n  (Only the file names listed in WATCH_ITEM_NAMES were printed above.");
    console.log("   Set WATCH_ITEM_NAMES = [] at the top to print every row instead.)");
  }

  console.log("\n=== 4. Summary — matches what the page itself would show this account ===");
  console.log(`  Visible in total:                ${visibleRows.length}`);
  console.log(`    of which Pending & decidable:  ${decidableRows.length}  ← this is "Requests N" at the top of the page`);
  console.log(`  Pending & visible but NOT decidable (you raised it yourself): ${
    visibleRows.filter((r) => r.Status === "Pending" && norm(r.RequestedBy) === myEmail && !systemAdmin).length
  }`);
  console.log(`  NOT visible at all:               ${notVisibleRows.length}`);
  if (notVisibleRows.length > 0 && !systemAdmin) {
    console.log("\n  Requests this account CANNOT see, and the unit each one needs:");
    for (const r of notVisibleRows) {
      const unitKey = norm(r.UnitTermGuid) || norm(r.Unit);
      console.log(
        `    #${r.Id} ${r.RequestType} · "${r.ItemName}" · needs unit ${unitKey || "(blank)"}` +
          ` — not in this account's APR/APRHC list above` +
          (r.RequestType === "Share" ? " (and not its DEPTVIEW list either)" : ""),
      );
    }
    console.log("\n  If any of these SHOULD be visible, the fix is a Group Map / group-membership");
    console.log("  question, not a code one: confirm the unit's UnitTermGuid above appears in");
    console.log("  this account's own APR/APRHC row set in section 2 — Group Management's own");
    console.log("  \"Quick Search\" for this email shows the same mapping and is the faster way to");
    console.log("  check it without this script.");
  }
})();

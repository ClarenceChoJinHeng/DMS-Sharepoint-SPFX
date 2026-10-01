/**
 * Diagnoses why a normal upload form self-approve did (or did not) fire for the signed-in user.
 *
 * Built 2026-09-10 because a system admin's own upload was reported as staying Pending. The
 * app's self-approve logic (`shared/selfApprove.ts` + `probeFolderApproveAccess` in
 * `shared/dmsFolderMap.ts`) has NO persona branch at all — it is the same three checks for
 * anyone who uploads:
 *
 *   1. `CRS Config` row `autoApproveOwnUpload` = "yes" (trimmed, case-insensitive). Anything
 *      else — absent, blank, "no" — is OFF by design (fails CLOSED, unlike most config here).
 *   2. A LIVE `EffectiveBasePermissions` read on the exact destination folder shows the
 *      signed-in user holds ApproveItems (bit index 4). Not a role/persona lookup.
 *   3. No unresolved name clash on the approved side for the file just uploaded (skips
 *      self-approve, leaves the item Pending for a human — this script cannot see that half,
 *      since it is a per-upload check against a specific filename).
 *
 * This script reproduces checks 1 and 2 exactly as the app performs them, PLUS reports whether
 * the signed-in account is a genuine Site Collection Administrator or a `CRS Owners` group
 * member (or neither) — because those two routes to "administrator" behave differently against
 * a folder's ACL: an SCA bypasses permission checks entirely; a plain Owners-group member only
 * has Full Control on a folder if Owners was actually re-granted there (reconciliation does
 * this on every folder whose inheritance it breaks — see FolderManager.tsx's `ownerGroupId`
 * restore calls — but ONLY once reconciliation has actually run since that folder existed).
 *
 * READ ONLY. Writes nothing, safe to run repeatedly, safe to run as any account.
 *
 * ── How to run ───────────────────────────────────────────────────────────────
 * 1. Set FOLDER_PATH below to the server-relative path of the folder the test file was
 *    uploaded into (the one the "Document folder information" card showed) — e.g.
 *    "/sites/CRS/ApprovalDocument/GHO/GCA/GCBC/2024/Tax Return".
 * 2. Sign in as the account you want to check (system admin, then an approver, etc.) and open
 *    any page of that site.
 * 3. Paste this whole file into the DevTools console.
 * 4. Repeat as each different account, and compare the ApproveItems line.
 */

(async () => {
  "use strict";

  // ⚠ FILL THIS IN — the folder the test upload actually landed in. Leave blank to skip the
  // per-folder ACL check and only see the config value + admin-status checks.
  const FOLDER_PATH = "";

  const SITE =
    (typeof _spPageContextInfo !== "undefined" &&
      _spPageContextInfo &&
      _spPageContextInfo.webAbsoluteUrl) ||
    location.origin + ((location.pathname.match(/^\/sites\/[^/]+/) || [""])[0]);

  const get = (url) =>
    fetch(url, {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
      },
    });

  const ok = (m) => console.log(`%c  PASS  %c${m}`, "color:#0f6c3f;font-weight:bold", "");
  const warn = (m) => console.log(`%c  ??    %c${m}`, "color:#8a4b00;font-weight:bold", "");
  const fail = (m) => console.log(`%c  FAIL  %c${m}`, "color:#a4262c;font-weight:bold", "");

  console.log(`%cSite: ${SITE}`, "font-weight:bold");

  /* ── Who is signed in ─────────────────────────────────────────────────────── */
  const meRes = await get(`${SITE}/_api/web/currentuser?$select=Title,Email,IsSiteAdmin,Id`);
  if (!meRes.ok) {
    fail(`could not read currentuser (HTTP ${meRes.status}) — cannot continue`);
    return;
  }
  const me = await meRes.json();
  console.log(`%c\nSigned in as: ${me.Title} <${me.Email || "no email"}>`, "font-weight:bold;font-size:13px");

  if (me.IsSiteAdmin) {
    ok("IsSiteAdmin = true — a genuine Site Collection Administrator. SCAs bypass SharePoint's " +
       "permission checks entirely, so EffectiveBasePermissions should read Full Control on " +
       "ANY folder regardless of its ACL.");
  } else {
    warn("IsSiteAdmin = false — not a Site Collection Administrator. Whether this account " +
         "reads as an administrator anywhere in the app depends on CRS Owners membership " +
         "(checked next), and folder-level access depends entirely on that folder's real ACL.");
  }

  const ownersRes = await get(`${SITE}/_api/web/associatedownergroup?$expand=Users&$select=Users/Email,Users/Id`);
  if (ownersRes.ok) {
    const owners = await ownersRes.json();
    const inOwners = (owners.Users || []).some((u) => u.Id === me.Id);
    if (inOwners) {
      ok("member of the site's Owners group (CRS Owners) — Full Control on the WEB, and " +
         "reconciliation explicitly re-grants this group Full Control on every folder whose " +
         "inheritance it breaks. If this folder was created BEFORE reconciliation last ran, " +
         "or reconciliation has never run since, Owners may not actually be on its ACL yet.");
    } else {
      warn("NOT a member of the site's Owners group. Whatever access this account has on the " +
           "destination folder comes from an ordinary persona grant (or nothing).");
    }
  } else {
    warn(`could not read the site's Owners group (HTTP ${ownersRes.status})`);
  }

  /* ── Check 1: the config row ──────────────────────────────────────────────── */
  console.log(`%c\nCheck 1 — CRS Config row "autoApproveOwnUpload"`, "font-weight:bold;font-size:13px");

  function reportConfigRow(rows) {
    if (rows.length === 0) {
      fail(
        'no "autoApproveOwnUpload" row exists — this fails CLOSED by design, so self-approve ' +
          "is OFF for EVERY account on this site, admin or approver. This is the most likely " +
          "cause if nobody's upload is self-approving. Add the row with SettingValue = \"yes\" " +
          "if the client wants this on.",
      );
      return;
    }
    const raw = rows[0].SettingValue;
    const on = (raw || "").trim().toLowerCase() === "yes";
    if (on) {
      ok(`SettingValue = "${raw}" — self-approve is ON site-wide. So a failure to self-approve ` +
         "is NOT the config; it is check 2 (the ACL) or check 3 (a same-named-file clash) for " +
         "THIS specific account and folder.");
    } else {
      fail(`SettingValue = "${raw}" — anything other than exactly "yes" is OFF. Self-approve ` +
           "will not fire for ANY account, admin or approver, until this is corrected.");
    }
  }

  const cfgRes = await get(
    `${SITE}/_api/web/lists/getbytitle('CRS Config')/items?$select=Title,SettingValue&$filter=Title eq 'autoApproveOwnUpload'&$top=1`,
  );
  if (!cfgRes.ok) {
    // The list may still be titled the legacy "DMS Config" on a site this has never been
    // renamed on — retry once before concluding the row is unreadable.
    const legacyRes = await get(
      `${SITE}/_api/web/lists/getbytitle('DMS Config')/items?$select=Title,SettingValue&$filter=Title eq 'autoApproveOwnUpload'&$top=1`,
    );
    if (!legacyRes.ok) {
      fail(`could not read CRS Config or DMS Config (HTTP ${cfgRes.status} / ${legacyRes.status})`);
    } else {
      reportConfigRow((await legacyRes.json()).value || []);
    }
  } else {
    reportConfigRow((await cfgRes.json()).value || []);
  }

  /* ── Check 2: ApproveItems on the actual destination folder ──────────────── */
  console.log(`%c\nCheck 2 — ApproveItems on the destination folder`, "font-weight:bold;font-size:13px");
  if (!FOLDER_PATH) {
    warn("FOLDER_PATH is blank — set it to the folder the test file landed in and re-run to " +
         "see this account's real ApproveItems answer for that exact folder.");
  } else {
    const idRes = await get(
      `${SITE}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeURIComponent(FOLDER_PATH)}'&$select=UniqueId,ServerRelativeUrl`,
    );
    if (!idRes.ok) {
      fail(`could not resolve "${FOLDER_PATH}" (HTTP ${idRes.status}) — check the path is exact, ` +
           "server-relative, and that this account can at least see the folder");
    } else {
      const folder = await idRes.json();
      console.log(`        folder: ${folder.ServerRelativeUrl}  (UniqueId ${folder.UniqueId})`);
      const permRes = await get(
        `${SITE}/_api/web/GetFolderById(guid'${encodeURIComponent(folder.UniqueId)}')` +
          `/ListItemAllFields/EffectiveBasePermissions`,
      );
      if (permRes.status === 401 || permRes.status === 403) {
        fail("EffectiveBasePermissions came back 401/403 — this exact probe in the app reads " +
             "this as DENIED, so self-approve is skipped and the item is left Pending.");
      } else if (permRes.status === 404) {
        fail("EffectiveBasePermissions came back 404 — the app reads this as the folder being " +
             "MISSING (or security-trimmed away from this account), self-approve is skipped.");
      } else if (!permRes.ok) {
        warn(`EffectiveBasePermissions came back HTTP ${permRes.status} — the app treats this ` +
             "as inconclusive and leaves the item Pending.");
      } else {
        const perms = await permRes.json();
        const low = Number(perms.Low);
        // ApproveItems is bit index 4 (viewListItems=1, addListItems=2, editListItems=3,
        // deleteListItems=4, approveItems=5) — arithmetic, never `&`: Full Control returns
        // Low = "4294967295", which JS bitwise coerces to a signed 32-bit int.
        const hasApprove = Math.floor(low / Math.pow(2, 4)) % 2 === 1;
        console.log(`        Low = ${perms.Low}`);
        if (hasApprove) {
          ok("ApproveItems IS granted on this folder for this account — check 2 passes. If " +
             "the file still stayed Pending, the cause is check 3 (a name clash on the " +
             "approved side skipped self-approve) or the self-approve MERGE itself failed — " +
             "look in the browser console from the upload for \"Self-approve MERGE failed\" " +
             "or \"Self-approve skipped\".");
        } else {
          fail("ApproveItems is NOT granted on this folder for this account. This is why the " +
               "upload stayed Pending — check 2 fails regardless of the config setting. If " +
               "this account is a CRS Owners member, reconciliation likely has not (re)run on " +
               "this folder since it was created; if it is neither an SCA nor an Owners " +
               "member, it simply holds no Approve role here (expected for a PIC).");
        }
      }
    }
  }

  console.log(
    "%c\nNot checked here: a same-named-file clash on the approved side (per-filename, not " +
      "checkable from a script), and the self-approve MERGE's own success — both only show up " +
      "as console.warn lines in the browser at the moment of upload.",
    "color:#666",
  );
})();

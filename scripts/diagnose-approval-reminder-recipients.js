/*
 * Why do the two approval-reminder flows email an admin/test account instead of the real
 * per-unit approvers?
 * -------------------------------------------------------------------------------------------
 * `CRS — Reminding Approver to Approve` and `CRS — HC approval reminder` (exported 2026-09-21)
 * both do the SAME four steps for every stale pending document: read its `UnitTid`, find the
 * `CRS Group Map` row(s) mapping `Role eq 'APR'` (or `'APRHC'`) to that exact unit term GUID,
 * list that group's MEMBERS, then email every member except the document's own author.
 *
 * READ ONLY. Every request is a GET — nothing is written, deleted or sent. Safe to re-run.
 *
 * This reproduces those four steps exactly, for every currently pending document on this site,
 * and prints who each flow would ACTUALLY email right now. Four distinct failure shapes are
 * told apart, because they need completely different fixes:
 *
 *   A. UnitTid is blank on the document           -> a metadata problem on the file itself
 *   B. No Group Map row maps that unit to APR(HC)  -> the unit's approver group was never mapped
 *   C. The mapped group has genuine members,
 *      but they ALL equal the document's author    -> the flow's own exclusion empties the list
 *      (self-loop: nobody but the uploader is in the group at all)
 *   D. The mapped group's members are a TEST/ADMIN account, not real approver staff
 *      -> the flow is working exactly as built; the GROUP itself needs real people added.
 *      This is the one CLAUDE.md already flagged once on this tenant (`crs@sdguthrie.com`
 *      sitting in `GHO_COSEC_GUTHRIE_APPROVER` etc. as leftover test setup) — this script
 *      checks it directly rather than assuming it is still true.
 *
 * ── How to run ────────────────────────────────────────────────────────────────────────────
 * Paste this whole file into the DevTools console on the CRS site, signed in as an account
 * that can read `CRS Group Map` and the approval libraries (an admin account is fine — this
 * is a read, not the live send, so nothing here is skipped for an admin the way the app's own
 * "does this viewer hold the role" checks sometimes are).
 */
(async () => {
  "use strict";

  /* ── Configure ──────────────────────────────────────────────────────────────────────── */

  // Leave blank to auto-detect. Set it if the auto-detected site below is wrong.
  const SITE_OVERRIDE = "";

  // Same list GUIDs the two flow exports use — resolved by GUID first (exact match to what the
  // flow itself reads), falling back to a title probe only if that GUID does not resolve here
  // (e.g. this script is being run against a different site than the one the flows target).
  const LIBRARIES = [
    {
      label: "Normal (Approval for Document)",
      guid: "eeb1bb19-ec53-4c41-8dc9-ecf255979c9b",
      titleCandidates: ["Approval for Document", "Approval Document"],
      approverRole: "APR",
      staleCutoffDays: 3,
    },
    {
      label: "HC (Approval for Highly Confidential Document)",
      guid: "d935aa6d-dc48-4832-ba7e-eca485ecdff3",
      titleCandidates: [
        "Approval for Highly Confidential Document",
        "HC Approval Document",
      ],
      approverRole: "APRHC",
      staleCutoffDays: 1,
    },
  ];

  /* ── Resolve the site — three routes, in order. `_spPageContextInfo` is missing on some
     modern pages, and parsing the page path breaks if this is run from a library/settings page. */
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
      /* fall through */
    }
    const m = location.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
    return (location.origin + (m ? m[1] : "")).replace(/\/+$/, "");
  };
  const web = resolveWeb();
  console.log("Site:", web);

  const bust = () => `_=${Date.now()}${Math.random().toString(36).slice(2)}`;
  const get = async (url) => {
    const sep = url.indexOf("?") > -1 ? "&" : "?";
    try {
      const r = await fetch(`${url}${sep}${bust()}`, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Cache-Control": "no-cache",
          Pragma: "no-cache",
        },
        cache: "no-store",
      });
      const type = r.headers.get("content-type") || "";
      if (type.indexOf("json") === -1) {
        return { ok: false, status: r.status, body: { __notJson: true } };
      }
      let body;
      try {
        body = await r.json();
      } catch (e) {
        body = { __unparseable: true };
      }
      return { ok: r.ok, status: r.status, body };
    } catch (e) {
      return { ok: false, status: 0, body: { __error: String((e && e.message) || e) } };
    }
  };

  const norm = (v) => (v ?? "").toString().trim().toLowerCase();

  console.log(
    "\n" +
      "NOTE ON WHAT THIS CANNOT SEE: if the flow's OWN run history shows every run FAILING\n" +
      "(rather than succeeding and mailing the wrong people), that means Power Automate's own\n" +
      "failure-alert email is what is landing in the admin's inbox — a completely different\n" +
      "cause from anything below. Check Run History in the flow first: if the 'admin' email is\n" +
      "literally titled something like 'A cloud flow that you own has failed' rather than\n" +
      "'REMINDER: Approval Required for ...', stop here and read the failed action's error\n" +
      "instead — this script only diagnoses a flow that is running and mis-resolving recipients.",
  );

  for (const lib of LIBRARIES) {
    console.log(`\n${"=".repeat(70)}\n${lib.label}\n${"=".repeat(70)}`);

    // Resolve the title for the exact GUID, the same way the flow's own `dataset`/`table` pair
    // does — and fall back to a title probe so this still runs on a site with different GUIDs.
    let listTitle;
    const byGuid = await get(
      `${web}/_api/web/lists(guid'${lib.guid}')?$select=Title`,
    );
    if (byGuid.ok) {
      listTitle = byGuid.body.Title;
      console.log(`Resolved by GUID: "${listTitle}"`);
    } else {
      for (const cand of lib.titleCandidates) {
        const probe = await get(
          `${web}/_api/web/lists/getbytitle('${encodeURIComponent(cand)}')?$select=Title`,
        );
        if (probe.ok) {
          listTitle = probe.body.Title;
          console.log(`GUID did not resolve here — matched by title instead: "${listTitle}"`);
          break;
        }
      }
    }
    if (!listTitle) {
      console.log("✗ Could not resolve this library at all on this site — skipping.");
      continue;
    }

    /* ── GetStale, exactly as the flow filters it ─────────────────────────────────────── */
    const cutoff = new Date(Date.now() - lib.staleCutoffDays * 24 * 60 * 60 * 1000).toISOString();
    const filter =
      `FSObjType eq 0 and OData__ModerationStatus eq 2 and Created lt '${cutoff}'`;
    const staleRead = await get(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
        `?$select=Id,FileLeafRef,Created,Unit,UnitTid,Author/Title,Author/EMail` +
        `&$expand=Author&$filter=${encodeURIComponent(filter)}&$top=200`,
    );
    if (!staleRead.ok) {
      console.log(
        `✗ Could not read this library's pending items (HTTP ${staleRead.status}). Nothing else` +
          " here can be checked without this.",
      );
      console.log("  Body:", staleRead.body);
      continue;
    }
    const stale = staleRead.body.value || [];
    console.log(
      `${stale.length} document(s) pending AND older than ${lib.staleCutoffDays} day(s) — this is` +
        ` exactly ${lib.label.indexOf("HC") === 0 ? "'CRS — HC approval reminder'" : "'CRS — Reminding Approver to Approve'"}'s own GetStale.`,
    );
    if (stale.length === 0) {
      // Also check ANY pending item, unfiltered by age, so a genuinely empty stale set is not
      // mistaken for "nothing pending at all".
      const anyPending = await get(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
          `?$select=Id&$filter=${encodeURIComponent("FSObjType eq 0 and OData__ModerationStatus eq 2")}&$top=1`,
      );
      const pendingCount = anyPending.ok ? (anyPending.body.value || []).length : undefined;
      console.log(
        pendingCount === 0
          ? "  Nothing is pending at all in this library right now — nothing for this flow to remind about."
          : "  Something IS pending, but nothing has crossed the staleness cutoff yet — the flow" +
              " has genuinely had nothing to send on its most recent runs. Not a bug.",
      );
      continue;
    }

    for (const item of stale) {
      const unitTid = (item.UnitTid || "").trim();
      const author = (item.Author && item.Author.EMail) || "";
      console.log(
        `\n  • ${item.FileLeafRef}  (Id ${item.Id}, uploaded ${item.Created})` +
          `\n      Unit="${item.Unit || ""}"  UnitTid="${unitTid || "(BLANK)"}"  Author=${author}`,
      );

      if (!unitTid) {
        console.log(
          "    ✗ A. UnitTid is blank on this document — GetApproverGroup's filter becomes" +
            " `UnitTermGuid eq ''`, which matches no real group. This is a metadata problem on the" +
            " FILE (it was filed with no Unit tier value), not on the flow.",
        );
        continue;
      }

      /* ── GetApproverGroup, exactly as the flow filters it ──────────────────────────── */
      const grpRead = await get(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent("CRS Group Map")}')/items` +
          `?$select=GroupId,GroupName&$filter=${encodeURIComponent(
            `Role eq '${lib.approverRole}' and UnitTermGuid eq '${unitTid}'`,
          )}&$top=5`,
      );
      if (!grpRead.ok) {
        console.log(
          `    ✗ Could not read CRS Group Map for this unit (HTTP ${grpRead.status}). Can't` +
            " check further for this document.",
        );
        continue;
      }
      const groups = grpRead.body.value || [];
      if (groups.length === 0) {
        console.log(
          `    ✗ B. No 'CRS Group Map' row maps Role='${lib.approverRole}' to this exact unit` +
            " term GUID. This unit's approver group was never mapped for this role — the flow" +
            " correctly finds nobody and its own `else` branch fires, silently, sending no email" +
            " to anyone (not even an admin) for THIS document.",
        );
        continue;
      }
      console.log(
        `    Matched group: "${groups[0].GroupName || groups[0].GroupId}" (GroupId ${groups[0].GroupId})` +
          (groups.length > 1
            ? ` — plus ${groups.length - 1} more; GetMembers only reads the FIRST one`
            : ""),
      );

      /* ── GetMembers + MembersWithEmail, exactly as the flow computes them ──────────── */
      const memRead = await get(
        `${web}/_api/web/sitegroups(${groups[0].GroupId})/users?$select=Title,Email`,
      );
      if (!memRead.ok) {
        console.log(
          `    ✗ Could not read that group's membership (HTTP ${memRead.status}).`,
        );
        continue;
      }
      const members = memRead.body.value || [];
      const recipients = members.filter(
        (m) => m.Email && norm(m.Email) !== norm(author),
      );
      console.log(
        `    Group has ${members.length} member(s): ${members
          .map((m) => `${m.Title} <${m.Email || "no email"}>`)
          .join(", ") || "(empty)"}`,
      );
      console.log(
        `    After excluding the document's own author (${author || "n/a"}), the flow would` +
          ` email: ${recipients.length ? recipients.map((r) => r.Email).join(", ") : "NOBODY"}`,
      );

      if (recipients.length === 0 && members.length > 0) {
        console.log(
          "    ✗ C. Every member of this group IS the document's own author — a self-loop." +
            " The flow's exclusion (correctly) empties the list, so nothing sends for this" +
            " document. If this account uploaded a document into a unit where they are ALSO" +
            " the only approver, that is expected, not a bug.",
        );
      } else if (recipients.length === 0 && members.length === 0) {
        console.log(
          "    ✗ The mapped group genuinely has NO members at all. Nobody has ever been added" +
            " to it — the group exists and is correctly mapped, but is empty.",
        );
      } else if (recipients.length > 0) {
        console.log(
          "    NOTE: if any of the addresses above is an admin/test/service account rather than" +
            " real approver staff, this is finding D — the flow is working exactly as designed," +
            " it is genuinely emailing every current member of the mapped approver group. The" +
            " fix is adding the real approver to this SharePoint group (and, if it is a leftover" +
            " test account like a service account added during earlier testing, removing IT from" +
            " the group so it stops receiving live reminders it should never see).",
        );
      }
    }
  }
})().catch((e) => console.error("Script error:", e));

/*
 * Which pending documents on the two approval libraries are due for a reminder right now, under
 * the per-file `NextReminderAt` cadence — and which aren't yet.
 * -------------------------------------------------------------------------------------------
 * Built 2026-09-24 while live-testing `CRS — Reminding Approver to Approve`'s rebuilt GetStale
 * filter, then extended the same day to also cover `CRS — HC approval reminder` once that flow
 * got the identical treatment (see
 * docs/superpowers/specs/2026-09-22-per-file-reminder-cadence-runbook.md). Reproduces the EXACT
 * filter each flow now uses — including the `BulkImport ne true` exclusion added after the first
 * live test on the normal flow accidentally re-triggered a separate bulk-auto-approve flow — so
 * you can compare what this script says against what each flow's own `GetStale` action actually
 * returned in its run history, and, after a test run, confirm which item(s) got their
 * `NextReminderAt` stamped forward by 3 days.
 *
 * READ ONLY. Every request is a GET — nothing is written. Safe to re-run before AND after a test.
 *
 * ── How to run ────────────────────────────────────────────────────────────────────────────
 * Paste this whole file into the DevTools console on the CRS site, on any page (site resolution
 * is automatic), signed in as an account that can read both approval libraries.
 */
(async () => {
  "use strict";

  /* ── Configure ──────────────────────────────────────────────────────────────────────── */

  // Leave blank to auto-detect. Set it if the auto-detected site below is wrong.
  const SITE_OVERRIDE = "";

  const LIBRARIES = [
    {
      label: "Normal (Approval for Document)",
      titleCandidates: ["Approval for Document", "Approval Document"],
    },
    {
      label: "HC (Approval for Highly Confidential Document)",
      titleCandidates: ["Approval for Highly Confidential Document", "HC Approval Document"],
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

  const now = new Date();
  const threeDaysAgo = new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000);

  for (const lib of LIBRARIES) {
    console.log(`\n${"#".repeat(70)}\n${lib.label}\n${"#".repeat(70)}`);

    let listTitle;
    for (const cand of lib.titleCandidates) {
      const probe = await get(
        `${web}/_api/web/lists/getbytitle('${encodeURIComponent(cand)}')?$select=Title`,
      );
      if (probe.ok) {
        listTitle = probe.body.Title;
        break;
      }
    }
    if (!listTitle) {
      console.log("✗ Could not resolve this library on this site — skipping.");
      continue;
    }
    console.log(`Library: "${listTitle}"`);

    /* Every pending FILE, unfiltered by age — we compute age/due-status ourselves so a file that
       is NOT yet due still shows up, with a reason, instead of silently vanishing from the list.
       BulkImport is read too, purely to REPORT the exclusion — GetStale's own filter now excludes
       it server-side, but seeing which items would have matched except for that flag is worth
       knowing after the 2026-09-24 incident where a bulk-import file's reminder cycle accidentally
       re-triggered a separate auto-approve flow. */
    const read = await get(
      `${web}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
        `?$select=Id,FileLeafRef,Created,NextReminderAt,BulkImport` +
        `&$filter=${encodeURIComponent("FSObjType eq 0 and OData__ModerationStatus eq 2")}` +
        `&$orderby=Created asc&$top=500`,
    );
    if (!read.ok) {
      console.log(`✗ Could not read pending items (HTTP ${read.status}).`, read.body);
      continue;
    }
    const items = read.body.value || [];
    console.log(`${items.length} document(s) currently pending.\n`);

    const dueNow = [];
    const notYetDue = [];
    const excludedBulk = [];
    const noCreated = [];

    for (const item of items) {
      if (!item.Created) {
        noCreated.push(item);
        continue;
      }
      const created = new Date(item.Created);
      const ageDays = (now.getTime() - created.getTime()) / (24 * 60 * 60 * 1000);
      const nextReminderAt = item.NextReminderAt ? new Date(item.NextReminderAt) : null;

      // Mirrors GetStale's own filter exactly:
      //   BulkImport ne 1 and (
      //     (NextReminderAt eq null and Created le now-3days)
      //     or (NextReminderAt ne null and NextReminderAt le now)
      //   )
      // ⚠ The filter uses "ne 1", not "ne true" — SharePoint's REST $filter on this endpoint does
      // not reliably evaluate the bare `true`/`false` keyword against a Yes/No column (confirmed
      // live 2026-09-24: `BulkImport eq true` and `BulkImport eq false` returned the IDENTICAL
      // result set, while `BulkImport ne 1`/`eq 0` correctly distinguished them). This script's own
      // classification below reads the JSON value directly (`item.BulkImport === true`), which is
      // unaffected by that quirk — only the flow's server-side $filter needed the numeric form.
      let due;
      let reason;
      if (!nextReminderAt) {
        due = created.getTime() <= threeDaysAgo.getTime();
        reason = due
          ? `never reminded, ${ageDays.toFixed(1)} days old (>= 3) — would match on the FIRST branch`
          : `never reminded, only ${ageDays.toFixed(1)} days old (< 3) — not due yet`;
      } else {
        due = nextReminderAt.getTime() <= now.getTime();
        const daysUntil = (nextReminderAt.getTime() - now.getTime()) / (24 * 60 * 60 * 1000);
        reason = due
          ? `NextReminderAt (${item.NextReminderAt}) has passed — would match on the SECOND branch`
          : `NextReminderAt (${item.NextReminderAt}) is ${daysUntil.toFixed(1)} days in the future — not due yet`;
      }

      const line = `Id ${item.Id}  "${item.FileLeafRef}"  Created=${item.Created}  NextReminderAt=${item.NextReminderAt || "(blank)"}\n    → ${reason}`;

      if (item.BulkImport === true) {
        if (due) excludedBulk.push(line + "\n    → EXCLUDED from GetStale by \"BulkImport ne true\", even though the date/clock logic alone would have matched");
        // A bulk-import item that is not yet due isn't worth calling out separately — it wouldn't
        // have matched either way, so it just falls into the ordinary "not yet due" bucket.
        else notYetDue.push(line);
      } else if (due) {
        dueNow.push(line);
      } else {
        notYetDue.push(line);
      }
    }

    console.log(`${"=".repeat(70)}\nDUE FOR A REMINDER RIGHT NOW (${dueNow.length})\n${"=".repeat(70)}`);
    console.log(dueNow.length ? dueNow.join("\n\n") : "(none)");

    console.log(`\n${"=".repeat(70)}\nEXCLUDED — BULK IMPORT, WOULD OTHERWISE BE DUE (${excludedBulk.length})\n${"=".repeat(70)}`);
    console.log(excludedBulk.length ? excludedBulk.join("\n\n") : "(none)");

    console.log(`\n${"=".repeat(70)}\nNOT YET DUE (${notYetDue.length})\n${"=".repeat(70)}`);
    console.log(notYetDue.length ? notYetDue.join("\n\n") : "(none)");

    if (noCreated.length) {
      console.log(`\n⚠ ${noCreated.length} item(s) had no readable Created value — skipped:`, noCreated);
    }
  }

  console.log(
    "\nRun this again AFTER a flow's test run finishes: every item that was in \"DUE\" above\n" +
      "should now show a NextReminderAt roughly 3 days from now (Created+6 the first time, since it\n" +
      "started blank), and should have moved into \"NOT YET DUE\" — that's the confirmation the\n" +
      "mechanism actually advanced the clock, not just sent an email. Nothing in the EXCLUDED list\n" +
      "should ever get a NextReminderAt stamp or an email from these flows at all.",
  );
})().catch((e) => console.error("Script error:", e));

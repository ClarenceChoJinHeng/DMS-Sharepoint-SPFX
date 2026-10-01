/**
 * Dumps a segment's whole folder tree across all six CRS libraries and REPORTS WHAT IS WRONG WITH IT.
 *
 * READ ONLY — every request is a GET. Paste into the browser console on any page of a CRS site,
 * signed in as someone who can see the libraries. Nothing is written, moved or deleted.
 *
 * ── Why this exists ──────────────────────────────────────────────────────────
 * After a structure migration the only honest confirmation is the FOLDERS, not the run log. That
 * distinction has already cost this project a day: on 2026-09-07 a migration was recorded as finished
 * from its own report (14 moved, 27 tidied) and the next morning's scan found 131 folders still in
 * the old shape. A run reporting success is not the same as the state being right.
 *
 * Checking that by hand means opening every Year and every Document Type in six libraries — for GHO,
 * 8 departments and 62 units. Nobody does that twice, which is exactly how a half-migrated segment
 * goes unnoticed.
 *
 * ── What it flags, and why each one matters ──────────────────────────────────
 *   MISPLACED FILE   a document not sitting at the end of the chain. The standing rule is that a
 *                    document ALWAYS lands at the end; one sitting shallower is a file the migration
 *                    did not move, and it is invisible from any single folder view.
 *   WRONG SHAPE      a folder holding documents whose parent is not a four-digit Year. Two causes,
 *                    and with an OPTIONAL level they cannot be told apart from the folder tree
 *                    alone: the branch is still in the old order (a reorder missed it), or it is a
 *                    leftover from an older chain at a depth the optional level makes legal. Both
 *                    need the folder opening. Replaced a positional "NOT A YEAR" check that
 *                    compared depth to CHAIN.indexOf("Year") and could not be right for both a unit
 *                    with a Sub Unit and one without.
 *   TOO DEEP         a folder below the end of the chain. Left over from an older chain (an
 *                    "Archive 1" under a removed tier, say). Harmless to browsing, and reconciliation
 *                    will report it as a stray for ever.
 *   EMPTY LEAF       a Document Type folder holding nothing. Usually FINE and reported separately —
 *                    its documents were approved and routed away, leaving the folder behind. It is
 *                    listed because "tidied 27 empty folders" is only checkable against a list.
 *
 * ⚠ IT CANNOT TELL YOU THE METADATA IS RIGHT. A reorder re-stamps zero documents by design, and a
 * file uploaded during a broken-chain window keeps a blank Year or Document Type that nothing
 * repairs and no folder view shows. After running this, still filter those two columns for blanks.
 *
 * ── Two traps this script is built around ────────────────────────────────────
 * 1. GetFolderByServerRelativeUrl is called with an OData PARAMETER ALIAS, never an inline quoted
 *    path. An inline literal returns HTTP 400 — not 404 — once the path is deep enough (~330 chars,
 *    six levels), because of the many encoded slashes in one literal. Form.tsx used the inline form
 *    and silently reported real folders as missing.
 * 2. The library is addressed by its URL SEGMENT, never its title. A rename changes the title and
 *    never the URL, and this client renames libraries routinely — twice in one afternoon on
 *    2026-08-27. So "Restricted & Confidential Document" is still at /Shared Documents.
 */
(async () => {
  /* ── Configure ─────────────────────────────────────────────────────────── */

  const SEGMENT = "MHO"; // the segment's TOP FOLDER name (StagingFolder), not its label
  /* The chain BELOW the segment folder, in order.
     A plain string is a level EVERY path has. `{ name, optional: true }` is a PER-UNIT level that
     only some units have — a Sub Unit authored under some unit terms and not others.
     GHO: ["Department", "Unit", "Year", "Document Type"]
     MHO: as below — Sub Unit is optional, so a leaf sits at 5 OR 6 levels.
     Buah: ["Buah Category", "Buah Number", "State", "Year", "Document Type"] */
  const CHAIN = [
    "Department",
    "Unit",
    { name: "Sub Unit", optional: true },
    "Minasmas Archive 2",
    "Year",
    "Document Type",
  ];
  /* ⚠ AN OPTIONAL LEVEL BREAKS EVERY POSITIONAL CHECK, which is why they are all derived here.
     With Sub Unit optional, a leaf is at MIN_DEPTH (units without one) or MAX_DEPTH (units with
     one) — so "the folder at position 4 is the Year" is simply false half the time, and a fixed
     depth would report every unit without a sub unit as MISPLACED. Running it that way would bury
     the real findings in false positives, which is worse than not running it at all. */
  const NAMES = CHAIN.map((c) => (typeof c === "string" ? c : c.name));
  const MAX_DEPTH = CHAIN.length;
  const MIN_DEPTH = CHAIN.filter((c) => typeof c === "string").length;
  const SHAPE = NAMES.join("/") + (MIN_DEPTH === MAX_DEPTH ? "" : "  (Sub Unit optional)");

  /* ⚠ THE YEAR IS FOUND FROM THE LEAF, NOT FROM A POSITION — the one check that survives an optional
     level. Whatever the depth, the last two segments of a document's folder are always
     `<Year>/<Document Type>`, so a branch still in the OLD order fails this wherever it sits. */
  const yearOk = (parts) => parts.length >= 2 && /^\d{4}$/.test(parts[parts.length - 2]);

  const CONCURRENCY = 6; // ten simultaneous calls is where a tenant starts answering 429

  /* ⚠ URL SEGMENTS, NOT TITLES — see trap 2 above. "Shared Documents" really does have a space in
     it; that is the Documents library's segment, and it is the one people get wrong. */
  const LIBS = [
    ["Approval (normal)", "ApprovalDocument"],
    ["Documents (normal)", "Shared Documents"],
    ["Approval (HC)", "HCApprovalDocument"],
    ["Documents (HC)", "HCDocuments"],
    ["Archive (normal)", "Archive"],
    ["Archive (HC)", "HCArchive"],
  ];

  /* ── Site ──────────────────────────────────────────────────────────────── */

  /* ⚠ FROM SHAREPOINT, NOT THE ADDRESS BAR. Splitting location.href produces a wrong base on a
     /Lists/ or /Forms/ page, and the symptom is not a clean 404 — it is a sign-in page parsed as
     JSON: Unexpected token '<', "<!DOCTYPE". The split stays only as a last resort. */
  const spx = window._spPageContextInfo || {};
  const site = (spx.webAbsoluteUrl
    || (spx.webServerRelativeUrl ? location.origin + spx.webServerRelativeUrl : "")
    || location.href.split("/_layouts")[0].split("/SitePages")[0]).replace(/\/$/, "");
  const webRel = new URL(site).pathname.replace(/\/$/, "");
  console.log("site: " + site);
  console.log("segment: " + SEGMENT + "   chain: " + SHAPE);

  /* ── Fetch, with the retry the search page had to learn ────────────────── */

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  /* 429 and 503 ONLY. A 400 is a malformed request, a 403 is permissions and a 404 is an absent
     folder — retrying any of those asks the same unanswerable question twice and hides the status. */
  const get = async (url) => {
    for (let attempt = 0; ; attempt++) {
      const res = await fetch(url, {
        headers: { Accept: "application/json;odata=nometadata" },
        credentials: "same-origin",
      });
      if (res.ok) return res.json();
      if ((res.status === 429 || res.status === 503) && attempt < 4) {
        const after = Number(res.headers.get("Retry-After"));
        await sleep(Number.isFinite(after) && after > 0 ? after * 1000 : Math.min(8000, 500 * 2 ** attempt));
        continue;
      }
      const err = new Error("HTTP " + res.status);
      err.status = res.status;
      throw err;
    }
  };

  /* One request per folder: its subfolders and its files together. The alias form is trap 1. */
  const read = async (path) => {
    const url = site + "/_api/web/GetFolderByServerRelativeUrl(@f)?@f='"
      + encodeURIComponent(path) + "'&$expand=Folders,Files"
      + "&$select=Folders/Name,Folders/ServerRelativeUrl,Files/Name";
    const d = await get(url);
    return { folders: d.Folders || [], files: (d.Files || []).map((f) => f.Name) };
  };

  /* ── Walk ──────────────────────────────────────────────────────────────── */

  const problems = [];
  const tree = [];
  let folderCount = 0;
  let fileCount = 0;

  const flag = (kind, library, path, note) => problems.push({ kind, library, path, note });

  /* A small pool rather than Promise.all over everything: GHO is thousands of folders across six
     libraries, and firing them all at once is how a tenant starts refusing them. */
  const pool = async (jobs) => {
    let i = 0;
    const workers = Array.from({ length: Math.min(CONCURRENCY, jobs.length) }, async () => {
      while (i < jobs.length) await jobs[i++]();
    });
    await Promise.all(workers);
  };

  const walk = async (libLabel, path, depth) => {
    let here;
    try {
      here = await read(path);
    } catch (e) {
      /* A 404 at the very top means the segment is simply not in this library, which is a normal
         state — nothing archived yet, or no HC documents. Deeper down it is worth saying. */
      if (e.status === 404 && depth === 0) return;
      flag("UNREADABLE", libLabel, path, e.message);
      return;
    }
    folderCount++;
    fileCount += here.files.length;

    const cut = path.indexOf("/" + SEGMENT + "/");
    const rel = cut === -1 ? SEGMENT : path.slice(cut + 1);
    tree.push({
      library: libLabel, depth, path: rel,
      folders: here.folders.length, files: here.files.length,
    });

    const parts = rel.split("/");
    // At or past the shallowest legitimate leaf. NOT an equality any more — see MIN/MAX_DEPTH.
    const atEnd = depth >= MIN_DEPTH;

    /* ⚠ TOO DEEP IS TESTED FIRST AND RETURNS, so a leftover folder is named ONCE. Ordering the
       misplaced-file check above this reported the same folder twice — once as TOO DEEP and once as
       a misplaced file — and printed "depth 5 (undefined)", because CHAIN[depth - 1] is off the end
       of the chain down here. Caught by running the script against a synthetic tree rather than by
       reading it. The file count goes in the note: a leftover holding documents is a different
       problem from an empty one, and that is the part worth acting on. */
    if (depth > MAX_DEPTH) {
      flag("TOO DEEP", libLabel, rel,
        (here.files.length > 0 ? here.files.length + " file(s), " : "empty, ")
        + "below the end of " + SHAPE + " — left over from an older structure");
      return; // descending would report every child of it too
    }

    /* ⚠ THE TWO FILE FAULTS ARE MUTUALLY EXCLUSIVE, deliberately. Reporting one folder under both
       kinds is the double-report bug this script already shipped once (see TOO DEEP above), and it
       makes the problem table longer than the truth.
         too shallow  → never moved            → MISPLACED FILE
         deep enough, wrong year position → moved into the old ORDER → NOT A YEAR */
    if (here.files.length > 0 && !atEnd) {
      flag("MISPLACED FILE", libLabel, rel,
        here.files.length + " file(s) at the "
        + (depth === 0 ? "segment folder" : NAMES[depth - 1] + " level") + " — expected them under "
        + SHAPE);
    } else if (here.files.length > 0 && !yearOk(parts)) {
      /* ⚠ TWO CAUSES, AND AN OPTIONAL LEVEL MAKES THEM INDISTINGUISHABLE FROM HERE — so the note
         names both rather than asserting one. Either this branch is still in the OLD ORDER (the
         reorder missed it), or it is a LEFTOVER folder from an older chain sitting at a depth the
         optional Sub Unit makes legal: `…/2024/Legal Opinion/Archive 1` is six levels down, exactly
         like a valid leaf under a unit that HAS a sub unit. Telling them apart needs the Sub Unit
         and Archive-2 term lists, which this script deliberately does not read.
         Both need the folder opening, which is what the report is for. */
      flag("WRONG SHAPE", libLabel, rel,
        here.files.length + ' file(s) whose folder is not <Year>/<Document Type> — "'
        + parts[parts.length - 2] + '" sits where the Year should be. Either this branch is still in'
        + " the old order, or it is a leftover folder from an older chain. Open it.");
    }
    if (atEnd && here.files.length === 0 && here.folders.length === 0) {
      flag("EMPTY LEAF", libLabel, rel, "no documents and no subfolders");
    }

    /* ⚠ THE OLD POSITIONAL YEAR CHECK IS GONE — it compared `depth` to `CHAIN.indexOf("Year")`,
       which cannot be right for both a unit with a Sub Unit and one without. The leaf-relative
       version above replaces it and is strictly stronger: it catches a wrong order at ANY depth. */
    const jobs = here.folders.map((f) => async () => {
      await walk(libLabel, f.ServerRelativeUrl, depth + 1);
    });
    await pool(jobs);
  };

  /* ── Run ───────────────────────────────────────────────────────────────── */

  const started = Date.now();
  for (const pair of LIBS) {
    console.log("walking " + pair[0] + " ...");
    await walk(pair[0], webRel + "/" + pair[1] + "/" + SEGMENT, 0);
  }
  const secs = Math.round((Date.now() - started) / 1000);

  /* ── Report ────────────────────────────────────────────────────────────── */

  console.log("%c" + folderCount + " folder(s), " + fileCount + " file(s) read in " + secs + "s",
    "font-weight:700");

  const empties = problems.filter((p) => p.kind === "EMPTY LEAF");
  const real = problems.filter((p) => p.kind !== "EMPTY LEAF");

  if (real.length === 0) {
    console.log("%cNo misplaced files, no out-of-order Year folders, nothing below the end of the "
      + "chain. " + empties.length + " empty leaf folder(s), which is normal.",
      "color:#0f6c3f;font-weight:700");
  } else {
    console.log("%c" + real.length + " problem(s) — every one of these is worth opening",
      "color:#a4262c;font-weight:700");
    console.table(real);
  }
  if (empties.length > 0) {
    console.log("%c" + empties.length + " empty leaf folder(s) — normal, listed so a tidy count is "
      + "checkable", "color:#8a8886");
    console.table(empties);
  }

  /* Left on window so the tree can be read, filtered or pasted into a ticket without re-running the
     walk — which on GHO is thousands of requests. */
  window.crsDump = { tree, problems, folderCount, fileCount };
  console.log("Full tree in window.crsDump.tree — copy(JSON.stringify(crsDump, null, 2)) puts it on "
    + "the clipboard.");

  /* ⚠ SAID ON EVERY RUN, because it is the one thing this cannot check and this is the only moment
     anyone would. */
  console.log("%cStill to check by hand: filter Year and Document Type for BLANKS. A file uploaded "
    + "during a broken-chain window keeps an empty value that no folder view shows and no backfill "
    + "repairs.", "color:#8a4b00");
})();

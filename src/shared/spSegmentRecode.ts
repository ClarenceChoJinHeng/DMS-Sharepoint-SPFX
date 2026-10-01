import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { probeFolderByPath, renameFolder } from "./dmsFolderMap";
import { sanitizeFolderSegment } from "./formModel";
import { archiveAvailable, cachedListTitle, libraryTargets, LIST_SUFFIX } from "./naming";
import { ExistingSegment } from "./newSegment";
import { SegmentCounts, unknownCounts } from "./segmentDeletion";
import { primeNames } from "./spNaming";

/**
 * Segment lifecycle primitives — the folder/file/row reads and writes shared by the Delete and
 * Re-code flows, plus the two SPFx-facing entry points those flows are built on:
 * `countSegmentDocuments` and `performSegmentRecode`.
 *
 * Extracted from `SegmentCreator.tsx` (`countSegment`, `onRecode`, and the primitives they both
 * depended on) so a SECOND screen — the CRS Term Abbreviations page — can offer the same re-code
 * action without a second, drifting copy of destructive-write logic. See
 * docs/superpowers/specs/2026-09-15-segment-rename-and-group-rename-design.md.
 *
 * Everything here reads/writes SharePoint directly; nothing is testable in isolation the way
 * `segmentRecode.ts`'s pure rules are. Callers own all component state (busy flags, progress bars,
 * result banners, the audit-row write, and reloading their own segment list) — this module only
 * ever returns a plain result, never a React state update.
 */

/* ── Which libraries a retire/recode actually touches ─────────────────────────────────────────
   ONE definition, read by the count, the delete flow and the recode flow alike. See the original
   comment at `SegmentCreator.tsx` (moved here 2026-09-15) for the two-bug history this replaced —
   the archive pair is excluded deliberately (7-year retention outlives the segment). */

export function retireLibraries(): { key: string; title: string; urlSegment: string }[] {
  return libraryTargets().filter((t) => t.key !== "Archive" && t.key !== "ArchiveHC");
}

/** The other half of `retireLibraries()` — Archive/ArchiveHC only, whichever exist on this site. */
export function archiveTargets(): { key: string; title: string; urlSegment: string }[] {
  return libraryTargets().filter((t) => t.key === "Archive" || t.key === "ArchiveHC");
}

/** Every subfolder path beneath `root`, depth-first. Throws — the caller reports `unknown`. */
export async function walkFolders(
  context: WebPartContext,
  siteUrl: string,
  lib: string,
  root: string,
): Promise<string[]> {
  const found: string[] = [];
  const queue = [`${siteUrl.replace(/^https?:\/\/[^/]+/, "")}/${lib}/${root}`];
  while (queue.length > 0 && found.length < 5000) {
    const here = queue.shift() as string;
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Folders?@f='${encodeURIComponent(here)}'&$select=ServerRelativeUrl,Name`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    // 404 = this branch does not exist in this library, which is not an error: a segment can have
    // folders in one library and not the other.
    if (res.status === 404) continue;
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const kids = ((await res.json()).value ?? []) as Array<{ ServerRelativeUrl?: string; Name?: string }>;
    for (const k of kids) {
      const url = (k.ServerRelativeUrl ?? "").trim();
      // Forms is SharePoint's own; it is not part of anyone's segment.
      if (!url || (k.Name ?? "") === "Forms") continue;
      found.push(url);
      queue.push(url);
    }
  }
  return found;
}

/**
 * Files directly in one folder. Throws — the caller reports `unknown`.
 *
 * ⚠ NEVER `Files/$count` — see the original note at `SegmentCreator.tsx`'s history: that scalar
 * OData segment throws unconditionally on SDG's real tenant, and a naive 404-means-empty handler
 * conflated "SharePoint refuses this endpoint shape" with "confirmed empty". List the collection
 * and count the array instead.
 */
export async function countFiles(
  context: WebPartContext,
  siteUrl: string,
  folderUrl: string,
): Promise<number> {
  const res: SPHttpClientResponse = await context.spHttpClient.get(
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Files?@f='${encodeURIComponent(folderUrl)}'&$select=Name&$top=5000`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (res.status === 404) return 0;
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const files = ((await res.json()).value ?? []) as unknown[];
  return files.length;
}

/**
 * Documents in Archive/ArchiveHC under a raw top-folder CODE, before any segment row exists for it
 * — used by both the recode gate and the new-segment archive-clash guard. `undefined` on any
 * failure, including "there is genuinely no archive on this site" (callers gate on
 * `archiveAvailable()` themselves before reaching here) — never guess empty.
 */
export async function countArchiveDocumentsForCode(
  context: WebPartContext,
  siteUrl: string,
  rawRoot: string,
): Promise<number | undefined> {
  const root = rawRoot.trim();
  if (!root) return undefined;
  const targets = archiveTargets();
  if (targets.length === 0) return undefined;
  try {
    let total = 0;
    for (const target of targets) {
      const lib = target.urlSegment;
      const rootUrl = `${siteUrl.replace(/^https?:\/\/[^/]+/, "")}/${lib}/${root}`;
      total += await countFiles(context, siteUrl, rootUrl);
      const subs = await walkFolders(context, siteUrl, lib, root);
      for (const f of subs) total += await countFiles(context, siteUrl, f);
    }
    return total;
  } catch {
    return undefined;
  }
}

export function countArchiveDocuments(
  context: WebPartContext,
  siteUrl: string,
  seg: ExistingSegment,
): Promise<number | undefined> {
  return countArchiveDocumentsForCode(context, siteUrl, seg.stagingFolder ?? "");
}

/** The Group Map rows carrying this segment, matched on the term-set GUID stored in `Segment`. */
export async function loadSegmentGroupMapRows(
  context: WebPartContext,
  siteUrl: string,
  seg: ExistingSegment,
): Promise<number[]> {
  const guid = (seg.termSetGuid ?? "").trim();
  if (!guid) return [];
  const list = encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap));
  const res: SPHttpClientResponse = await context.spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('${list}')/items?$select=Id,Segment&$top=5000`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const rows = ((await res.json()).value ?? []) as Array<{ Id: number; Segment?: string }>;
  return rows
    .filter((r) => (r.Segment ?? "").trim().toLowerCase() === guid.toLowerCase())
    .map((r) => r.Id);
}

/** How many term-store requests the ancestry walk below may spend before giving up. */
const MAX_TERM_WALK_REQUESTS = 400;

/**
 * Every term GUID under this segment's term set, at any depth. The only reliable way to answer
 * "does this abbreviation row belong to segment X", since abbreviation rows carry no segment field
 * of their own. Bounded and reported INCOMPLETE rather than partial — an outage that stops the walk
 * part-way must never look like "this segment has only 6 abbreviation rows".
 */
export async function loadSegmentTermGuids(
  context: WebPartContext,
  siteUrl: string,
  seg: ExistingSegment,
): Promise<{ complete: boolean; guids: Set<string> }> {
  const setGuid = (seg.termSetGuid ?? "").trim();
  if (!setGuid) return { complete: true, guids: new Set() };
  const guids = new Set<string>();
  let requests = 0;
  let frontier: string[] = [""]; // "" walks the set's own top level
  try {
    while (frontier.length > 0) {
      const next: string[] = [];
      const kidLists = await Promise.all(
        frontier.map(async (parentId) => {
          if (requests >= MAX_TERM_WALK_REQUESTS) return [] as Array<{ id?: string }>;
          requests++;
          const url = parentId
            ? `${siteUrl}/_api/v2.1/termStore/sets/${setGuid}/terms/${parentId}/children?$select=id`
            : `${siteUrl}/_api/v2.1/termStore/sets/${setGuid}/children?$select=id`;
          const res: SPHttpClientResponse = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          const data = await res.json();
          return (data.value ?? []) as Array<{ id?: string }>;
        }),
      );
      if (requests >= MAX_TERM_WALK_REQUESTS) return { complete: false, guids };
      for (const kids of kidLists) {
        for (const k of kids) {
          const id = (k.id ?? "").toLowerCase();
          if (id && !guids.has(id)) {
            guids.add(id);
            next.push(k.id as string);
          }
        }
      }
      frontier = next;
    }
    return { complete: true, guids };
  } catch {
    return { complete: false, guids };
  }
}

/** `CRS Term Abbreviation` rows, Id + TermGuid only — all this pass needs to count and delete. */
export async function loadAbbreviationTermRows(
  context: WebPartContext,
  siteUrl: string,
): Promise<Array<{ id: number; termGuid: string }>> {
  const list = encodeURIComponent(cachedListTitle(LIST_SUFFIX.abbreviation));
  const res: SPHttpClientResponse = await context.spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('${list}')/items?$select=Id,TermGuid&$top=5000`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const rows = ((await res.json()).value ?? []) as Array<{ Id: number; TermGuid?: string }>;
  return rows.map((r) => ({ id: r.Id, termGuid: (r.TermGuid ?? "").trim().toLowerCase() }));
}

/**
 * The segment's abbreviation-row ids, matched by walking the live term tree. Shared by the preview
 * count and the actual deletion, so the two can never disagree about which rows belong to this
 * segment.
 */
export async function loadSegmentAbbreviationRowIds(
  context: WebPartContext,
  siteUrl: string,
  seg: ExistingSegment,
): Promise<{ complete: boolean; ids: number[] }> {
  const walk = await loadSegmentTermGuids(context, siteUrl, seg);
  if (!walk.complete) return { complete: false, ids: [] };
  if (walk.guids.size === 0) return { complete: true, ids: [] };
  const rows = await loadAbbreviationTermRows(context, siteUrl);
  return {
    complete: true,
    ids: rows.filter((r) => r.termGuid && walk.guids.has(r.termGuid)).map((r) => r.id),
  };
}

export async function deleteItem(
  context: WebPartContext,
  siteUrl: string,
  listTitle: string,
  itemId: number,
): Promise<void> {
  const res: SPHttpClientResponse = await context.spHttpClient.post(
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE",
      },
    },
  );
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
}

/** Recycle (not purge) a folder, so it is restorable for 93 days. */
export async function recycleFolder(
  context: WebPartContext,
  siteUrl: string,
  lib: string,
  root: string,
): Promise<boolean> {
  const url = `${siteUrl.replace(/^https?:\/\/[^/]+/, "")}/${lib}/${root}`;
  const res: SPHttpClientResponse = await context.spHttpClient.post(
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Recycle()?@f='${encodeURIComponent(url)}'`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (res.status === 404) return false;
  if (!res.ok) throw new Error(`${lib}: HTTP ${res.status}`);
  return true;
}

/**
 * Count what a segment holds, across both non-archive libraries, plus its Group Map rows,
 * abbreviation rows and — where the site has one — its archive.
 *
 * Any failure makes the whole count `unknown` rather than a partial number. A partial count is
 * worse than none here: it would read as authoritative and could show "0 documents" for a segment
 * whose second library simply did not answer.
 */
export async function countSegmentDocuments(
  context: WebPartContext,
  siteUrl: string,
  seg: ExistingSegment,
): Promise<SegmentCounts> {
  const root = (seg.stagingFolder ?? "").trim();
  if (!root) {
    return unknownCounts("this segment has no top folder recorded, so its folders cannot be found");
  }
  let folders = 0;
  let documents = 0;
  try {
    for (const target of retireLibraries()) {
      const lib = target.urlSegment;
      const rootUrl = `${siteUrl.replace(/^https?:\/\/[^/]+/, "")}/${lib}/${root}`;
      documents += await countFiles(context, siteUrl, rootUrl);
      const subs = await walkFolders(context, siteUrl, lib, root);
      folders += subs.length;
      for (const f of subs) documents += await countFiles(context, siteUrl, f);
    }
  } catch (e) {
    return unknownCounts(`the folders could not be read (${(e as Error).message})`);
  }

  // The Group Map read is separate and NOT fatal to the count: a segment can have unreadable
  // folders and readable rows, or the reverse. An unreadable Group Map is reported as unknown too,
  // because deleting rows we could not see is the same class of blind act.
  let groupMapRows = 0;
  try {
    const rows = await loadSegmentGroupMapRows(context, siteUrl, seg);
    groupMapRows = rows.length;
  } catch (e) {
    return unknownCounts(
      `the ${cachedListTitle(LIST_SUFFIX.groupMap)} list could not be read (${(e as Error).message})`,
    );
  }

  // Abbreviation rows this segment's LIVE term tree still covers — retiring deletes these
  // unconditionally, so an incomplete walk makes the whole count unknown, same as an unreadable
  // Group Map above.
  let abbreviationRows = 0;
  try {
    const abbrev = await loadSegmentAbbreviationRowIds(context, siteUrl, seg);
    if (!abbrev.complete) {
      return unknownCounts(
        `the ${cachedListTitle(LIST_SUFFIX.abbreviation)} list or the term store could not be fully read`,
      );
    }
    abbreviationRows = abbrev.ids.length;
  } catch (e) {
    return unknownCounts(
      `the ${cachedListTitle(LIST_SUFFIX.abbreviation)} list could not be read (${(e as Error).message})`,
    );
  }

  // Archive is counted independently and is NEVER allowed to make the whole count `unknown` — a
  // failed or missing archive read only means "leave the archive alone", never "we cannot say
  // anything about this segment at all". Only asked when this site HAS an archive at all.
  const archiveDocuments = archiveAvailable()
    ? await countArchiveDocuments(context, siteUrl, seg)
    : undefined;

  return { state: "counted", folders, documents, groupMapRows, abbreviationRows, archiveDocuments };
}

/**
 * Re-code a segment's top folder by RECYCLING the (confirmed-empty) old tree and letting
 * reconciliation rebuild it fresh under the new name — the write sequence, minus every
 * component-local state update and minus the audit-row write, both of which the caller owns
 * (callers differ on `source:` and on what happens to their own list/reload afterward).
 *
 * ⚠ KEPT, NOT DELETED, BUT NO LONGER CALLED FROM EITHER UI SCREEN (2026-09-16). Both
 * `SegmentCreator.tsx` and `AbbreviationManager.tsx` now always call `performLiveSegmentRecode`
 * instead, even for a confirmed-empty segment — the client watched an empty segment's folders get
 * recycled and asked for that not to happen ("Can we ensure for renaming the term abbreviations
 * doesn't move it to recycle bin if its empty"), and the in-place rename below already handles an
 * empty segment correctly (a library where the old folder does not exist is skipped, not a
 * failure) with strictly less destruction — nothing goes to the recycle bin, and any
 * already-provisioned subfolders survive rather than being blown away. This function is left
 * defined in case a reason to prefer recycling ever comes back.
 *
 * Never throws: every failure is caught and folded into `{ ok: false, renamed, lines }`, so a
 * caller need only render what comes back.
 *
 * ⚠ `ok` AND `renamed` ANSWER DIFFERENT QUESTIONS, AND A CALLER MUST NOT COLLAPSE THEM. `ok` is
 * "did everything go cleanly" — folding in a non-fatal Folder Map tidy-up failure alongside a
 * genuine mode-row write failure. `renamed` is "did the mode row's `StagingFolder` actually get
 * written" — true the moment step 2 succeeds, `false` if the write never happened or was refused.
 * A caller that reflects the new name locally (or reloads expecting to see it) on `ok` alone would
 * do so even when the write never landed, showing a folder name the server does not have.
 *
 * `counts` must already satisfy `canOfferRecode` — this function does not re-check it, the same way
 * `onRecode` never did; the caller is expected to have gated the button on it.
 */
export async function performSegmentRecode(
  context: WebPartContext,
  siteUrl: string,
  seg: ExistingSegment,
  newFolder: string,
): Promise<{ ok: boolean; renamed: boolean; lines: string[] }> {
  const oldRoot = (seg.stagingFolder ?? "").trim();
  const newRoot = sanitizeFolderSegment(newFolder).trim();
  const lines: string[] = [];
  let ok = true;
  let renamed = false;
  try {
    // ⚠ RE-PRIME FIRST, NEVER TRUST THE CALLER. `retireLibraries()` (below) reads
    // `cachedHcLibraries()` — a SYNCHRONOUS module-level cache `naming.ts` owns — and if that
    // hasn't resolved yet (a fast click right after mount, or a caller that forgot to await
    // `primeNames` at all), the HC pair is silently OMITTED from this run rather than reported as
    // missing: the same "capped read reads as absent" shape this codebase keeps re-learning, here
    // applied to a library LIST rather than a paged collection. `primeNames` memoises every probe
    // as a module-level promise, so re-awaiting it once already resolved costs nothing.
    await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);

    // 1. Recycle the OLD (empty, per the caller's gate) tree in every library it might exist in.
    for (const target of retireLibraries()) {
      try {
        const went = await recycleFolder(context, siteUrl, target.urlSegment, oldRoot);
        lines.push(
          went
            ? `${target.title}/${oldRoot} moved to the recycle bin.`
            : `${target.title}/${oldRoot} did not exist.`,
        );
      } catch (e) {
        lines.push(`Could not recycle ${target.title}/${oldRoot} — ${(e as Error).message}`);
        ok = false;
      }
    }

    // 2. The mode row. If THIS fails, stop — a run that recycled the old folders but left the row
    //    naming the old key is a segment now pointing at a folder that no longer exists.
    if (seg.itemId === undefined) {
      throw new Error("this segment's configuration row has no id, so it cannot be updated");
    }
    const write: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items(${seg.itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          // JSON light: no `__metadata`, and `odata-version` blanked because SPFx injects 4.0,
          // under which SharePoint cannot infer the entity set for a light payload.
          "odata-version": "",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        body: JSON.stringify({ StagingFolder: newRoot }),
      },
    );
    if (!write.ok) {
      throw new Error(
        `The old folders were recycled, but the segment row could not be updated (HTTP ` +
          `${write.status}). Fix the "${cachedListTitle(LIST_SUFFIX.config)}" row by hand — its ` +
          `Title is "${seg.key}" — and set StagingFolder to "${newRoot}".`,
      );
    }
    renamed = true;
    lines.push(`The segment's top folder is now "${newRoot}".`);

    // 3. Folder Map rows for the OLD Section — derivable, so they go exactly when the folder does.
    //    A row left behind would point at a UniqueId now in the recycle bin.
    try {
      const list = encodeURIComponent(cachedListTitle(LIST_SUFFIX.folderMap));
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${list}')/items?$select=Id,Section&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (res.ok) {
        const rows = ((await res.json()).value ?? []) as Array<{ Id: number; Section?: string }>;
        const mine = rows.filter(
          (r) => (r.Section ?? "").trim().toLowerCase() === oldRoot.toLowerCase(),
        );
        let gone = 0;
        for (const r of mine) {
          try {
            await deleteItem(context, siteUrl, cachedListTitle(LIST_SUFFIX.folderMap), r.Id);
            gone++;
          } catch {
            ok = false;
          }
        }
        if (mine.length > 0) lines.push(`Folder Map rows removed: ${gone} of ${mine.length}.`);
      }
    } catch {
      lines.push("The Folder Map rows could not be tidied up; reconciliation will report them.");
      ok = false;
    }

    lines.push("Run Folder Reconciliation to rebuild the folder tree under the new name.");
    return { ok, renamed, lines };
  } catch (e) {
    return { ok: false, renamed, lines: [...lines, `Stopped: ${(e as Error).message}`] };
  }
}

/**
 * Re-code a LIVE segment's top folder — rename it IN PLACE, in every library including the
 * archive, preserving every document, version, approval status and permission inside. This is
 * the write sequence for a segment that HOLDS documents (`canOfferRecode` is false,
 * `canOfferAnyRecode` is true) — the counterpart to `performSegmentRecode` above, which only ever
 * handles the confirmed-empty case by recycling the old tree and letting reconciliation rebuild.
 *
 * Uses `renameFolder` — the same `MoveTo`-within-parent primitive reconciliation already uses to
 * rename a live Department or Unit folder in place — never a recycle-and-recreate. `MoveTo` within
 * one parent changes only the PATH; the folder's UniqueId, every document beneath it, every
 * version, every approval status and every ACL survive untouched.
 *
 * Two phases, and the order is deliberate:
 *   1. PRE-CHECK every library (including the archive) for a collision at the NEW name via
 *      `probeFolderByPath`. Anything other than a CONFIRMED-missing answer — found, or could not
 *      tell — refuses the whole operation with NOTHING changed. A live segment's real documents
 *      must never be moved on the strength of a guess, and "could not tell" is exactly the
 *      uncertain case this project treats as unsafe to write through.
 *   2. Rename the segment's top folder in each library where it currently exists. A library with
 *      no such folder (404) is SKIPPED, not a failure — the same tolerance
 *      `retireLibraries()`-based callers already show a missing library. But if a rename is
 *      ATTEMPTED (the folder existed) and it fails, the whole run STOPS right there, before
 *      touching the mode row or any Folder Map row: a segment split across two folder names in
 *      different libraries is already a bad state, and letting the configuration also disagree
 *      with what SharePoint actually holds would make it worse, not better. Only once EVERY
 *      attempted rename has succeeded does the mode row's `StagingFolder` get written, followed by
 *      the Folder Map rows' `Section` field — MERGED from the old value to the new one, never
 *      deleted, because the underlying folders and their UniqueIds are unchanged; only the
 *      denormalized label naming which segment they belong to needs to catch up.
 *
 * Never throws: every failure is caught and folded into `{ ok: false, renamed, lines }`.
 *
 * `counts` must already satisfy `canOfferAnyRecode` (the count succeeded, whatever it found) —
 * this function does not re-check it, the same way `performSegmentRecode` trusts its own gate.
 */
export async function performLiveSegmentRecode(
  context: WebPartContext,
  siteUrl: string,
  seg: ExistingSegment,
  newFolder: string,
): Promise<{ ok: boolean; renamed: boolean; lines: string[] }> {
  const oldRoot = (seg.stagingFolder ?? "").trim();
  const newRoot = sanitizeFolderSegment(newFolder).trim();
  const lines: string[] = [];
  const siteRelative = siteUrl.replace(/^https?:\/\/[^/]+/, "");
  // ⚠⚠ RE-PRIME FIRST, NEVER TRUST THE CALLER — found live 2026-09-16, where a segment's Archive
  // folder was left behind under its OLD name while every other library correctly renamed:
  // `libraryTargets()` (next line) reads `cachedArchiveLibraries()`/`cachedHcLibraries()`, both
  // SYNCHRONOUS module-level caches `naming.ts` owns and fills only once `primeArchiveLibraries`/
  // `primeHcLibraries` resolve. If that hasn't happened yet when this runs (a fast click right
  // after the hosting screen mounts, or a caller that forgot to await `primeNames` at all), the
  // archive pair is silently OMITTED from `targets` below — not reported missing, not skipped with
  // a "no such folder" log line, simply never attempted. The rename then succeeds everywhere it WAS
  // attempted, the mode row's `StagingFolder` flips to the new code, and the NEXT full
  // reconciliation run — which by then has had plenty of time to warm the cache — ensure-creates a
  // brand-new folder under the new name in Archive, leaving the untouched old one sitting right
  // beside it. Same "capped read reads as absent" shape this codebase keeps re-learning, here
  // applied to a library LIST rather than a paged collection.
  //
  // `primeNames` memoises every probe as a module-level promise, so re-awaiting it once a caller
  // has already primed costs nothing — this is defence against a caller that has NOT, not a second
  // definition of when priming happens.
  await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
  const targets = libraryTargets(); // every library on this site, archive pair included
  let ok = true;
  let renamed = false;
  try {
    // 1. Pre-check every library for a collision at the NEW name. Only a confirmed-missing answer
    //    counts as safe — a found folder OR an unconfirmed status both refuse, with nothing
    //    changed anywhere.
    for (const target of targets) {
      const path = `${siteRelative}/${target.urlSegment}/${newRoot}`;
      const probe = await probeFolderByPath(context.spHttpClient, siteUrl, path);
      if (!probe.confirmedMissing) {
        const why = probe.folder
          ? "a folder already exists there"
          : `its status could not be confirmed (HTTP ${probe.status})`;
        return {
          ok: false,
          renamed: false,
          lines: [
            `Stopped before changing anything: "${newRoot}" in ${target.title} — ${why}. Pick a ` +
              `different name, or clear that folder first.`,
          ],
        };
      }
    }
    lines.push(`"${newRoot}" is free in every library — nothing to collide with.`);

    // 2. Rename the top folder in place, wherever it currently exists.
    let renameFailed = false;
    for (const target of targets) {
      const currentUrl = `${siteRelative}/${target.urlSegment}/${oldRoot}`;
      const result = await renameFolder(context.spHttpClient, siteUrl, currentUrl, newRoot);
      if (result.ok) {
        lines.push(`${target.title}/${oldRoot} renamed to ${target.title}/${newRoot}.`);
      } else if (result.status === 404) {
        // No such folder in this library at all — not a failure, nothing to rename here.
        lines.push(`${target.title} has no "${oldRoot}" folder — skipped.`);
      } else {
        renameFailed = true;
        const why = result.conflict
          ? "a folder with the new name already exists there"
          : `HTTP ${result.status}${result.detail ? ` — ${result.detail}` : ""}`;
        lines.push(`Could not rename ${target.title}/${oldRoot} — ${why}.`);
      }
    }
    if (renameFailed) {
      lines.push(
        "Stopped: at least one library's rename failed. The segment's configuration was NOT " +
          "changed, so libraries that DID rename may now disagree with the ones that did not — " +
          "check each library named above and fix it by hand before trying again.",
      );
      return { ok: false, renamed: false, lines };
    }

    // 3. The mode row. If THIS fails, stop — every folder is already renamed, so leaving the row
    //    naming the old key would mean reconciliation looks for folders that no longer exist.
    if (seg.itemId === undefined) {
      throw new Error("this segment's configuration row has no id, so it cannot be updated");
    }
    const write: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items(${seg.itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "odata-version": "",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        body: JSON.stringify({ StagingFolder: newRoot }),
      },
    );
    if (!write.ok) {
      throw new Error(
        `Every folder was renamed, but the segment row could not be updated (HTTP ` +
          `${write.status}). Fix the "${cachedListTitle(LIST_SUFFIX.config)}" row by hand — its ` +
          `Title is "${seg.key}" — and set StagingFolder to "${newRoot}".`,
      );
    }
    renamed = true;
    lines.push(`The segment's top folder is now "${newRoot}".`);

    // 4. Folder Map rows for the OLD Section — UPDATED, never deleted. The folders themselves and
    //    their UniqueIds did not change, only their path; only the denormalized `Section` label
    //    needs re-pointing.
    try {
      const list = encodeURIComponent(cachedListTitle(LIST_SUFFIX.folderMap));
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${list}')/items?$select=Id,Section&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (res.ok) {
        const rows = ((await res.json()).value ?? []) as Array<{ Id: number; Section?: string }>;
        const mine = rows.filter(
          (r) => (r.Section ?? "").trim().toLowerCase() === oldRoot.toLowerCase(),
        );
        let repointed = 0;
        for (const r of mine) {
          try {
            const merge: SPHttpClientResponse = await context.spHttpClient.post(
              `${siteUrl}/_api/web/lists/getbytitle('${list}')/items(${r.Id})`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: "application/json;odata=nometadata",
                  "Content-Type": "application/json;odata=nometadata",
                  "odata-version": "",
                  "X-HTTP-Method": "MERGE",
                  "IF-MATCH": "*",
                },
                body: JSON.stringify({ Section: newRoot }),
              },
            );
            if (merge.ok) repointed++;
            else ok = false;
          } catch {
            ok = false;
          }
        }
        if (mine.length > 0) {
          lines.push(`Folder Map rows re-pointed to the new name: ${repointed} of ${mine.length}.`);
        }
      } else {
        lines.push("The Folder Map rows could not be checked; reconciliation may report them.");
        ok = false;
      }
    } catch {
      lines.push("The Folder Map rows could not be re-pointed; reconciliation may report them.");
      ok = false;
    }

    lines.push("Run Folder Reconciliation to confirm everything still resolves.");
    return { ok, renamed, lines };
  } catch (e) {
    return { ok: false, renamed, lines: [...lines, `Stopped: ${(e as Error).message}`] };
  }
}

import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, parseLevels, stripFolderChars } from "../../../shared/formModel";
import { isAbbreviatedLevel, splitChain } from "../../../shared/folderChain";

/**
 * Grouping-key prefix for the top terms of a SHARED level's own term set.
 *
 * ⚠ Sibling uniqueness groups on `parentGuid`, and a top-level term carries "". Without a synthetic
 * key two coded shared sets would be judged siblings OF EACH OTHER, and of the segment's own
 * departments — so `ARC` in one set would clash with `ARC` in another, which are not siblings at
 * all. `findCollisions` treats the key as opaque, which is what makes this legitimate.
 */
const SHARED_SET_KEY = "set:";
import { abbrevListTitle } from "../../../shared/folderAbbreviation";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { Toast, ToastKind } from "../../../shared/toast";
import {
  AbbrevRowDraft,
  changedRows,
  groupRowsByParent,
  hasBlockingProblem,
  renamingRows,
  RowProblem,
  validateRows,
} from "../../../shared/abbreviationDraft";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

/**
 * Term Abbreviations — name the folders a segment's terms produce.
 *
 * Spec: docs/superpowers/specs/2026-08-12-term-abbreviation-page-design.md
 *
 * Folder names come from the abbreviation list, keyed by term GUID, and a term with no code is
 * SKIPPED by every reconciliation run — no folder, and that unit cannot upload. The list was
 * previously maintained in the SharePoint list view or by pasting a console script; the client will
 * do neither after handover.
 *
 * THE PAGE NEVER SHOWS A GUID. An admin picks a segment; the tree and the keys are resolved here.
 *
 * Saving writes list rows and NOTHING else — no folder is created or renamed until someone runs
 * Folder Reconciliation. That keeps this screen inert and repeatable, and leaves the destructive step
 * behind its own button.
 */

interface SegmentOption {
  key: string;
  label: string;
  termSetGuid: string;
  stagingFolder: string;
  /** Permissioned level names, shallowest first — these are the tiers that need codes. */
  levelNames: string[];
  /** Below-Unit tiers, which name their folders from the term label and need no code. */
  belowNames: string[];
  /**
   * The below-Unit chain itself, taken from `PendingLevels ?? Levels`.
   *
   * ⚠ THE PENDING CHAIN, NOT THE LIVE ONE, and that is load-bearing. Turning codes on for a level
   * STAGES the change; this screen is the very next step and has to show that level's terms before
   * the migration applies it. Read the live chain and a newly coded level's terms never appear on
   * the one step whose job is forcing their codes in.
   *
   * Only levels with `abbreviated` need rows at all — an uncoded level names its folders from the
   * label and needs no code, which is why the walk was capped in the first place.
   */
  below: Level[];
}

/** A term resolved from the store, with its position in the tree. */
interface TermNode {
  id: string;
  label: string;
  parentId: string;
  depth: number; // 1-based
}

const s: Record<string, React.CSSProperties> = {
  msg: { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.5 },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { ...NOTICE_ATTENTION },
  ok: { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  info: { background: "#f3f2f1", border: "1px solid #e1dfdd", color: "#323130" },
  label: { display: "block", fontSize: 12, fontWeight: 600, color: "#323130", margin: "0 0 4px" },
  select: { padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4, minWidth: 260 },
  input: { padding: "5px 8px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4, width: "100%", boxSizing: "border-box" },
  inputBad: { borderColor: "#a4262c", background: "#fdf3f3" },
  btn: { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "8px 16px", fontSize: 13, cursor: "pointer" },
  off: { background: "#f3f2f1", color: "#a19f9d", border: "1px solid #e1dfdd", borderRadius: 4, padding: "8px 16px", fontSize: 13, cursor: "not-allowed" },
  ghost: { background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4, padding: "4px 10px", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" },
  /**
   * A GRID, not a flex row. As flex the term label grew and the path hint shrank by content length,
   * so a 14-row tier had its inputs at 14 different offsets — and the one thing an admin scans this
   * page for is the EMPTY box. Fixed tracks keep every column aligned whatever the label is called.
   */
  row: {
    display: "grid",
    /* ⚠ THE FIELD SITS BESIDE THE TERM NAME (client, 2026-09-06: *"Move ALL text fields and place
       it beside the folder title"*). It used to be pushed right by a `1fr` label column, so on a
       wide screen the box a code goes in was half a screen away from the term it names.
       A FIXED label column rather than `max-content`, so every input still lines up: with
       `max-content` the boxes would step in and out with the length of each term. */
    gridTemplateColumns: "minmax(0, 260px) minmax(0, 200px) minmax(0, 1fr)",
    alignItems: "center",
    columnGap: 12,
    rowGap: 4,
    padding: "8px 10px",
    borderBottom: "1px solid #f3f2f1",
  },
  scrollBox: { maxHeight: "58vh", overflowY: "auto", overflowX: "hidden", border: "1px solid #ececec", borderRadius: 6, padding: "10px 12px", background: "#fff", marginBottom: 14 },
  tierHead: { display: "flex", alignItems: "center", gap: 12, padding: "10px 10px 6px", fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: "#605e5c", flexWrap: "wrap" },
  // The parent term's name above its children. Deliberately NOT uppercase like `tierHead`: this is a
  // real term label ("Group Legal, Risk ＆ Compliance"), and shouting it loses the casing the client
  // authored — and would sit level with the tier heading instead of under it.
  parentHead: { display: "flex", alignItems: "baseline", gap: 8, padding: "12px 10px 4px", fontSize: 13, fontWeight: 600, color: "#242424", flexWrap: "wrap", borderTop: "1px solid #f0f0f0" },
  parentCount: { fontSize: 11.5, fontWeight: 400, color: "#8a8886" },
  parentGroup: { paddingLeft: 10, borderLeft: "2px solid #eef2f5" },
  // Spans every track: a collision message names another term and its parent, and truncating it into
  // one column would hide the half that says which other row to look at.
  problem: { fontSize: 11, lineHeight: 1.5, marginTop: 3, gridColumn: "1 / -1" },
  hint: { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  pathHint: { fontSize: 11, color: "#8a8886", fontFamily: "Consolas, monospace", overflowWrap: "break-word" },
  /* The asterisk beside every term. `#a4262c` is this project's danger red, the same one the field
     errors and the attention banners use, so "required" reads the same everywhere. */
  req: { color: "#a4262c", fontWeight: 700 },
};

export interface AbbreviationManagerProps {
  context: WebPartContext;
  siteUrl: string;
  onDirtyChange?: (dirty: boolean) => void;
  /**
   * Hands this screen's own `save()` to the host, so a guided flow's **Next** can save before it
   * advances (client, 2026-09-06 — the Save button is gone).
   *
   * ⚠ REGISTERED, NOT RE-IMPLEMENTED. The host must call THE SAME function the screen would have
   * called: it validates, writes, re-baselines the rows and reports into the panel the admin is
   * looking at. A second write path in the runner would be a second definition of what saving means,
   * and the drifting copy would be the one nobody tests.
   *
   * Called with `undefined` on unmount, so a host holding the reference cannot save a screen that is
   * no longer there.
   */
  registerSave?: (save: (() => Promise<boolean>) | undefined) => void;
  /**
   * How many terms currently have no folder code, or `undefined` when that is not knowable.
   *
   * Reported upward because the guided flow gates its Next button on it and CANNOT afford to work it out
   * itself — the count needs a walk of the whole term tree (~115 requests for GHO). This screen has
   * already paid for that walk, so the number is free here and unaffordable anywhere else. Without it
   * `abbreviationsMissing` stayed `undefined`, unknown never gates, and Next sat enabled on a screen full
   * of "no folder will be created" warnings (client, 2026-08-17).
   *
   * `undefined` means NOT KNOWABLE — no segment chosen, the tree still loading, or a term set with no
   * terms — and must never be reported as 0, which would read as "all done".
   */
  onMissingChange?: (missing: number | undefined) => void;
  /**
   * True while a read is IN FLIGHT — the segment list or the term tree.
   *
   * Separate from the count, because the count is `undefined` for the whole read and a host cannot
   * tell "still reading" from "read and failed" out of one value. Only the first should hold a Next
   * button: it clears itself in seconds, while holding on a failure strands the admin for good.
   */
  onLoadingChange?: (loading: boolean) => void;
  /**
   * Pre-selects the segment picker below, so a guided flow that already asked which segment does not
   * ask a second time. Same pattern as `SubtreeMigrator`'s `initialSegmentKey` — matches on the mode
   * row's `Title`, the key both screens build their options from. Found missing live 2026-08-26: the
   * "Rename or re-code a folder" flow already names the segment in its own header, then this screen
   * opened on a blank "Select a segment..." anyway.
   */
  initialSegmentKey?: string;
}

export default function AbbreviationManager({
  context,
  siteUrl,
  onDirtyChange,
  registerSave,
  onMissingChange,
  onLoadingChange,
  initialSegmentKey,
}: AbbreviationManagerProps): React.ReactElement {
  const [segments, setSegments] = useState<SegmentOption[]>([]);
  const [chosen, setChosen] = useState<string>("");
  const [rows, setRows] = useState<AbbrevRowDraft[]>([]);
  const [loading, setLoading] = useState(true);
  const [treeLoading, setTreeLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  /**
   * The last outcome, shown as a floating Toast rather than a banner at the top of the page — this
   * screen is 70-odd rows long, so the top is off-screen when Save is pressed and nothing appears to
   * happen (client, 2026-08-17).
   *
   * `warn` marks a HALF-success: the rows were written and the audit row was not. It matters because a
   * success auto-dismisses and a warning does not — a partial outcome that clears itself would be read
   * as a clean save.
   */
  const [result, setResult] = useState<{ ok: boolean; warn?: boolean; text: string } | undefined>(
    undefined,
  );
  /** Existing list item ids by term GUID, so a save updates rather than duplicating. */
  const [itemIds, setItemIds] = useState<Record<string, number>>({});

  const segment = (): SegmentOption | undefined => segments.filter((x) => x.key === chosen)[0];

  const configList = (): string => encodeURIComponent(cachedListTitle(LIST_SUFFIX.config));
  const abbrevList = (): string => encodeURIComponent(abbrevListTitle());

  /* ── Segments ──────────────────────────────────────────────────────────────── */

  useEffect(() => {
    const load = async (): Promise<void> => {
      // NAMES FIRST. Every read on this page goes through cachedListTitle/abbrevListTitle, and an
      // unprimed cache resolves to the LEGACY `DMS …` titles — which 404 on a renamed site and
      // surface as "the configuration list returned HTTP 404", i.e. as a MISSING list rather than a
      // list called something else. Mounted inside Folder Administration this raced the parent's own
      // priming and won only when an earlier read had already warmed the cache, so the failure came
      // and went with page-load timing. Priming is idempotent and cached; a FAILURE must not stop the
      // read, because the legacy title is still correct on a site that was never renamed.
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${configList()}')/items` +
          `?$select=Title,ModeLabel,TermSetGuid,StagingFolder,Levels&$filter=ConfigType eq 'mode'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) throw new Error(`the configuration list returned HTTP ${res.status}`);
      const raw = ((await res.json()).value ?? []) as Array<{
        Title?: string;
        ModeLabel?: string;
        TermSetGuid?: string;
        StagingFolder?: string;
        Levels?: string;
      }>;

      /* ⚠ ITS OWN REQUEST, AND ITS FAILURE CHANGES NOTHING.
         `PendingLevels` is created ON DEMAND by the Folder levels screen the first time a structure
         change is staged, so on a site where that has never happened the column DOES NOT EXIST — and
         one unknown name fails the WHOLE `$select` with HTTP 400 (gotcha #11). Folded into the read
         above, this screen would report "could not read the segments" on a site that is otherwise
         perfectly provisioned. `FolderAdmin` paid exactly that price on 2026-08-18 and both
         `StructureManager` and `SubtreeMigrator` split the read for the same reason.

         A failure here leaves every segment on its LIVE chain, which is what this screen did before
         today — degraded, never broken. */
      const pendingBy: Record<string, string> = {};
      try {
        const pres: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${configList()}')/items` +
            `?$select=Title,PendingLevels&$filter=ConfigType eq 'mode'&$top=200`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (pres.ok) {
          (
            ((await pres.json()).value ?? []) as Array<{ Title?: string; PendingLevels?: string }>
          ).forEach((r) => {
            const staged = (r.PendingLevels ?? "").trim();
            if (staged) pendingBy[(r.Title ?? "").trim()] = staged;
          });
        }
      } catch {
        /* left on the live chain — see above */
      }

      const opts = raw
        .filter((r) => (r.TermSetGuid ?? "").trim() !== "")
        .map((r) => {
          const staged = pendingBy[(r.Title ?? "").trim()] ?? "";
          const chain: Level[] = parseLevels(staged || (r.Levels ?? ""));
          const { permissioned, onDemand } = splitChain(chain);
          return {
            key: (r.Title ?? "").trim(),
            label: (r.ModeLabel ?? r.Title ?? "").trim(),
            termSetGuid: (r.TermSetGuid ?? "").trim(),
            stagingFolder: (r.StagingFolder ?? "").trim(),
            // A row predating the Levels schema still needs codes for its terms, so fall back to
            // the pilot's shared chain rather than rendering an empty page.
            levelNames:
              permissioned.length > 0 ? permissioned.map((l) => l.label) : ["Department", "Unit"],
            belowNames: onDemand.map((l) => l.label),
            below: onDemand,
          };
        })
        .sort((a, b) => a.label.localeCompare(b.label));
      setSegments(opts);
      // ⚠ MUST decide the pre-select HERE, in the SAME pass that sets `segments` — not in a separate
      // effect keyed on `chosen !== ""`. `setSegments` and this selection both fire from one load, so
      // a second effect guarding on "has anything already been chosen" cannot tell "the admin picked
      // one" apart from "this effect already auto-picked the first one a moment ago" — it always loses
      // that race, because by the time it runs, `chosen` is no longer "". Found live 2026-08-26: a flow
      // that already named the segment in its header still opened this screen on a blank picker.
      const preselect =
        initialSegmentKey && opts.filter((o) => o.key === initialSegmentKey).length > 0
          ? initialSegmentKey
          : opts.length > 0
          ? opts[0].key
          : undefined;
      if (preselect) setChosen(preselect);
      setLoading(false);
    };
    load().catch((e) => {
      setResult({ ok: false, text: `Could not read the segments — ${(e as Error).message}` });
      setLoading(false);
    });
  }, []);

  /* ── Tree + existing codes for the chosen segment ───────────────────────────── */

  useEffect(() => {
    const seg = segment();
    if (!seg) return undefined;
    let cancelled = false;
    const load = async (): Promise<void> => {
      setTreeLoading(true);
      setResult(undefined);

      /* Walk as deep as there are permissioned levels, PLUS one more when a below-Unit level is
         named by codes.

         ⚠ EVERY BELOW-UNIT LEVEL EXCEPT YEAR AND DOCUMENT TYPE. Those two are named from the term
         label — there is no useful abbreviation for `2024` — and listing their terms would inflate
         `missing`, which HOLDS the guided flow's Next button: the screen would demand codes nobody
         needs before an admin could carry on. `isAbbreviatedLevel` is the one definition, shared with
         the upload form and the migrator, so a level cannot be named one way and listed the other.

         ⚠ ONE EXTRA LEVEL, never more. `allowedTierPositions` permits exactly one per-unit level and
         the contiguity rule puts it directly under the deepest permissioned tier, so its terms are
         always one step below the leaf. */
      const codedBelow = (seg.below ?? []).filter(isAbbreviatedLevel);
      const codedPerUnit = codedBelow.filter((l) => !(l.termSet ?? "").trim())[0];
      const maxDepth = seg.levelNames.length + (codedPerUnit ? 1 : 0);
      const nodes: TermNode[] = [];
      const walk = async (parentId: string, depth: number): Promise<void> => {
        if (depth > maxDepth) return;
        const url = parentId
          ? `${siteUrl}/_api/v2.1/termStore/sets/${seg.termSetGuid}/terms/${parentId}/children?$select=id,labels`
          : `${siteUrl}/_api/v2.1/termStore/sets/${seg.termSetGuid}/children?$select=id,labels`;
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json" } },
        );
        if (!res.ok) throw new Error(`the term store returned HTTP ${res.status}`);
        const kids = (
          ((await res.json()).value ?? []) as Array<{
            id?: string;
            labels?: Array<{ name?: string }>;
          }>
        ).map((t) => ({
          id: t.id ?? "",
          label: (t.labels ?? [])[0]?.name ?? "",
          parentId,
          depth,
        }));
        for (const k of kids) {
          if (!k.id) continue;
          nodes.push(k);
          await walk(k.id, depth + 1);
        }
      };
      await walk("", 1);

      /* Coded SHARED levels draw from their own term set, not the segment's, so each needs its own
         read — this screen has only ever walked `seg.termSetGuid`.

         ⚠ THE PARENT KEY IS SYNTHETIC, AND WITHOUT IT THE COLLISION CHECK IS WRONG. Sibling
         uniqueness groups on `parentGuid`, and a top-level term carries "" — so two coded shared
         sets would be judged siblings OF EACH OTHER, and of the segment's own departments. `ARC` on
         one set and `ARC` on another are not siblings and must not clash; `ARC` twice inside ONE set
         must. `findCollisions` treats the key as opaque, which is what makes this legitimate.

         ⚠ ONE LEVEL DEEP, because a shared set is REQUIRED to be flat (1.0.491.0 refuses a nested
         one outright), so its top terms are the whole set.

         ⚠ AN UNREADABLE SET THROWS rather than contributing no rows. Omitting them silently would
         make this screen report every code filled in while terms sat unlisted — and that count is
         what releases the flow's Next button. "Could not read" is the honest answer. */
      const sharedRows: Array<{
        id: string;
        label: string;
        level: string;
        parentKey: string;
      }> = [];
      for (const lvl of codedBelow) {
        const set = (lvl.termSet ?? "").trim();
        if (!set) continue;
        const sres: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/v2.1/termStore/sets/${set}/children?$select=id,labels`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json" } },
        );
        if (!sres.ok) {
          throw new Error(
            `the term store returned HTTP ${sres.status} for the "${lvl.label}" level's own term set`,
          );
        }
        (
          ((await sres.json()).value ?? []) as Array<{
            id?: string;
            labels?: Array<{ name?: string }>;
          }>
        ).forEach((t) => {
          if (!t.id) return;
          sharedRows.push({
            id: t.id,
            label: (t.labels ?? [])[0]?.name ?? "",
            level: lvl.label,
            parentKey: `${SHARED_SET_KEY}${set.toLowerCase()}`,
          });
        });
      }

      // Existing rows, with their item ids so a save updates in place.
      const existing: Record<string, { id: number; abbreviation: string }> = {};
      const cur: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${abbrevList()}')/items?$select=Id,TermGuid,Abbreviation&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!cur.ok) throw new Error(`${abbrevListTitle()} returned HTTP ${cur.status}`);
      (
        ((await cur.json()).value ?? []) as Array<{
          Id: number;
          TermGuid?: string;
          Abbreviation?: string;
        }>
      ).forEach((r) => {
        const key = (r.TermGuid ?? "").trim().toLowerCase();
        if (key) existing[key] = { id: r.Id, abbreviation: (r.Abbreviation ?? "").trim() };
      });

      if (cancelled) return;
      const ids: Record<string, number> = {};
      const drafts: AbbrevRowDraft[] = nodes.map((n) => {
        const found = existing[n.id.toLowerCase()];
        if (found) ids[n.id] = found.id;
        return {
          termGuid: n.id,
          label: n.label,
          // Past the permissioned depth the only thing down there is the coded per-unit level.
          level:
            seg.levelNames[n.depth - 1] ??
            (codedPerUnit ? codedPerUnit.label : `Level ${n.depth}`),
          parentGuid: n.parentId,
          abbreviation: found ? found.abbreviation : "",
          original: found ? found.abbreviation : "",
        };
      }).concat(
        sharedRows.map((r) => {
          const found = existing[r.id.toLowerCase()];
          if (found) ids[r.id] = found.id;
          return {
            termGuid: r.id,
            label: r.label,
            level: r.level,
            parentGuid: r.parentKey,
            abbreviation: found ? found.abbreviation : "",
            original: found ? found.abbreviation : "",
          };
        }),
      );
      setItemIds(ids);
      setRows(drafts);
      setTreeLoading(false);
    };
    load().catch((e) => {
      if (cancelled) return;
      setRows([]);
      setResult({ ok: false, text: `Could not load "${seg.label}" — ${(e as Error).message}` });
      setTreeLoading(false);
    });
    return () => {
      cancelled = true;
    };
  }, [chosen, segments.length]);

  /* ── Dirty ─────────────────────────────────────────────────────────────────── */

  const pending = changedRows(rows);
  const dirty = pending.length > 0;

  useEffect(() => {
    if (onDirtyChange) onDirtyChange(dirty);
  }, [dirty]);

  useEffect(() => {
    if (!dirty) return undefined;
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = "";
      return "";
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [dirty]);

  /* ── Editing ───────────────────────────────────────────────────────────────── */

  const problems: Record<string, RowProblem> = validateRows(rows);
  const blocked = hasBlockingProblem(problems);
  const renames = renamingRows(rows);

  const setCode = (termGuid: string, value: string): void =>
    setRows(rows.map((r) => (r.termGuid === termGuid ? { ...r, abbreviation: value } : r)));

  /* ⚠ `sameAsName` AND `sameAsNameForLevel` ARE DELETED (client, 2026-09-06: *"remove those
     buttons. force client to type"*). They filled a row, or every empty row of a level, from the
     term's own name.

     Deleted rather than parked: an unused function is a lint warning, and this project has learned
     twice that something kept alive doing exactly the withdrawn thing gets wired back up. Both were
     one line — `{ ...r, abbreviation: r.label }` — so restoring either is trivial if the client
     changes their mind.

     ⚠ WHAT WENT WITH THEM IS NOT ONLY CONVENIENCE. "Same as term name" stored the literal folder
     name, which meant skip-not-guess, sibling uniqueness and rename-on-change all applied to it for
     free. Typing the same value by hand gets the same treatment, so nothing is weaker — it is
     slower, which is what was asked for. */

  /* ── Save ──────────────────────────────────────────────────────────────────── */

  /**
   * Write the changed rows. Answers whether the step may move on.
   *
   * ⚠ IT RETURNS A VERDICT NOW, because the flow's Next calls it (see `registerSave`) and must NOT
   * advance on a failure — the codes are in component state and nowhere else, so a step change
   * would lose them silently. `false` for a refusal as well as an error: a sibling collision is not
   * an exception, and reconciliation would abort on it three steps later.
   */
  const save = async (): Promise<boolean> => {
    const seg = segment();
    // Nothing to save is SUCCESS — an admin who changed nothing must still be able to walk on.
    if (!seg) return false;
    if (blocked) return false;
    if (!dirty) return true;
    setBusy(true);
    setResult(undefined);
    try {
      const digestRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/contextinfo`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      const digest = (await digestRes.json()).FormDigestValue as string;

      let written = 0;
      for (const r of pending) {
        const id = itemIds[r.termGuid];
        const base = `${siteUrl}/_api/web/lists/getbytitle('${abbrevList()}')/items`;
        const res: SPHttpClientResponse = await context.spHttpClient.post(
          id ? `${base}(${id})` : base,
          SPHttpClient.configurations.v1,
          {
            // Built as an explicit Record<string,string>: an inline `id ? {...} : {}` spread widens
            // the header values to `string | undefined`, which SPHttpClient's HeadersInit rejects.
            headers: ((): Record<string, string> => {
              const h: Record<string, string> = {
                Accept: "application/json;odata=nometadata",
                "Content-Type": "application/json;odata=nometadata",
                "X-RequestDigest": digest,
              };
              // An existing row is UPDATED, never re-created — a second row for the same term GUID
              // would leave reconciliation picking between two codes for one folder.
              if (id) {
                h["X-HTTP-Method"] = "MERGE";
                h["IF-MATCH"] = "*";
              }
              return h;
            })(),
            body: JSON.stringify({
              // The LIVE term label, not anything typed here. Title is what reconciliation writes to
              // each folder's Full Name column, and it is also the only thing that survives a term
              // being deleted and recreated — which is how an orphaned row is ever repaired.
              Title: r.label,
              TermGuid: r.termGuid,
              Level: r.level,
              Abbreviation: r.abbreviation.trim(),
            }),
          },
        );
        if (!res.ok) {
          const body = await res.text().catch(() => "");
          throw new Error(
            `"${r.label}" could not be saved (HTTP ${res.status}). ${body.slice(0, 160)} ` +
              `${written} row(s) were saved before this one.`,
          );
        }
        written++;
      }

      const renamed = renames.length;

      // Recorded BEFORE re-baselining, because `original` is about to be overwritten and the
      // old → new pair is the whole value of the record: it is the only place that says which folder
      // the next reconciliation will rename, and why.
      const auditOk = await writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.abbreviationChanged,
        source: "AbbreviationManager",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        segment: seg.label,
        summary:
          `Abbreviations saved for ${seg.label} — ${written} changed` +
          (renamed > 0 ? `, ${renamed} rename${renamed === 1 ? "" : "s"}` : ""),
        details: pending
          .map(
            (r) =>
              // `original` is optional on the draft, and an absent one reads the same as an empty
              // one here — both mean "there was no stored code", which is what a reader needs.
              `${r.level} "${r.label}": ` +
              `${(r.original ?? "").trim() === "" ? "(blank)" : (r.original ?? "").trim()} → ` +
              `${r.abbreviation.trim() === "" ? "(blank)" : r.abbreviation.trim()}`,
          )
          .concat([
            renamed > 0
              ? "No folder has changed yet. The next reconciliation will RENAME the folders whose code changed."
              : "No folder has changed yet. Reconciliation creates the folders.",
          ]),
      });

      // Re-baseline in place. A reload would repeat the tree walk for no gain, and this keeps the
      // rows on screen exactly as saved.
      setRows(rows.map((r) => ({ ...r, original: r.abbreviation.trim() })));
      setResult({
        ok: true,
        warn: !auditOk,
        text:
          `Saved ${written} abbreviation${written === 1 ? "" : "s"} for ${seg.label}. ` +
          `No folder has changed yet — run Folder Reconciliation to create or rename them.` +
          (renamed > 0
            ? ` ${renamed} of these replaced an existing code, so reconciliation will RENAME ${
                renamed === 1 ? "that folder" : "those folders"
              } in both libraries. Documents, permissions and approval status are kept.`
            : "") +
          // Said on the panel already being read, rather than as a second banner. The save itself
          // succeeded and must not be made to look otherwise.
          (auditOk ? "" : " (This change could not be recorded in the audit log.)"),
      });
      return true;
    } catch (e) {
      setResult({ ok: false, text: (e as Error).message });
      return false;
    } finally {
      setBusy(false);
    }
  };

  /* ⚠ RE-REGISTERED ON EVERY RENDER, AND AN EMPTY DEP ARRAY WOULD BE THE BUG. `save` closes over
     `rows`, so a reference captured once would go on writing the codes as they were when the screen
     mounted — the stale-closure trap that has already cost this project a reconciliation run's
     permissions and an evening of diagnosis. Registering each render keeps the host holding the
     CURRENT closure; the cleanup clears it so a host cannot save a screen that is no longer there. */
  useEffect(() => {
    if (registerSave) registerSave(save);
    return () => {
      if (registerSave) registerSave(undefined);
    };
  });

  const seg = segment();
  const missing = rows.filter((r) => r.abbreviation.trim() === "").length;

  /* Report the count up to whatever hosts this screen — the guided flow gates Next on it.
     THE THREE UNKNOWN CASES ARE THE POINT: still loading, no segment chosen, or a term set with no terms
     at all. Each would compute `missing === 0` from an empty `rows`, and 0 means "all done" to a caller,
     which is the one answer that must never be inferred from an absence. `rows.length === 0` covers the
     no-terms case explicitly rather than relying on the loading flag having cleared.

     Reports the LIVE value, not the saved one, so typing a code lifts the gate immediately and clearing
     one puts it back — the count on screen and the gate can never disagree.

     ⚠ IT MUST STAY ABOVE THE `if (loading)` RETURN BELOW. Placed after it, the hook was skipped on the
     first render and called on the next, which is a hook-count change — React throws and the whole page
     renders BLANK, with the flow's heading and rail gone too (2026-08-17). Hooks are unconditional or
     they are a crash; an early return is exactly the kind of line that hides one. */
  useEffect(() => {
    if (!onMissingChange) return;
    const knowable = !loading && !treeLoading && seg !== undefined && rows.length > 0;
    onMissingChange(knowable ? missing : undefined);
  }, [onMissingChange, loading, treeLoading, seg, rows.length, missing]);

  /* And whether a read is IN FLIGHT, which is not the same fact.
     Both are reported because neither implies the other: the count is `undefined` while loading AND
     when the read failed, and only the first should hold the flow's Next button. Under the same
     unconditional-hook rule as above — never move this below an early return. */
  useEffect(() => {
    if (!onLoadingChange) return;
    onLoadingChange(loading || treeLoading);
  }, [onLoadingChange, loading, treeLoading]);

  /* ── Render ────────────────────────────────────────────────────────────────── */

  if (loading) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading segments&hellip;</p>;

  return (
    <div>
      {/* A FLOATING toast, not a banner in the flow of the page. Save sits at the bottom of 70-odd rows,
          so a message at the top is off-screen exactly when it is written and pressing Save reads as
          doing nothing (client, 2026-08-17).

          Three kinds, because the dismissal differs: a success clears itself, while a FAILURE and a
          half-success (rows written, audit row not) stay until closed. A toast that clears itself after an
          error is worse than the banner it replaced — it can be missed entirely rather than merely be out
          of view. */}
      {result && (
        <Toast
          kind={(!result.ok ? "err" : result.warn ? "warn" : "ok") as ToastKind}
          text={result.text}
          onDismiss={() => setResult(undefined)}
        />
      )}

      {/* ⚠ A PLAIN LINE WHEN A FLOW ALREADY CHOSE THE SEGMENT, not a pre-selected dropdown. Reported
          on the first live test of the new step (2026-09-09): the structure flow renders its own
          segment switcher, so this step showed TWO "Segment" dropdowns stacked, both naming the same
          thing. Neither had been on one screen before, because this flow only gained an abbreviations
          step today.

          Exactly the treatment the migrate screen already carries (1.0.456.0) for the same complaint
          — client, 2026-09-06: *"You haven't remove the select... just put a message indicator is
          enough."* Gated on the HOST having supplied a segment rather than on which flow this is, so
          the standalone tab keeps its real picker. */}
      {initialSegmentKey ? (
        <div style={{ marginBottom: 16 }}>
          <label style={s.label}>Segment</label>
          <div style={{ fontSize: 13, fontWeight: 600, color: "#242424" }}>
            {segment()?.label ?? initialSegmentKey}
          </div>
        </div>
      ) : (
      <div style={{ marginBottom: 16 }}>
        <label style={s.label}>Segment</label>
        <select
          style={s.select}
          value={chosen}
          onChange={(e) => {
            // Switching segment discards unsaved edits, so refuse rather than confirm — Save is
            // right there, and losing the work behind something that looks like navigation is the
            // same silent discard the other structure screens guard against.
            if (dirty) {
              setResult({
                ok: false,
                text: "Save your changes first — switching segment would lose them.",
              });
              return;
            }
            setChosen(e.target.value);
          }}
        >
          {segments.map((o) => (
            <option key={o.key} value={o.key}>
              {o.label}
            </option>
          ))}
        </select>
        {/* The "top folder is set on the segment, not here" hint came off on the client's
            instruction (2026-09-06), along with the below-Unit note and the missing-code banner.
            The screen now says one thing: every field is required. */}
      </div>
      )}


      {treeLoading ? (
        <p style={{ fontSize: 13, color: "#605e5c" }}>Reading the term store&hellip;</p>
      ) : // NO SEGMENT is tested BEFORE no terms. Reversed, a page that could not read the segment
      // list at all blamed the term store — "this segment's term set has no terms yet" under an
      // empty dropdown, sending an admin to fix a term set that was never the problem.
      !seg ? null : rows.length === 0 ? (
        <div style={{ ...s.msg, ...s.warn }}>
          This segment&apos;s term set has no terms yet. Add its structure to the term store first —
          reconciliation builds one folder level per term level, so there is nothing to name.
        </div>
      ) : (
        <>

          {/* SCROLLS (client, 2026-08-18): 7 departments plus 60 units is 67 rows, so Save sat far
              below the fold and the warning banner above scrolled out of sight while an admin worked
              through the list.

              Everything that must stay visible is OUTSIDE this box — the collision banner and the
              rename warning — so the only thing scrolling is the rows themselves. (The Save button
              it used to keep out of the box is gone; Next saves.) Safe to cap unconditionally, unlike the group list on Folder Access: these
              rows hold text boxes and buttons, and nothing absolutely positioned that a scroll
              container could clip.

              `overflowX: hidden` because the row is a flex layout that already fits; without it a
              long term label can raise a horizontal scrollbar under every row. */}
          <div style={s.scrollBox}>
          {/* ⚠ CODED BELOW-UNIT LEVELS ARE APPENDED, and leaving them out is a DEADLOCK rather than
              a missing section. Their rows are built and counted as missing, so `abbreviationsMissing`
              is non-zero and the flow's Next button is HELD — while the rows themselves render
              nowhere, because this list only ever held the permissioned levels. The admin would be
              told codes are outstanding with nothing on screen to fill in.

              Derived from the same `isAbbreviatedLevel` rule the walk uses, so a level can never be walked
              into rows without also being rendered. */}
          {seg.levelNames
            .concat((seg.below ?? []).filter(isAbbreviatedLevel).map((l) => l.label))
            .map((levelName) => {
            const levelRows = rows.filter((r) => r.level === levelName);
            if (levelRows.length === 0) return null;
            return (
              <div key={levelName} style={{ marginBottom: 18 }}>
                <div style={s.tierHead}>
                  <span>
                    {levelName} ({levelRows.length})
                  </span>

                </div>
                {/* GROUPED UNDER THE PARENT TERM (client, 2026-08-17). A flat `UNIT (63)` list is
                    alphabetical across the whole segment, so `Tax` sat between `SDGI` and `Treasury`
                    with nothing saying which department any of them belonged to. It also makes the rule
                    visible: codes must be unique among SIBLINGS, so two departments may each hold a
                    `TAX` unit — which looks like a duplicate in a flat list and is perfectly legal
                    here. */}
                {groupRowsByParent(levelRows, rows).map((g) => {
                  /* ⚠ A SHARED LEVEL'S TERMS SIT AT THE TOP OF THEIR OWN SET, so they carry the
                     synthetic grouping key `set:<guid>` — which is what stops two different sets
                     being judged siblings of each other AND of the segment's own departments. There
                     is no term behind that key, so the parent header fell through to "Parent term
                     not found" and read as an error (seen live 2026-09-09), with a count that
                     pluralised the level name into "2 minasmas archive 2s".
                     A flat set has exactly ONE group, so the header adds nothing anyway — the level
                     heading above already names it. Rendered flat, like the top-level case. */
                  const flatSet = g.parentGuid.indexOf(SHARED_SET_KEY) === 0;
                  const grouped = g.parentGuid !== "" && !flatSet;
                  return (
                  <div key={g.parentGuid || "__top"} style={grouped ? s.parentGroup : undefined}>
                    {grouped ? (
                      <div style={s.parentHead}>
                        {/* An UNRESOLVED parent is labelled, never hidden: a row with no code still
                            needs one, and dropping it is the one failure this page exists to stop. */}
                        <span>{g.parentLabel || "Parent term not found"}</span>
                        <span style={s.parentCount}>
                          {g.rows.length} {levelName.toLowerCase()}
                          {g.rows.length === 1 ? "" : "s"}

                        </span>
                      </div>
                    ) : undefined}
                    {g.rows.map((r) => {
                  const p = problems[r.termGuid];
                  return (
                    <div key={r.termGuid} style={s.row}>
                      {/* ⚠ THE ASTERISK IS NOT DECORATION — every code is REQUIRED now (client,
                          2026-09-06: *"ensure it is a mandatory to fill in"*), and Next is held
                          while any is blank. It marks the term, not the box, because the box has no
                          label of its own. */}
                      <span style={{ fontSize: 13 }}>
                        {r.label} <span style={s.req}>*</span>
                      </span>
                      <input
                        style={{ ...s.input, ...(p && p.error ? s.inputBad : {}) }}
                        value={r.abbreviation}
                        placeholder="code"
                        /* ⚠ 20 CHARACTERS (client's cap). A folder code is a SHORT stand-in for a
                           term — `GHO`, `CORU` — and the whole point of the list is that a path
                           stays readable. `maxLength` refuses the 21st keystroke rather than
                           reporting the problem afterwards, which is the right shape for a limit
                           somebody meets while typing. */
                        maxLength={20}
                        /* ⚠ STRIPPED AS IT IS TYPED (client, 2026-09-09: *"did you prevent any
                           illegal chracter to not enter in the term abbrevition input?"* — it did
                           not). `folderNameFor` sanitizes when the folder is NAMED, so typing `MA/D`
                           stored `MA/D` in this list and created a folder called `MAD`: no error, no
                           warning, and the list silently stopped saying what the folder is called —
                           which is the one thing it exists to say.
                           ⚠ `stripFolderChars`, never `sanitizeFolderSegment`: that one also trims,
                           so a space could never be typed and a two-word code would be unreachable. */
                        onChange={(e) => setCode(r.termGuid, stripFolderChars(e.target.value))}
                      />
                      {/* ⚠ "Same as term name" AND THE `/code` PREVIEW BOTH REMOVED (client,
                          2026-09-06: *"remove those buttons. force client to type"*). The rename
                          marker stays — it is the ONLY place an admin can see that this edit will
                          move a live folder, and it appears BEFORE they commit rather than in the
                          save summary afterwards. */}
                      {p && p.note ? (
                        <span style={{ fontSize: 11, color: "#7a4f00", fontWeight: 600 }}>
                          ✎ {p.note}
                        </span>
                      ) : (
                        <span />
                      )}
                      {p && (p.error || p.warn) && (
                        <div style={{ ...s.problem, color: p.error ? "#a4262c" : "#7a4f00" }}>
                          {p.error ?? p.warn}
                        </div>
                      )}
                    </div>
                  );
                })}
                  </div>
                  );
                })}
              </div>
            );
          })}
          </div>

          {/* ⚠ THE SAVE BUTTON IS GONE, AND NEXT SAVES INSTEAD (client, 2026-09-06: *"Remove Save
              button, after System Admin keyed in the abbreviation, click Next for clear instruction
              to proceed to the next page"*).

              Two controls for one intention was the problem: an admin who typed the codes and
              pressed Next was told to press Save first — a gate explaining a button a foot away,
              which reads as the page arguing with itself.

              ⚠ SAVING IS STILL A REAL STEP, NOT AN AUTOSAVE. `registerSave` hands this component's
              own `save()` up to the flow runner, which awaits it and advances ONLY if it succeeded.
              A failed write therefore holds the step with its own message rather than moving on and
              losing the codes — which is the outcome the removed warning existed to prevent.

              ⚠ AND A COLLISION STILL BLOCKS. `save()` refuses while two siblings would share a
              name, so Next refuses too: reconciliation would abort on it anyway, and it must not be
              possible to walk past a state that stops the run three steps later. */}
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 8 }}>
            {busy && <span style={{ fontSize: 12, color: "#605e5c" }}>Saving&hellip;</span>}
            {blocked && (
              <span style={{ fontSize: 12, color: "#a4262c" }}>
                Two folders would share a name — fix the rows in red first.
              </span>
            )}
            {!blocked && dirty && renames.length > 0 && (
              <span style={{ fontSize: 12, color: "#7a4f00" }}>
                {renames.length} existing code{renames.length === 1 ? "" : "s"} changed — live folders
                will be renamed at the next reconciliation.
              </span>
            )}
          </div>
        </>
      )}
    </div>
  );
}

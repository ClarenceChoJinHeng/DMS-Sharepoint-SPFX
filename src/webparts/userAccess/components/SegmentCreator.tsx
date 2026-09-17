import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, sanitizeFolderSegment } from "../../../shared/formModel";
import { effectiveOnDemandTiers } from "../../../shared/folderChain";
import { EVENT } from "../../../shared/auditLog";
// `libraryTargets()` (used inside `shared/spSegmentRecode.ts`'s `retireLibraries`/`archiveTargets`)
// carries BOTH names per library. Use `.urlSegment` for anything that builds a PATH and `.title`
// only for display: the two differ ("Approval Document" vs "/ApprovalDocument", "Documents" vs
// "/Shared Documents") and a title in a URL fails SILENTLY — gotcha #12.
import { allLibraryTitles, archiveAvailable, cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { ensureColumn } from "../../../shared/spColumns";
import { openInNewTab } from "../../../shared/newTab";
import {
  SegmentCounts,
  canDeleteArchive,
  canOfferFolderDelete,
  confirmationMatches,
  deletionSummary,
  needsTypedConfirmation,
  survivorLines,
  unknownCounts,
} from "../../../shared/segmentDeletion";
import {
  canOfferAnyRecode,
  folderRecodeConflict,
  folderRecodeIsNoOp,
  recodeRefusalReason,
  recodeSummaryLive,
} from "../../../shared/segmentRecode";
import {
  archiveTargets,
  countArchiveDocuments,
  countArchiveDocumentsForCode,
  countSegmentDocuments,
  deleteItem,
  loadSegmentAbbreviationRowIds,
  loadSegmentGroupMapRows,
  performLiveSegmentRecode,
  recycleFolder,
  retireLibraries,
} from "../../../shared/spSegmentRecode";
import {
  buildSegmentLevels,
  columnNameFor,
  columnsForDraft,
  depthVerdict,
  ExistingSegment,
  fieldConflicts,
  isGuid,
  MAX_PERMISSIONED_TIERS,
  modeKeyFor,
  NewSegmentDraft,
  nextSortOrder,
  normalizeGuid,
  validateNewSegment,
  requiredFieldErrors,
} from "../../../shared/newSegment";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

/**
 * Add a whole business segment — slice B of `2026-08-10-structure-manager-ui-design.md`,
 * specced in `2026-08-12-add-segment-design.md`.
 *
 * Slice A edits the levels of a segment that already exists. This creates one: the DMS Config
 * `mode` row plus the tier columns in both libraries. Eight of the twelve intended segments are
 * still unbuilt and they do NOT share the head-office shape — Upstream Ops needs
 * Region → Estate/Mill, I&T needs a single tier — so the admin NAMES the permissioned tiers and
 * this creates their columns. A fixed Department/Unit prefix could not onboard them.
 *
 * It deliberately does NOT create the term set, the abbreviations, the groups, or any folder.
 * That is why it ends on a checklist rather than a success message: a mode row is one of six
 * things a segment needs, and the abbreviation step silently creates nothing when it is missed.
 */


/** How many term-store requests the depth walk may spend before giving up. See measureDepth. */
const DEPTH_REQUEST_CAP = 400;
/** Concurrency for the walk — quick enough to feel instant, low enough not to invite throttling. */
const DEPTH_BATCH = 8;

type SetCheck =
  | { state: "blank" }
  | { state: "malformed" }
  | { state: "checking" }
  | { state: "found"; name: string; count: number }
  | { state: "notfound" }
  | { state: "unknown"; status: number };

/**
 * Whether a re-check is worth offering — same rule as `StructureManager.tsx`'s identical helper.
 *
 * ONLY WHERE THE VERDICT CAME FROM THE SERVER. `blank` and `malformed` are decided locally from the
 * field itself, so re-asking cannot change either — a button there would be a control that visibly
 * does nothing, which is how the ones that DO work stop being trusted.
 */
function setCheckIsRecheckable(c: SetCheck): boolean {
  return (
    c.state === "checking" || c.state === "found" ||
    c.state === "notfound" || c.state === "unknown"
  );
}

const s: Record<string, React.CSSProperties> = {
  msg: { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.5 },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { ...NOTICE_ATTENTION },
  ok: { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  card: { border: "1px solid #e1e1e1", borderRadius: 8, padding: "14px 16px", marginBottom: 12, background: "#fff" },
  label: { display: "block", fontSize: 12, fontWeight: 600, color: "#323130", margin: "14px 0 4px" },
  input: { width: "100%", boxSizing: "border-box", padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4 },
  // Spread OVER `input`, so the red border is the only difference and the two can never drift apart.
  inputBad: { border: "1px solid #d13438", background: "#fdf6f6" },
  fieldErr: { fontSize: 12, color: "#a4262c", marginTop: 4, marginBottom: 8, lineHeight: 1.45, fontWeight: 600 },
  hint: { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  btn: { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "8px 16px", fontSize: 13, cursor: "pointer" },
  ghost: { background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  off: { background: "#f3f2f1", color: "#a19f9d", border: "1px solid #e1dfdd", borderRadius: 4, padding: "8px 16px", fontSize: 13, cursor: "not-allowed" },
  iconBtn: { background: "#fff", border: "1px solid #c8c8c8", borderRadius: 4, padding: "3px 8px", fontSize: 12, cursor: "pointer", marginRight: 4 },
  danger: { background: "#fff", color: "#a4262c", border: "1px solid #e6b3b5", borderRadius: 4, padding: "4px 9px", fontSize: 12, cursor: "pointer" },
  // Term Store link + Re-check button (client QA item #57, 2026-09-13), mirroring
  // `StructureManager.tsx`'s identical pairing beside its own Term set ID field.
  // ⚠ `marginBottom: 4` MATCHES `label` BELOW — dropped when this was split into a flex row to
  // hold the "Open the Term Store" link and the re-check button beside the label, so the row's
  // bottom edge sat flush against the input underneath it (found live, 2026-09-17: "button is
  // sticking the input"). `label` is a plain block and always had this gap; this row needs the
  // same one, or the taller bordered re-check button reads as glued to the box below it.
  labelRow: { display: "flex", alignItems: "center", gap: 5, marginTop: 14, marginBottom: 4 },
  labelLink: { fontSize: 12, fontWeight: 400, marginLeft: "auto" },
  recheckBtn: { border: "1px solid #c7c7c7", background: "#fff", borderRadius: 4, cursor: "pointer", fontSize: 12, lineHeight: 1, padding: "3px 8px", color: "#0f6c3f", fontWeight: 400 },
  recheckOff: { border: "1px solid #e1dfdd", background: "#f3f2f1", borderRadius: 4, cursor: "not-allowed", fontSize: 12, lineHeight: 1, padding: "3px 8px", color: "#a19f9d", fontWeight: 400 },
  tierRow: { display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 6, background: "#fafafa", marginBottom: 6, fontSize: 13, flexWrap: "wrap" },
  tierLock: { background: "#f3f2f1", color: "#605e5c" },
  tierName: { fontWeight: 600, flex: "0 0 170px" },
  tierMeta: { fontSize: 11, color: "#8a8886", flex: "1 1 160px" },
  path: { fontSize: 12, color: "#605e5c", fontFamily: "Consolas, monospace", marginTop: 6, wordBreak: "break-all" },
  h: { fontSize: 13, fontWeight: 700, color: "#1b1b1b", margin: "0 0 8px" },
  li: { fontSize: 13, lineHeight: 1.6, marginBottom: 4 },
};

function setCheckMessage(c: SetCheck): string {
  switch (c.state) {
    case "checking":
      return "Checking…";
    case "malformed":
      return "That is not a term set ID. Copy the ID from the term store — it looks like 023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf.";
    case "notfound":
      return "No term set with that ID exists on this site. If you copied it from another site, or copied a term instead of the term set, it will not work here.";
    case "unknown":
      return `Could not check that ID right now${c.status ? ` (HTTP ${c.status})` : ""}. You can still continue, but confirm the ID is right.`;
    case "found":
      return c.count === 0
        ? `Found "${c.name}", but it has no terms yet. The segment's structure has to exist in the term store before reconciliation can build any folders.`
        : `Found "${c.name}" — ${c.count} top-level ${c.count === 1 ? "term" : "terms"}.`;
    default:
      return "";
  }
}

function setCheckStyle(c: SetCheck): React.CSSProperties {
  if (c.state === "notfound" || c.state === "malformed") return { color: "#a4262c" };
  if (c.state === "unknown" || (c.state === "found" && c.count === 0)) return { color: "#7a4f00" };
  if (c.state === "found") return { color: "#0f6c3f" };
  return {};
}

export interface SegmentCreatorProps {
  context: WebPartContext;
  siteUrl: string;
  /** Fired as the form gains or loses unsaved input, so the page can guard a tab switch. */
  onDirtyChange?: (dirty: boolean) => void;
  /**
   * Fired once a segment has actually been written, with its DERIVED key and its label.
   *
   * Exists so a host can react to the creation instead of inferring it. The guided flow could only
   * infer it by re-reading the mode rows and asking the admin which one was new — which is why the
   * form stayed open under its own success message (client, 2026-08-17). `key` is the value nobody
   * types (`mode_<slug>`), so it is also the only reliable way for a caller to select the new row.
   */
  onCreated?: (key: string, label: string) => void;
  /**
   * Whether the existing-segments list offers **Delete**. Absent means yes.
   *
   * Set false by the "Add a new segment" flow (client, 2026-08-17: *"why do we allow client to delete in
   * the Add a New Segment flow? we should ony allow them to create not delete"*). Retiring a segment has
   * its own guided flow, which moves the documents out first and asks for a typed confirmation — putting
   * the same Delete button one click from a creation form offers the destructive half with none of that.
   *
   * DEFAULTS TO SHOWN on purpose. The Retire flow's delete step mounts this very screen, so a default of
   * hidden would silently remove the button from the flow whose entire purpose is to use it — and the
   * standalone Segments tab needs it too. Only the create flow opts out.
   *
   * WARN: IT ALSO DECIDES WHETHER THE SEGMENTS LIST RENDERS AT ALL (client, 2026-09-09). This
   * used to read "the list itself STAYS either way" and that is no longer true: the list is shown
   * only where it can be acted on, which is everywhere except the create flow. The rows are still
   * READ regardless - the duplicate-name refusal depends on them.
   */
  allowDelete?: boolean;
  /**
   * Mirror of `allowDelete`, same reasoning inverted. The Retire flow mounts this exact screen for
   * step 2 ("Segments -> Delete"), and that flow has no use for the create form below the segments
   * list — client, 2026-08-26: *"Retiring a segment is just to delete, not to create."* DEFAULTS TO
   * SHOWN, so the standalone Segments tab and the "Add a new segment" flow are both untouched; only
   * Retire opts out. The segments list is what that flow exists to show, and it survives because
   * Retire leaves `allowDelete` alone - see the warning on that prop.
   */
  allowCreate?: boolean;
  /**
   * Whether the segments list offers **Re-code folder**. Absent means yes.
   *
   * Set false by the Retire flow (client, 2026-09-13: they found it sitting on the retire screen and
   * asked why — it was never a deliberate choice, just a gap: Re-code was built after Retire's
   * "hide everything except delete" wiring already existed, and nobody folded it in). Retiring and
   * recoding are different actions on the same row and must not be offered side by side on the
   * screen whose whole purpose is deleting.
   *
   * DEFAULTS TO SHOWN, matching `allowDelete`/`allowCreate` — only Retire opts out; the standalone
   * Segments tab and "Add a new segment" are unaffected.
   */
  allowRecode?: boolean;
}

export default function SegmentCreator({
  context,
  siteUrl,
  onDirtyChange,
  onCreated,
  allowDelete,
  allowCreate,
  allowRecode,
}: SegmentCreatorProps): React.ReactElement {
  const [existing, setExisting] = useState<ExistingSegment[]>([]);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [loaded, setLoaded] = useState(false);

  const [label, setLabel] = useState("");
  const [family, setFamily] = useState<"BusinessSegment" | "Project">("BusinessSegment");
  const [termSetGuid, setTermSetGuid] = useState("");
  const [stagingFolder, setStagingFolder] = useState("");
  // EMPTY, not ["Department", "Unit"] (client, 2026-08-15: "it is always showing department and Unit
  // this will confuse the client").
  //
  // The seed was not merely confusing — it was a trap the depth check could not catch. That check
  // compares the term set's depth to the NUMBER of tiers, so a 2-deep Upstream Ops set matched a
  // seeded Department/Unit perfectly: right count, wrong names. It would have created
  // `Department`/`Unit` columns for a segment whose tiers are Region and Estate/Mill, with nothing
  // failing, surfacing later as a detail panel labelled in another segment's vocabulary.
  //
  // The admin NAMES the tiers here (spec 2026-08-12, "The admin NAMES the permissioned tiers"), so a
  // default is a name nobody chose. The examples stay in the hint, where they teach without filling in.
  const [tiers, setTiers] = useState<string[]>([]);
  /* ⚠ DERIVED FROM THE SHARED CONSTANT, never a literal 2 typed here. `validateNewSegment` refuses
     on the same number, so the control and the rule cannot drift into disagreeing — a greyed Add
     button beside a form that would have accepted the level, or the reverse, is worse than either. */
  const tiersFull = tiers.length >= MAX_PERMISSIONED_TIERS;
  const [below, setBelow] = useState<Level[]>([]);
  const [newTier, setNewTier] = useState("");

  // Deletion. `delCounts === undefined` means "still counting" — distinct from a count that
  // finished and came back unknown, which is a state the dialog has to render differently.
  const [deleting, setDeleting] = useState<ExistingSegment | undefined>(undefined);
  const [delCounts, setDelCounts] = useState<SegmentCounts | undefined>(undefined);
  const [delFolders, setDelFolders] = useState(false);
  const [delTyped, setDelTyped] = useState("");
  const [delLog, setDelLog] = useState<string[]>([]);
  /**
   * Deleting a segment is up to ~1,000+ sequential row deletes (one HTTP call each, since a bulk
   * delete endpoint does not exist) — genuinely slow, not stuck. Found live 2026-08-26, retiring GHO's
   * 1157 mappings: a static "Deleting…" label with no count gives an admin nothing to judge whether it
   * is working or hung, over a run that can run for minutes. `undefined` before the row count is known.
   */
  // `label` distinguishes which step is running — the mapping cleanup and the abbreviation
  // cleanup (2026-09-11) share this one progress bar, and without a label the bar would go on
  // reading "Removing folder-access mappings" while it is actually deleting abbreviation rows.
  const [delProgress, setDelProgress] = useState<
    { done: number; total: number; label: string } | undefined
  >(undefined);

  /* Re-coding a segment's TOP FOLDER — spec 2026-09-11-segment-recode-design.md. Separate state from
     the delete dialog: a different count-based gate (documents === 0, not the delete flow's
     folder/document/mapping trio) and a different write (MERGE StagingFolder + recycle the OLD tree,
     never the mode row itself). `recodeCounts === undefined` means "still counting", same convention
     as `delCounts`. */
  const [recoding, setRecoding] = useState<ExistingSegment | undefined>(undefined);
  const [recodeCounts, setRecodeCounts] = useState<SegmentCounts | undefined>(undefined);
  const [recodeNewFolder, setRecodeNewFolder] = useState("");
  const [recodeTyped, setRecodeTyped] = useState("");
  const [recodeLog, setRecodeLog] = useState<string[]>([]);

  const [check, setCheck] = useState<SetCheck>({ state: "blank" });
  /* Bumped by the "Re-check" button (client QA item #57, 2026-09-13), mirroring
     `StructureManager.tsx`'s identical below-Unit-level pairing. It is a DEP on the check effect
     rather than a second call site for the check — one definition of "run the check", so the debounce
     and the `cancelled` guard apply to a manual re-check exactly as they do to typing. */
  const [recheck, setRecheck] = useState(0);
  const [busy, setBusy] = useState(false);
  /* Red on the required fields, and ONLY after a save has actually been refused.
     Client, 2026-09-08: *"I notice a bug when I click on the Create Segment it doesnt show the
     error message. Highlight the Term set ID and Top Folder Name red."*
     Never set on load: a blank form lit up red reads as broken, and people then stop reading
     red anywhere. Each field's marker is DERIVED from its own current value, so it clears
     itself as soon as that field is fixed without waiting for another attempt. */
  const [showErrors, setShowErrors] = useState(false);
  const resultRef = useRef<HTMLDivElement>(null);
  const [progress, setProgress] = useState("");
  const [result, setResult] = useState<
    { ok: boolean; text: string; checklist?: boolean; warn?: string } | undefined
  >(undefined);

  const draft = (): NewSegmentDraft => ({
    label,
    family,
    termSetGuid,
    stagingFolder,
    permissioned: tiers.map((t) => ({ label: t })),
    below,
  });

  /* ── Load ──────────────────────────────────────────────────────────────────── */

  useEffect(() => {
    const load = async (): Promise<void> => {
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const config = encodeURIComponent(cachedListTitle(LIST_SUFFIX.config));

      // Id and TermSetGuid are for DELETION: the item id to remove the row, the term-set GUID to
      // find the segment's Group Map rows (a folder row's Segment holds it). Creation needs
      // neither — see ExistingSegment.
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${config}')/items` +
          `?$select=Id,Title,ModeLabel,StagingFolder,SortOrder,TermSetGuid&$filter=ConfigType eq 'mode'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) throw new Error(`the configuration list returned HTTP ${res.status}`);
      const rows = ((await res.json()).value ?? []) as Array<{
        Id?: number;
        Title?: string;
        ModeLabel?: string;
        StagingFolder?: string;
        SortOrder?: number;
        TermSetGuid?: string;
      }>;
      setExisting(
        rows.map((r) => ({
          key: (r.Title ?? "").trim(),
          label: (r.ModeLabel ?? r.Title ?? "").trim(),
          stagingFolder: (r.StagingFolder ?? "").trim(),
          sortOrder: r.SortOrder,
          itemId: r.Id,
          termSetGuid: (r.TermSetGuid ?? "").trim(),
        })),
      );

      // Seed the below-Unit tiers with the SAME built-in pair slice A seeds, from the same
      // helper. A new segment with an empty below-Unit list runs on that pair implicitly anyway,
      // so showing an empty list would make the first added level look like it REPLACED Year and
      // Document Type — dropping them from every future path with no error.
      const setRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${config}')/items` +
          `?$select=Title,SettingValue&$filter=ConfigType eq 'setting'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      let year = "";
      let docType = "";
      if (setRes.ok) {
        const srows = ((await setRes.json()).value ?? []) as Array<{
          Title?: string;
          SettingValue?: string;
        }>;
        const get = (k: string): string =>
          (srows.filter((r) => (r.Title ?? "").trim() === k)[0]?.SettingValue ?? "").trim();
        year = get("termSet_yearPeriod");
        docType = get("termSet_documentType");
      }
      setBelow(effectiveOnDemandTiers([], year, docType));
      setLoaded(true);
    };
    load().catch((e) => {
      setLoadError((e as Error).message);
      setLoaded(true);
    });
  }, []);

  /* ── Term set: resolve as it is typed ──────────────────────────────────────── */

  useEffect(() => {
    const guid = normalizeGuid(termSetGuid);
    if (!guid) {
      setCheck({ state: "blank" });
      return undefined;
    }
    if (!isGuid(guid)) {
      setCheck({ state: "malformed" });
      return undefined;
    }
    setCheck({ state: "checking" });
    let cancelled = false;
    const timer = setTimeout(() => {
      const run = async (): Promise<void> => {
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/v2.1/termStore/sets/${guid}?$select=id,localizedNames`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json" } },
        );
        if (cancelled) return;
        // 404 is the answer, not an error: term sets are per-site here, so one copied from
        // another site's config lands here too — as does a TERM's id pasted by mistake.
        if (res.status === 404) {
          setCheck({ state: "notfound" });
          return;
        }
        if (!res.ok) {
          setCheck({ state: "unknown", status: res.status });
          return;
        }
        const set = (await res.json()) as { localizedNames?: Array<{ name?: string }> };
        const name = (set.localizedNames ?? [])[0]?.name ?? "";
        let count = 0;
        const kids: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/v2.1/termStore/sets/${guid}/children?$select=id`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json" } },
        );
        if (kids.ok) count = (((await kids.json()).value ?? []) as unknown[]).length;
        if (!cancelled) setCheck({ state: "found", name, count });
      };
      run().catch(() => {
        if (!cancelled) setCheck({ state: "unknown", status: 0 });
      });
    }, 400);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
    // `recheck` IS A DEP SO THE BUTTON RE-RUNS THIS EFFECT, rather than calling the check function a
    // second way — see the state's own comment and `StructureManager.tsx`'s identical shape.
  }, [termSetGuid, recheck]);

  /* ── Dirty tracking ────────────────────────────────────────────────────────── */

  const dirty =
    label.trim() !== "" ||
    termSetGuid.trim() !== "" ||
    stagingFolder.trim() !== "" ||
    // Any tier at all is now an edit, since the list starts empty.
    tiers.length > 0;

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

  /* ⚠ THE MESSAGE WAS ALWAYS SET; IT WAS SIMPLY OFF SCREEN, AND THAT IS THE WHOLE BUG.
     `result` renders near the TOP of this page and the Create button sits about two hundred
     lines of JSX below it, so a refusal appeared somewhere the person who pressed the button
     could not see - reported as *"it doesnt show the error message"*. Identical to the
     Requests page's outcome banner (1.0.325.0), and fixed the same way.

     ⚠ DECLARED HERE, WITH THE OTHER HOOKS, ABOVE `if (!loaded) return`. Below that return it
     would run a different number of times on the render after loading finishes - *Rendered more
     hooks than during the previous render* - which blanks the whole web part with no error UI.
     That has cost this project three separate outages; the comment is the guard. */
  useEffect(() => {
    if (!result) return;
    try {
      resultRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
    } catch {
      /* Older browsers reject the options object. A message that did not scroll is the
         behaviour before this existed, so there is nothing to report. */
    }
  }, [result]);

  /* ── Depth ─────────────────────────────────────────────────────────────────── */

  /**
   * How many levels deep the term set actually goes.
   *
   * Walks level by level: the deepest level that still holds terms IS the depth. `undefined`
   * means the walk could not finish — a failed request, or the request cap — which
   * `depthVerdict` turns into a warning rather than a mismatch, because not knowing is not the
   * same as being wrong.
   *
   * It stops one level past the named tier count: at that point "deeper than you named" is
   * already settled and asking the rest changes nothing.
   */
  const measureDepth = async (guid: string, tierCount: number): Promise<number | undefined> => {
    let spent = 0;
    const childrenOf = async (termId?: string): Promise<string[] | undefined> => {
      if (spent >= DEPTH_REQUEST_CAP) return undefined;
      spent++;
      const url = termId
        ? `${siteUrl}/_api/v2.1/termStore/sets/${guid}/terms/${termId}/children?$select=id`
        : `${siteUrl}/_api/v2.1/termStore/sets/${guid}/children?$select=id`;
      const res: SPHttpClientResponse | undefined = await context.spHttpClient
        .get(url, SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } })
        .catch(() => undefined);
      if (!res || !res.ok) return undefined;
      const rows = ((await res.json()).value ?? []) as Array<{ id?: string }>;
      return rows.map((r) => r.id ?? "").filter((x) => x !== "");
    };

    const top = await childrenOf();
    if (top === undefined) return undefined;
    if (top.length === 0) return 0;

    let frontier = top;
    let depth = 1;
    while (depth <= tierCount && frontier.length > 0) {
      const next: string[] = [];
      for (let i = 0; i < frontier.length; i += DEPTH_BATCH) {
        const batch = frontier.slice(i, i + DEPTH_BATCH);
        setProgress(
          `Checking the term set — level ${depth + 1}, ${spent} of ${DEPTH_REQUEST_CAP} checks used…`,
        );
        const results = await Promise.all(batch.map((id) => childrenOf(id)));
        for (const r of results) {
          // A single unreadable branch makes the whole depth unknown. Treating it as "no
          // children" would report a SHALLOWER set than exists — the dangerous direction, since
          // it is the too-deep case that mis-provisions permissions silently.
          if (r === undefined) return undefined;
          next.push(...r);
        }
      }
      if (next.length === 0) return depth;
      depth++;
      frontier = next;
    }
    return depth;
  };

  /* ── Tier editing ──────────────────────────────────────────────────────────── */

  const addTier = (): void => {
    const name = newTier.trim();
    if (!name) return;
    setTiers([...tiers, name]);
    setNewTier("");
  };

  const moveTier = (i: number, delta: number): void => {
    const to = i + delta;
    if (to < 0 || to >= tiers.length) return;
    const next = tiers.slice();
    const item = next[i];
    next.splice(i, 1);
    next.splice(to, 0, item);
    setTiers(next);
  };

  /* ── Folder/file primitives shared by Create, Delete and Re-code ──────────────
     Extracted 2026-09-15 to `shared/spSegmentRecode.ts` so the CRS Term Abbreviations screen can
     offer the same re-code action with no second, drifting copy of this destructive-write logic.
     `retireLibraries`, `archiveTargets`, `walkFolders`, `countFiles`, `countArchiveDocumentsForCode`,
     `countArchiveDocuments`, `loadSegmentGroupMapRows`, `loadSegmentTermGuids`,
     `loadAbbreviationTermRows`, `loadSegmentAbbreviationRowIds`, `deleteItem`, `recycleFolder`,
     `countSegmentDocuments` (was `countSegment`) and `performSegmentRecode` all live there now,
     unchanged in behaviour — only parameterised on `(context, siteUrl, …)` instead of closing over
     component props. See docs/superpowers/specs/2026-09-15-segment-rename-and-group-rename-design.md. */

  /* ── Create ────────────────────────────────────────────────────────────────── */

  const create = async (): Promise<void> => {
    setBusy(true);
    setResult(undefined);
    setProgress("");
    try {
      const d = draft();
      // Computed once here, reused by the archive-clash check below AND the mode-row write at the
      // end — one derivation, so the two can never disagree about which folder is being claimed.
      const folder = sanitizeFolderSegment(d.stagingFolder).trim();

      // 1. VALIDATE FIRST. A rejected draft must leave nothing behind — no column, no row.
      const errors = validateNewSegment(d, existing);
      if (errors.length > 0) {
        // Mark the required fields red as well as printing the summary. The summary explains; the
        // red says WHICH box, which is what a long form actually needs.
        setShowErrors(true);
        setResult({ ok: false, text: errors.join(" ") });
        return;
      }
      // Past validation, so nothing is outstanding — clear the markers rather than leaving a form
      // that succeeded still wearing them.
      setShowErrors(false);

      /* 1.5. ARCHIVE CODE CLASH — added 2026-09-13, agreed with the client (via Reene) after the
         PCAR incident. See docs/superpowers/specs/2026-09-13-archive-code-reuse-guard-design.md.

         A segment's top folder carries no term — it is matched purely by NAME — and retiring a
         segment never deletes a non-empty Archive folder (7-year retention). So a DIFFERENT,
         unrelated segment reusing this same code would land its documents in the SAME Archive
         folder as whatever a previous, retired segment left behind: one folder literally named
         (say) "PCAR", visibly holding two unrelated businesses' records with nothing telling them
         apart. `validateNewSegment` above cannot catch this — it only compares against LIVE mode
         rows, and this collision is with something that no longer has one.

         ⚠ HARD REFUSAL, NO OVERRIDE, PER THE CLIENT'S OWN INSTRUCTION. Nothing needs one: once
         this check exists, the collision can never happen going forward, and a segment retired
         with a genuinely EMPTY archive already has that folder deleted (the 2026-09-12 feature),
         so a code becomes reusable again on its own the moment it is actually safe.

         ⚠ FAILS CLOSED on an unreadable check, same reasoning as `canDeleteArchive`: guessing
         "clear" when the read merely failed risks the exact silent collision this exists to
         prevent, discovered only much later by a confused client.

         Scoped to the segment's OWN top-folder code only — never department/unit abbreviations.
         SharePoint refuses two sibling folders with the same name, so once this one code is
         guaranteed unique, nothing beneath it can ever collide with a DIFFERENT segment's history:
         it is always created fresh under a folder that has never existed before. */

      /* ⚠ SUSPECTED LIVE 2026-09-13, the first time this guard was tested: it appeared to let "PCT"
         be reused with 3 real files still in Archive/PCT, with no refusal. Re-tested on a hard cache
         clear and the guard fired correctly — so THIS specific incident was very likely a stale
         package/cached tab (this project's own single most common false alarm), not this race. The
         race described below is real and worth guarding regardless, just not confirmed as the
         actual cause here: `archiveAvailable()` reads a cache `primeNames()` fills — awaited ONCE in
         this component's mount effect (line ~288) but NOT blocking the render, since React does not
         wait on effects before a component becomes interactive. Moving through "Create the term
         set" -> "New segment" -> Create fast enough on a freshly loaded page COULD reach this exact
         line before that background priming has settled, silently skipping the whole block below.
         The documented fix for this shape of race elsewhere in this codebase (the 1.0.207.0 race,
         GroupManager's own `loadModes`) is the same every time: await priming again right where it
         is needed, never trust a mount effect finished in time. `primeNames` memoises each of its
         own probes, so a second call here is cheap — instant if priming already finished, or it
         simply joins the same in-flight promise if not. */
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      if (archiveAvailable()) {
        setProgress("Checking the Archive library…");
        const archiveCount = await countArchiveDocumentsForCode(context, siteUrl, folder);
        if (archiveCount === undefined) {
          setResult({
            ok: false,
            text:
              `Could not confirm whether "${folder}" is already used in the Archive library — ` +
              `try again in a moment. Nothing has been created.`,
          });
          return;
        }
        if (archiveCount > 0) {
          setResult({
            ok: false,
            text:
              `"${folder}" already exists in the Archive library — it holds ${archiveCount} ` +
              `archived document${archiveCount === 1 ? "" : "s"}. Staging and Documents show it as ` +
              `free because that previous segment was retired, but its Archive folder was kept (7-year ` +
              `retention) — reusing this code now would mix that old segment's records in with this ` +
              `new one. Choose a different top folder name.`,
          });
          return;
        }
      }

      // 2. Depth. THE check: reconciliation walks the term tree and caps at the permissioned
      //    tier count, so a mismatch puts the folder ACLs on the wrong level — silently.
      setProgress("Checking the term set…");
      const depth = await measureDepth(normalizeGuid(d.termSetGuid), d.permissioned.length);
      const verdict = depthVerdict(depth, d.permissioned.length);
      if (!verdict.ok) {
        setResult({ ok: false, text: verdict.error ?? "The term set depth does not match." });
        return;
      }

      // 3. Columns, in BOTH libraries. A column with no row is a harmless orphan; a row naming a
      //    missing column breaks every upload in the segment — and validateUpdateListItem
      //    returns HTTP 200 with HasException, so that failure is not even loud.
      const created: string[] = [];
      for (const col of columnsForDraft(d)) {
        // FOUR libraries where the site has Highly Confidential, not two. A tier column absent
        // from one library fails the WHOLE metadata write for a document filed there — and
        // validateUpdateListItem answers HTTP 200 with HasException, so it is not even loud.
        for (const lib of allLibraryTitles()) {
          setProgress(`Creating ${col.internal} in ${lib}…`);
          if (await ensureColumn(context.spHttpClient, siteUrl, lib, col.internal, col.display)) {
            created.push(`${col.internal} (${lib})`);
          }
        }
      }

      // 4. The mode row, last.
      setProgress("Writing the configuration row…");
      const key = modeKeyFor(d.label);
      const sortOrder = nextSortOrder(existing);
      const write: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
          },
          body: JSON.stringify({
            Title: key,
            ConfigType: "mode",
            ModeLabel: d.label.trim(),
            Category: d.family,
            TermSetGuid: normalizeGuid(d.termSetGuid),
            StagingFolder: folder,
            SortOrder: sortOrder,
            Levels: JSON.stringify(buildSegmentLevels(d)),
          }),
        },
      );
      if (!write.ok) {
        const b = await write.text().catch(() => "");
        throw new Error(
          `The columns were created but the segment could not be saved (HTTP ${write.status}). ` +
            `${b.slice(0, 180)} Those columns are harmless and will be reused when you try again.`,
        );
      }

      setExisting([...existing, { key, label: d.label.trim(), stagingFolder: folder, sortOrder }]);
      // Announced only past every failure path above, so a host can treat it as proof the row exists
      // rather than as "Create was pressed". Guarded because the standalone page passes no handler.
      if (onCreated) onCreated(key, d.label.trim());
      setResult({
        ok: true,
        checklist: true,
        warn: verdict.warn,
        text:
          `"${d.label.trim()}" is configured — key ${key}, top folder ${folder}.` +
          (created.length > 0
            ? ` Created ${created.length} column(s): ${created.join(", ")}. They are not shown in` +
              ` any library view yet, which is deliberate so your existing views keep the layout` +
              ` you set — add them with the view's "Show or hide columns".`
            : " Every column it needs already existed, so none were created."),
      });
      // Recorded before the form is cleared, while the values are still in hand. What matters most
      // here is the DERIVED pair — key and sort order — because nobody types either, so if the
      // segment later fails to appear this row is the only place that says what they came out as.
      writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.segmentCreated,
        source: "SegmentCreator",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        segment: d.label.trim(),
        summary: `Segment created — ${d.label.trim()} (${folder})`,
        details: [
          `Key: ${key}`,
          `Top folder: ${folder}`,
          `Sort order: ${sortOrder}`,
          `Permissioned tiers: ${tiers.join(" → ")}`,
          `Term set: ${termSetGuid.trim()}`,
          created.length > 0
            ? `Columns created in both libraries: ${created.join(", ")}`
            : "No columns created — every one it needs already existed.",
          "Not yet done: abbreviations, groups and Group Map rows, then reconciliation.",
        ],
      }).catch(() => undefined);

      setLabel("");
      setTermSetGuid("");
      setStagingFolder("");
      setTiers([]);
      setNewTier("");
    } catch (e) {
      setResult({ ok: false, text: (e as Error).message });
    } finally {
      setBusy(false);
      setProgress("");
    }
  };

  /**
   * Re-read the mode rows only.
   *
   * The mount effect also seeds the below-Unit tiers from the settings rows, which must NOT be
   * re-run here: it would overwrite whatever the admin has typed into the create form while the
   * delete dialog was open.
   */
  const reload = async (): Promise<void> => {
    const config = encodeURIComponent(cachedListTitle(LIST_SUFFIX.config));
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${config}')/items` +
        `?$select=Id,Title,ModeLabel,StagingFolder,SortOrder,TermSetGuid&$filter=ConfigType eq 'mode'&$top=200`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return;
    const rows = ((await res.json()).value ?? []) as Array<{
      Id?: number;
      Title?: string;
      ModeLabel?: string;
      StagingFolder?: string;
      SortOrder?: number;
      TermSetGuid?: string;
    }>;
    setExisting(
      rows.map((r) => ({
        key: (r.Title ?? "").trim(),
        label: (r.ModeLabel ?? r.Title ?? "").trim(),
        stagingFolder: (r.StagingFolder ?? "").trim(),
        sortOrder: r.SortOrder,
        itemId: r.Id,
        termSetGuid: (r.TermSetGuid ?? "").trim(),
      })),
    );
  };

  /* ── Delete a segment ──────────────────────────────────────────────────────────
     Spec: docs/superpowers/specs/2026-08-14-delete-segment-design.md

     Delete RETIRES a segment: the mode row plus its Group Map rows. It touches no document and
     no column. Deleting the folders is a separate opt-in, off by default, because that is the
     only irreversible half — and even that goes to the recycle bin.

     The rules that decide what may be offered live in shared/segmentDeletion.ts, tested, because
     the failure here is not a wrong number on screen: it is an archive nobody could confirm was
     empty being deleted.

     ⚠ `retireLibraries`, `archiveTargets`, `walkFolders`, `countFiles`, `countArchiveDocumentsForCode`,
     `countArchiveDocuments`, `loadSegmentGroupMapRows`, `loadSegmentTermGuids`,
     `loadAbbreviationTermRows`, `loadSegmentAbbreviationRowIds`, `deleteItem`, `recycleFolder` and
     `countSegmentDocuments` (was `countSegment`) — every low-level primitive this delete flow is
     built on — now live in `shared/spSegmentRecode.ts` (moved 2026-09-15, alongside the recode
     flow's own write sequence, for the same reason). Look there for their definitions; nothing
     about their behaviour changed, only where they are declared. */

  const openDelete = (seg: ExistingSegment): void => {
    setDeleting(seg);
    setDelCounts(undefined);
    setDelFolders(false);
    setDelTyped("");
    setDelLog([]);
    setDelProgress(undefined);
    countSegmentDocuments(context, siteUrl, seg)
      .then((c) => {
        setDelCounts(c);
        /**
         * PRE-TICKED ONLY FOR AN EMPTY SHELL (client, 2026-08-26: *"the whole point of retiring a
         * segment is to delete everything"* — they had left GHO's folders behind by not ticking it).
         *
         * ⚠ TIED TO THE DOCUMENT COUNT, NEVER FLAT-ON. The checkbox's own label is conditional: with
         * documents present it reads "including N documents", so a flat default would point the
         * default action at recycling real files. Retiring usually happens AFTER the migrator has
         * moved documents out, which is exactly the `documents === 0` case — so this gives the
         * intended behaviour where it is harmless and withholds it where it is not.
         *
         * `state === "counted"` is required as well: an UNKNOWN count means `canOfferFolderDelete`
         * refuses to render the checkbox at all, and a hidden-but-true flag would be a deletion the
         * admin was never shown. `alsoFolders` in `onDelete` re-checks the same guard, so this is
         * belt-and-braces rather than the only thing standing in the way.
         *
         * The typed confirmation is unaffected — `needsTypedConfirmation` returns true whenever
         * `deleteFolders` is true, so pre-ticking never removes a gate.
         */
        setDelFolders(c.state === "counted" && c.documents === 0);
      })
      .catch((e) => setDelCounts(unknownCounts((e as Error).message)));
  };

  const onDelete = async (): Promise<void> => {
    if (!deleting || !delCounts) return;
    const seg = deleting;
    const alsoFolders = delFolders && canOfferFolderDelete(delCounts);
    setBusy(true);
    const lines: string[] = [];
    let ok = true;
    try {
      // 1. Group Map rows. A row under a segment that no longer exists is a grant nobody can see
      //    or manage, while reconciliation keeps re-applying it.
      let rowsRemoved = 0;
      let rowIds: number[] = [];
      try {
        rowIds = await loadSegmentGroupMapRows(context, siteUrl, seg);
      } catch (e) {
        lines.push(`Could not read the folder-access mappings (${(e as Error).message}) — none were removed.`);
        ok = false;
      }
      setDelProgress(
        rowIds.length > 0
          ? { done: 0, total: rowIds.length, label: "Removing folder-access mappings" }
          : undefined,
      );
      for (let i = 0; i < rowIds.length; i++) {
        try {
          await deleteItem(context, siteUrl, cachedListTitle(LIST_SUFFIX.groupMap), rowIds[i]);
          rowsRemoved++;
        } catch {
          ok = false;
        }
        setDelProgress({ done: i + 1, total: rowIds.length, label: "Removing folder-access mappings" });
      }
      if (rowIds.length > 0) {
        lines.push(`Folder-access mappings removed: ${rowsRemoved} of ${rowIds.length}.`);
      }
      setDelProgress(undefined);

      // 2. Abbreviation rows. MANDATORY, no opt-out (2026-09-11, client: "no need checkbox as an
      //    option, make it mandatory") — see docs/superpowers/specs/2026-09-11-retire-deletes-
      //    abbreviations-design.md. Walked FRESH here rather than trusting delCounts, because state
      //    may have moved between opening the dialog and pressing the button. Non-blocking, same
      //    as the Group Map cleanup above: a failed read or a partial delete is logged and the
      //    retire continues — this is cleanup, not the critical step.
      let abbrevRemoved = 0;
      let abbrevIds: number[] = [];
      let abbrevKnown = true;
      try {
        const abbrev = await loadSegmentAbbreviationRowIds(context, siteUrl, seg);
        if (!abbrev.complete) {
          abbrevKnown = false;
          lines.push("Could not fully read the term store, so its abbreviation rows were not removed — check for leftovers after reconciliation runs.");
          ok = false;
        } else {
          abbrevIds = abbrev.ids;
        }
      } catch (e) {
        abbrevKnown = false;
        lines.push(`Could not read the abbreviation rows (${(e as Error).message}) — none were removed.`);
        ok = false;
      }
      if (abbrevKnown) {
        setDelProgress(
          abbrevIds.length > 0
            ? { done: 0, total: abbrevIds.length, label: "Removing abbreviation rows" }
            : undefined,
        );
        for (let i = 0; i < abbrevIds.length; i++) {
          try {
            await deleteItem(context, siteUrl, cachedListTitle(LIST_SUFFIX.abbreviation), abbrevIds[i]);
            abbrevRemoved++;
          } catch {
            ok = false;
          }
          setDelProgress({ done: i + 1, total: abbrevIds.length, label: "Removing abbreviation rows" });
        }
        if (abbrevIds.length > 0) {
          lines.push(`Abbreviation rows removed: ${abbrevRemoved} of ${abbrevIds.length}.`);
        }
        setDelProgress(undefined);
      }

      // 3. The mode row. If THIS fails, stop: a run that removed the grants but left the segment
      //    live is a segment whose uploaders have quietly lost their access.
      if (seg.itemId === undefined) throw new Error("this segment's configuration row has no id, so it cannot be deleted");
      await deleteItem(context, siteUrl, cachedListTitle(LIST_SUFFIX.config), seg.itemId);
      lines.push("The segment is no longer offered in the upload form, and reconciliation will not walk it.");

      // 4. Folders, only if asked. After the row, never before — see the spec's D6.
      if (alsoFolders) {
        const root = (seg.stagingFolder ?? "").trim();
        for (const target of retireLibraries()) {
          try {
            // urlSegment builds the request; title is what the admin recognises in the log.
            const went = await recycleFolder(context, siteUrl, target.urlSegment, root);
            lines.push(
              went
                ? `${target.title}/${root} moved to the recycle bin.`
                : `${target.title}/${root} did not exist.`,
            );
          } catch (e) {
            lines.push(`Could not delete ${target.title}/${root} — ${(e as Error).message}`);
            ok = false;
          }
        }

        // 4b. Archive/ArchiveHC — ONLY when re-checked here as confirmed empty. Never trust
        // `delCounts` alone: it was read when the dialog opened, and state can move between then
        // and this button press (the same reasoning `loadSegmentGroupMapRows`/
        // `loadSegmentAbbreviationRowIds` above are re-walked fresh, not read from `delCounts`).
        // Spec: docs/superpowers/specs/2026-09-12-empty-archive-deletion-on-retire-design.md
        if (archiveAvailable()) {
          const freshArchiveCount = await countArchiveDocuments(context, siteUrl, seg).catch(() => undefined);
          if (canDeleteArchive({ ...delCounts, archiveDocuments: freshArchiveCount })) {
            for (const target of archiveTargets()) {
              try {
                const went = await recycleFolder(context, siteUrl, target.urlSegment, root);
                lines.push(
                  went
                    ? `${target.title}/${root} was empty and moved to the recycle bin.`
                    : `${target.title}/${root} did not exist.`,
                );
              } catch (e) {
                lines.push(`Could not delete ${target.title}/${root} — ${(e as Error).message}`);
                ok = false;
              }
            }
          } else {
            // Either it genuinely has content, or the re-check could not confirm it — either way,
            // fail CLOSED and say nothing was touched, rather than silently doing nothing with no
            // explanation. `survivorLines` below states which of the two applies, from `delCounts`
            // (the dialog-open-time read) — re-reading a second reason string here would risk it
            // disagreeing with what the dialog already told the admin.
            lines.push("The archive was left untouched — it either holds documents or could not be re-confirmed as empty.");
          }
        }

        // Folder Map rows are derivable, so they go exactly when the folders do — otherwise they
        // point at UniqueIds that no longer resolve.
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
              (r) => (r.Section ?? "").trim().toLowerCase() === root.toLowerCase(),
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
      }

      lines.push(...survivorLines(alsoFolders, archiveAvailable(), delCounts.archiveDocuments));
      if (rowsRemoved > 0) {
        lines.push("Folder permissions stay in place until Folder Reconciliation runs.");
      }
      setDelLog(lines);
      setResult({
        ok,
        text: ok
          ? `"${seg.label}" deleted. ${deletionSummary(delCounts, alsoFolders)}`
          : `"${seg.label}" was deleted, but not everything succeeded — see below.`,
      });

      writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.segmentDeleted,
        outcome: ok ? "Success" : "Failed",
        source: "SegmentCreator",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        segment: seg.label,
        summary: `Segment deleted — ${seg.label} (${seg.stagingFolder})`,
        details: [
          `Key: ${seg.key}`,
          `Top folder: ${seg.stagingFolder}`,
          `Folders deleted: ${alsoFolders ? "YES — moved to the recycle bin" : "no"}`,
          delCounts.state === "counted"
            ? `Counted before deleting: ${delCounts.folders} folder(s), ${delCounts.documents} document(s), ${delCounts.groupMapRows} mapping(s), ${delCounts.abbreviationRows} abbreviation(s)`
            : `Counts were UNKNOWN before deleting (${delCounts.reason ?? "unreadable"})`,
          ...lines,
        ],
      }).catch(() => undefined);

      setDeleting(undefined);
      await reload();
    } catch (e) {
      setDelProgress(undefined);
      setDelLog([...lines, `Stopped: ${(e as Error).message}`]);
      setResult({ ok: false, text: `"${seg.label}" was NOT deleted — ${(e as Error).message}` });
    } finally {
      setBusy(false);
    }
  };

  /* ── Re-code a segment's top folder ───────────────────────────────────────────
     Spec: docs/superpowers/specs/2026-09-11-segment-recode-design.md

     The segment container folder has NO term backing it — unlike Department and Unit, it has no
     abbreviation row and no term-keyed Folder Map row, so it is created BY NAME straight from the
     mode row's `StagingFolder`. Reconciliation cannot rename it the way it renames a Department or
     Unit folder (by UniqueId), so a bare edit of `StagingFolder` would leave a brand-new empty tree
     built under the new name while the OLD tree — documents included — sits unrecognised.

     Offered only for a genuinely EMPTY segment (documents === 0), reusing the exact same count and
     libraries the delete dialog already uses. Group Map rows and groups are UNTOUCHED: a row's
     `Segment` holds the term-SET GUID and `UnitTermGuid` a term GUID, neither of which depends on
     `StagingFolder` — and `suggestGroupName` derives a group's stem from the segment LABEL, never
     from this key. Only the mode row's `StagingFolder` cell and the Folder Map rows naming the OLD
     `Section` need touching; reconciliation rebuilds the new tree on its own next run. */

  const openRecode = (seg: ExistingSegment): void => {
    setRecoding(seg);
    setRecodeCounts(undefined);
    setRecodeNewFolder(seg.stagingFolder);
    setRecodeTyped("");
    setRecodeLog([]);
    countSegmentDocuments(context, siteUrl, seg)
      .then((c) => setRecodeCounts(c))
      .catch((e) => setRecodeCounts(unknownCounts((e as Error).message)));
  };

  /**
   * A thin wrapper around `performLiveSegmentRecode` — this component keeps only its own state
   * updates, the result banner, the audit-row write (with its own `source: "SegmentCreator"`) and
   * the list reload.
   *
   * ⚠ ALWAYS THE LIVE, IN-PLACE RENAME NOW — the recycle-and-rebuild path (`performSegmentRecode`)
   * is NO LONGER CALLED from here (client, 2026-09-16, watching the empty-segment case recycle its
   * folders: *"Can we ensure for renaming the term abbreviations doesn't move it to recycle bin if
   * its empty"*). `performLiveSegmentRecode` already handles a confirmed-empty segment correctly —
   * `renameFolder` treats a library where the old folder does not exist as a no-op skip, not a
   * failure — so there is no longer a reason to prefer the destructive recycle for that case: the
   * in-place rename does strictly less (nothing is ever sent to the recycle bin, and any already-
   * provisioned Department/Unit subfolders survive the rename rather than being blown away and
   * left for reconciliation to rebuild from scratch). `performSegmentRecode`/`recodeSummary` stay
   * defined in `shared/spSegmentRecode.ts`/`shared/segmentRecode.ts` — kept, not deleted — in case
   * a reason to recycle ever comes back; nothing here calls them any more.
   */
  const onRecode = async (): Promise<void> => {
    if (!recoding || !recodeCounts) return;
    const seg = recoding;
    const oldRoot = (seg.stagingFolder ?? "").trim();
    const newRoot = sanitizeFolderSegment(recodeNewFolder).trim();
    setBusy(true);
    try {
      const { ok, renamed, lines } = await performLiveSegmentRecode(context, siteUrl, seg, recodeNewFolder);
      setRecodeLog(lines);
      setResult({
        ok,
        text: ok
          ? `"${seg.label}" ${recodeSummaryLive(oldRoot, newRoot, recodeCounts)}`
          : `"${seg.label}"'s top folder change did not fully succeed — see below.`,
      });

      writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.segmentRecoded,
        outcome: ok ? "Success" : "Failed",
        source: "SegmentCreator",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        segment: seg.label,
        summary: `Segment top folder re-coded — ${seg.label}: "${oldRoot}" → "${newRoot}"`,
        details: [
          `Key: ${seg.key}`,
          `Old top folder: ${oldRoot}`,
          `New top folder: ${newRoot}`,
          "Mode: live rename (in place, every library, nothing recycled)",
          recodeCounts.state === "counted"
            ? `Counted before recoding: ${recodeCounts.folders} folder(s), ${recodeCounts.documents} document(s)`
            : `Counts were UNKNOWN before recoding (${recodeCounts.reason ?? "unreadable"})`,
          ...lines,
        ],
      }).catch(() => undefined);

      // Gated on `renamed`, NOT `ok`: `ok` also folds in a non-fatal Folder Map tidy-up failure,
      // and the mode row's `StagingFolder` may or may not actually have been written. Reflecting the
      // new name (or reloading expecting to see it) when the write never landed would show a folder
      // name the server does not have — `renamed` is true only once that write is confirmed.
      if (renamed) {
        // Reflect the new folder locally so the list does not show the stale value until a reload.
        setExisting(existing.map((e) => (e.key === seg.key ? { ...e, stagingFolder: newRoot } : e)));
        setRecoding(undefined);
        await reload();
      }
    } catch (e) {
      // Neither write sequence throws — this only catches a failure in the surrounding
      // bookkeeping (the audit write, the reload) so the admin still sees SOMETHING rather than a
      // silently dead button.
      setRecodeLog([`Stopped: ${(e as Error).message}`]);
      setResult({ ok: false, text: `"${seg.label}"'s top folder was NOT changed — ${(e as Error).message}` });
    } finally {
      setBusy(false);
    }
  };

  /* ── Render ────────────────────────────────────────────────────────────────── */

  const key = modeKeyFor(label);
  /* Live, per-field collisions — the same ones Create refuses, rendered next to the field that caused
     them. This form is taller than the viewport, so the summary at the top is off-screen exactly when it
     matters, and pressing Create reads as doing nothing (client, 2026-08-17). Recomputed every render
     rather than in state: it is three array scans over a list of segments, and state would be one more
     thing that can disagree with the fields. */
  const conflicts = fieldConflicts(draft(), existing);
  /* What each field shows, from TWO sources that cannot both fire on one field.
     `fieldConflicts` is a clash with an existing segment - it needs the field to hold something
     valid, so it is safe to show live, as you type. `requiredFieldErrors` is blank-or-malformed,
     which is true of an untouched form, so it is gated on a refused save. Merged here so the
     three inputs read ONE value each and cannot disagree about whether they are in error. */
  const required = requiredFieldErrors(draft());
  const fieldError = {
    label: conflicts.label || (showErrors ? required.label : ""),
    folder: conflicts.folder || (showErrors ? required.folder : ""),
    termSet: conflicts.termSet || (showErrors ? required.termSet : ""),
  };
  const folderPreview = sanitizeFolderSegment(stagingFolder).trim();
  const pathPreview =
    `/${folderPreview || "TOPFOLDER"}/` +
    [...tiers.map((t) => t || "?"), ...below.map((b) => b.label)].map((n) => `<${n}>`).join("/");

  if (!loaded) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading&hellip;</p>;

  const typedOk =
    delCounts !== undefined &&
    deleting !== undefined &&
    (!needsTypedConfirmation(delCounts, delFolders) || confirmationMatches(delTyped, deleting.label));

  /* Recode's own readiness: the count must have succeeded (empty OR holding documents — since
     2026-09-16 a documented segment is a VALID, different path, not a refusal), the sanitized new
     folder must not collide with another segment or be the current value unchanged, and — always,
     since this action never has a "harmless" case once offered — the segment's label must be
     typed out. */
  const recodeConflictMsg = recoding
    ? folderRecodeConflict(recodeNewFolder, existing, recoding.key)
    : "";
  const recodeIsNoOp = recoding ? folderRecodeIsNoOp(recodeNewFolder, recoding.stagingFolder) : true;
  const recodeSanitized = sanitizeFolderSegment(recodeNewFolder).trim();
  const recodeOk =
    recoding !== undefined &&
    recodeCounts !== undefined &&
    canOfferAnyRecode(recodeCounts) &&
    recodeSanitized.length > 0 &&
    !recodeConflictMsg &&
    !recodeIsNoOp &&
    confirmationMatches(recodeTyped, recoding.label);

  return (
    <div>
      {loadError && (
        <div style={{ ...s.msg, ...s.err }}>
          Could not read the existing segments — {loadError}. Adding one now risks a duplicate name
          or a shared top folder, so fix that first.
        </div>
      )}

      {/* ── The segments that exist, each removable ─────────────────────────────
          Spec: 2026-08-14-delete-segment-design.md. Delete RETIRES the segment; it deletes no
          document and no column unless the folder option is ticked in the dialog. */}
      {/* GATED ON `allowDelete`, WHICH IS FALSE IN THE CREATE FLOW AND NOWHERE ELSE (client,
          2026-09-09: *"Remove the the Segments on this site (7), client dont think its needed."*
          Their screenshot was step 2 of Add a new segment, where this list is purely informational
          and its own hint said so.

          The list is shown exactly where it can be ACTED on, so one flag decides both and they
          cannot drift: `hideSegmentDelete` is set only for `newSegment`, so Retire keeps the list
          (client, 2026-09-09: *"it should only show, Segments on this site"* - it IS that flow's
          whole screen) and the standalone Segments tab keeps it too, where deleting is why anyone
          opens it. Split them into a separate `hideSegmentList` prop only if a caller ever needs
          the list without the buttons or the reverse.

          WARN: `existing` IS STILL READ AND STILL POPULATED - `fieldConflicts` refuses a duplicate
          name or top folder from it, so this hides the RENDER and nothing else. Do not "tidy" the
          read away with it. */}
      {allowDelete !== false && existing.length > 0 && (
        <div style={s.card}>
          <p style={s.h}>Segments on this site ({existing.length})</p>
          {existing.map((seg) => (
            <div key={seg.key} style={{ ...s.tierRow, background: "#fff", border: "1px solid #f0f0f0" }}>
              <span style={s.tierName}>{seg.label || seg.key}</span>
              <span style={s.tierMeta}>
                top folder <strong>{seg.stagingFolder || "—"}</strong>
              </span>
              {/* Unconditional, and only because the whole list above is now gated on the same
                  flag - `tsc` caught the leftover `allowDelete !== false` here as provably always
                  true, which is what a redundant guard looks like from the outside. If the list is
                  ever shown where Delete must not be, bring the guard back rather than disabling
                  the button: a greyed Delete tells an admin the option belongs here and invites
                  hunting for the way to enable it. */}
              {/* Re-code the top folder — a physical rename, spec 2026-09-11-segment-recode-design.md.
                  A separate button rather than a mode on Delete: the two gate on different counts
                  (delete cares about mapping rows too; recode only about documents) and end in
                  opposite directions (delete removes the row, recode keeps it and rebuilds it).

                  GATED ON `allowRecode`, false only on the Retire screen (client, 2026-09-13). Retiring
                  and recoding are opposite actions on the same row and do not belong on one screen
                  together. */}
              {allowRecode !== false && (
                <button
                  style={s.ghost}
                  disabled={busy || seg.itemId === undefined}
                  title={
                    seg.itemId === undefined
                      ? "This row has no id, so it cannot be updated from here."
                      : "Rename the segment's top folder — offered only when the segment is empty"
                  }
                  onClick={() => openRecode(seg)}
                >
                  Re-code folder
                </button>
              )}
              <button
                style={s.danger}
                disabled={busy || seg.itemId === undefined}
                title={
                  seg.itemId === undefined
                    ? "This row has no id, so it cannot be deleted from here."
                    : "Remove this segment"
                }
                onClick={() => openDelete(seg)}
              >
                Delete
              </button>
            </div>
          ))}
          <p style={s.hint}>
            Deleting a segment stops it being offered, removes its folder-access mappings, and deletes its folder abbreviations (re-creating it later means retyping every code). It does not delete any document, and never deletes a column.
          </p>
        </div>
      )}

      {delLog.length > 0 && (
        <div style={{ ...s.msg, ...s.ok }}>
          {delLog.map((line, i) => (
            <div key={i} style={{ marginBottom: 3 }}>{line}</div>
          ))}
        </div>
      )}

      {recodeLog.length > 0 && (
        <div style={{ ...s.msg, ...s.ok }}>
          {recodeLog.map((line, i) => (
            <div key={i} style={{ marginBottom: 3 }}>{line}</div>
          ))}
        </div>
      )}

      {deleting !== undefined && (
        <div
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={() => setDeleting(undefined)}
        >
          <div
            style={{ background: "#fff", borderRadius: 8, padding: 20, width: "min(600px, 92vw)", maxHeight: "86vh", overflowY: "auto" }}
            onClick={(e) => e.stopPropagation()}
          >
            <p style={{ ...s.h, fontSize: 15 }}>Delete &quot;{deleting.label}&quot;?</p>

            {delCounts === undefined && (
              <p style={{ fontSize: 13, color: "#605e5c" }}>Counting what this segment holds&hellip;</p>
            )}

            {delCounts !== undefined && (
              <>
                <p style={{ fontSize: 13, margin: "0 0 10px" }}>
                  The segment stops being offered in the upload form, and Folder Reconciliation
                  stops walking it.
                </p>

                {delCounts.state === "counted" ? (
                  <p style={{ fontSize: 13, margin: "0 0 10px" }}>
                    It currently holds <strong>{delCounts.folders}</strong> folder
                    {delCounts.folders === 1 ? "" : "s"} and{" "}
                    <strong>{delCounts.documents}</strong> document
                    {delCounts.documents === 1 ? "" : "s"} across{" "}
                    {/* DERIVED, never "both": this counts four libraries on an HC site and two
                        without one, and it said "both" while silently reading only the approval
                        library until 2026-08-26. A number that cannot disagree with the loop that
                        produced it. */}
                    <strong>{retireLibraries().length}</strong> librar
                    {retireLibraries().length === 1 ? "y" : "ies"}, and{" "}
                    <strong>{delCounts.groupMapRows}</strong> folder-access mapping
                    {delCounts.groupMapRows === 1 ? "" : "s"}.{" "}
                    {/* Mandatory, no checkbox (2026-09-11) — this always happens, so it is stated
                        plainly rather than folded behind an option. */}
                    Deleting also removes <strong>{delCounts.abbreviationRows}</strong> abbreviation
                    {delCounts.abbreviationRows === 1 ? " row" : " rows"} — re-creating this segment
                    later means retyping every code.
                  </p>
                ) : (
                  /* Withholding the folder option is not enough on its own — an unexplained
                     missing checkbox reads as a broken page. Say which read failed. */
                  <div style={{ ...s.msg, ...s.warn, marginBottom: 10 }}>
                    Could not count what this segment holds — {delCounts.reason}. You can still
                    remove the segment, but <strong>deleting its folders is not offered</strong>,
                    because nothing here can confirm they are empty. Delete them by hand in
                    SharePoint if you need to.
                  </div>
                )}

                {canOfferFolderDelete(delCounts) && (
                  <label style={{ display: "block", fontSize: 13, margin: "0 0 10px", cursor: "pointer" }}>
                    <input
                      type="checkbox"
                      checked={delFolders}
                      disabled={busy}
                      onChange={(e) => setDelFolders(e.target.checked)}
                      style={{ marginRight: 8 }}
                    />
                    Also delete the folders{delCounts.documents > 0
                      ? ` — including ${delCounts.documents} document${delCounts.documents === 1 ? "" : "s"}`
                      : " (they are empty)"}
                  </label>
                )}

                <div style={{ ...s.msg, ...s.ok, marginBottom: 10 }}>
                  <strong>What survives</strong>
                  {survivorLines(
                    delFolders && canOfferFolderDelete(delCounts),
                    archiveAvailable(),
                    delCounts.archiveDocuments,
                  ).map((line, i) => (
                    <div key={i} style={{ marginTop: 4 }}>{line}</div>
                  ))}
                </div>

                {delCounts.groupMapRows > 0 && (
                  <div style={{ ...s.msg, ...s.warn, marginBottom: 10 }}>
                    Removing the mappings is not removing the access.{" "}
                    <strong>
                      The folder permissions they granted stay in place until Folder Reconciliation
                      runs.
                    </strong>
                  </div>
                )}

                {needsTypedConfirmation(delCounts, delFolders) && (
                  <>
                    <label style={s.label}>
                      Type <strong>{deleting.label}</strong> to confirm
                    </label>
                    <input
                      style={s.input}
                      value={delTyped}
                      disabled={busy}
                      onChange={(e) => setDelTyped(e.target.value)}
                    />
                  </>
                )}

                <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
                  <button
                    style={typedOk && !busy ? s.danger : s.off}
                    disabled={!typedOk || busy}
                    onClick={() => { onDelete().catch(() => undefined); }}
                  >
                    {busy && delProgress
                      ? `Deleting… (${delProgress.done} of ${delProgress.total})`
                      : busy
                      ? "Deleting…"
                      : "Delete segment"}
                  </button>
                  <button style={s.ghost} disabled={busy} onClick={() => setDeleting(undefined)}>
                    Cancel
                  </button>
                </div>
                {busy && delProgress && (
                  <div style={{ marginTop: 10 }}>
                    <div style={{ height: 6, borderRadius: 3, background: "#edebe9", overflow: "hidden" }}>
                      <div
                        style={{
                          height: "100%",
                          borderRadius: 3,
                          background: "#0f6c3f",
                          width: `${Math.round((delProgress.done / delProgress.total) * 100)}%`,
                          transition: "width 120ms linear",
                        }}
                      />
                    </div>
                    <p style={{ fontSize: 12, color: "#605e5c", margin: "6px 0 0" }}>
                      {delProgress.label} — {delProgress.done} of {delProgress.total}. This
                      is one request per row; do not close this tab until it finishes.
                    </p>
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      )}

      {/* ── Re-code a segment's top folder ───────────────────────────────────────
          Spec: 2026-08-14-delete-segment-design.md's sibling, 2026-09-11-segment-recode-design.md,
          widened 2026-09-16 to a segment holding documents too, then SIMPLIFIED the same day: the
          top folder is now ALWAYS renamed IN PLACE, in every library including the archive, with
          every document, version, approval status and permission preserved — whether the segment
          is empty or holds real documents. Nothing is ever recycled (client, watching the empty
          case recycle its folders: "Can we ensure for renaming the term abbreviations doesn't move
          it to recycle bin if its empty"). Reuses the same count the delete dialog runs; only the
          ONE refusal left (the count itself came back unreadable) still shows the error branch
          below. */}
      {recoding !== undefined && (
        <div
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={() => setRecoding(undefined)}
        >
          <div
            style={{ background: "#fff", borderRadius: 8, padding: 20, width: "min(600px, 92vw)", maxHeight: "86vh", overflowY: "auto" }}
            onClick={(e) => e.stopPropagation()}
          >
            <p style={{ ...s.h, fontSize: 15 }}>
              Re-code the top folder for &quot;{recoding.label}&quot;?
            </p>

            {/* WARN: the whole form used to be GATED on `recodeCounts !== undefined`, so for a
                large segment (GHO — 8 departments, 62 units, deep below-Unit chains, 4 libraries)
                the count's fully-sequential folder walk can take well over a minute, and for that
                whole minute the admin saw nothing but "Checking…" with no input anywhere on
                screen (client, 2026-09-16: "still no input for segment for term abbreviation").
                The FORM now renders immediately — typing and the typed confirmation do not need
                the count at all — and only the SUBMIT is held back until `recodeCounts` resolves,
                via the existing `recodeOk`/button-disabled logic below. Only once the count comes
                back UNKNOWN (a genuine failure, not merely "still running") does this swap to the
                refusal-only view, since at that point there is nothing left to fill in for. */}
            {recodeCounts !== undefined && !canOfferAnyRecode(recodeCounts) ? (
              <>
                <div style={{ ...s.msg, ...s.err, marginBottom: 10 }}>
                  {recodeRefusalReason(recodeCounts)}
                </div>
                <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
                  <button style={s.ghost} onClick={() => setRecoding(undefined)}>
                    Close
                  </button>
                </div>
              </>
            ) : (
              <>
                {recodeCounts === undefined && (
                  <p style={{ fontSize: 13, color: "#605e5c", marginBottom: 10 }}>
                    Checking what this segment holds — this can take a while for a large segment.
                    You can start filling in the form below now; the button unlocks once the check
                    finishes.
                  </p>
                )}

                {recodeCounts !== undefined && (
                  <>
                    <p style={{ fontSize: 13, margin: "0 0 10px" }}>
                      This segment holds <strong>{recodeCounts.documents}</strong> document
                      {recodeCounts.documents === 1 ? "" : "s"}. Its top folder is currently{" "}
                      <strong>{recoding.stagingFolder || "—"}</strong>.
                    </p>
                    <div style={{ ...s.msg, ...s.warn, marginBottom: 10 }}>
                      The top folder is renamed <strong>in place</strong>, in every library
                      including the archive — nothing is ever recycled or re-created, even for an
                      empty segment. Every document, version, approval status and permission
                      travels with it. <strong>Run Folder Reconciliation afterwards</strong> to
                      confirm everything still resolves. Groups and their mappings are untouched;
                      they are keyed on the segment&apos;s term set, never on this folder name.
                    </div>
                  </>
                )}

                <label style={s.label}>New top folder name</label>
                <input
                  style={{ ...s.input, ...(recodeConflictMsg ? s.inputBad : {}) }}
                  value={recodeNewFolder}
                  disabled={busy}
                  onChange={(e) => setRecodeNewFolder(e.target.value)}
                />
                {recodeConflictMsg ? (
                  <p style={s.fieldErr}>{recodeConflictMsg}</p>
                ) : recodeSanitized.length === 0 ? (
                  <p style={s.fieldErr}>Give the segment a top folder name.</p>
                ) : recodeIsNoOp ? (
                  <p style={s.hint}>That is the current folder — type a different code to change it.</p>
                ) : recodeCounts === undefined ? (
                  <p style={s.hint}>Ready — waiting for the check to finish before this can be submitted.</p>
                ) : (
                  <p style={s.hint}>
                    Every folder will be renamed to &quot;{recodeSanitized}&quot; here — in place.
                  </p>
                )}

                <label style={s.label}>
                  Type <strong>{recoding.label}</strong> to confirm
                </label>
                <input
                  style={s.input}
                  value={recodeTyped}
                  disabled={busy}
                  onChange={(e) => setRecodeTyped(e.target.value)}
                />

                <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
                  <button
                    style={recodeOk && !busy ? s.btn : s.off}
                    disabled={!recodeOk || busy}
                    onClick={() => { onRecode().catch(() => undefined); }}
                  >
                    {recodeCounts === undefined
                      ? "Waiting for the check…"
                      : busy
                        ? "Renaming…"
                        : "Re-code folder"}
                  </button>
                  <button style={s.ghost} disabled={busy} onClick={() => setRecoding(undefined)}>
                    Cancel
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {result && (
        <div ref={resultRef} style={{ ...s.msg, ...(result.ok ? s.ok : s.err) }}>
          <div>{result.text}</div>
          {result.warn && <div style={{ marginTop: 8, fontWeight: 600 }}>{result.warn}</div>}
          {result.checklist && (
            <div style={{ marginTop: 12 }}>
              <p style={s.h}>It is not usable yet. Four steps remain:</p>
              <ol style={{ margin: 0, paddingLeft: 18 }}>
                <li style={s.li}>
                  <strong>An abbreviation for every term</strong>, in{" "}
                  {cachedListTitle(LIST_SUFFIX.abbreviation)}. A term without one is skipped by
                  every reconciliation run — no folder, and that unit cannot upload.
                </li>
                <li style={s.li}>
                  <strong>Groups and their Group Map rows</strong> — the Folder Access page.
                </li>
                <li style={s.li}>
                  <strong>Run Folder Reconciliation</strong>. That is what creates the folders and
                  grants access; nothing on this page created a folder.
                </li>
                <li style={s.li}>
                  Check an uploader can see it. Until reconciliation has produced their folder the
                  segment stays hidden from them — deliberately, so nobody uploads into a
                  half-built segment.
                </li>
              </ol>
            </div>
          )}
        </div>
      )}

      {allowCreate !== false && (
      <>
      {/* ⚠ HIDDEN ONCE THE TERM SET RESOLVES — found live 2026-09-13, alongside the archive
          code-reuse guard. This banner is a standing reminder with no condition of its own, so it
          sat in the SAME red "attention" styling as a genuine validation error (the archive-clash
          refusal right below it, and the required-field errors elsewhere on this form). The moment
          the term set is actually FOUND — proven by the green confirmation under that field — this
          stops being useful advice and starts reading as a second active problem next to whatever
          real error the admin is looking at, even though nothing about the term set is wrong. It
          returns the instant the field goes blank or unresolved again (check.state leaves "found"),
          since that IS the case this banner exists for.
          ⚠⚠ ALSO HIDDEN ON "blank" — a second live report (2026-09-17): "the error box keeps
          showing the moment I go in, I haven't even trigger anything yet." An untouched, empty
          field is not a problem to flag red — it is the state every admin starts in, the same
          reasoning `showErrors`/`requiredFieldErrors` already apply elsewhere on this exact form
          (colouring or alarming a blank field before anyone has typed into it trains people to stop
          reading red). The banner now waits for either a genuine finding (malformed/notfound/
          unknown) or the field being actively checked, matching the hint right under the input two
          lines below, which already used this same "blank" exclusion. */}
      {check.state !== "found" && check.state !== "blank" && (
      <div style={{ ...s.msg, ...s.warn }}>
        The segment&apos;s <strong>term set must already exist</strong> in the term store, with its
        full structure of terms. This page does not create terms — reconciliation builds one folder
        level per term level, so the term set is what decides the shape.
      </div>
      )}

      <div style={s.card}>
        {/* "Project name" under Group Project, matching the term this project already uses
            elsewhere for that family's top-level identifier (CLAUDE.md: "Project Name" is what
            "Business Segment" is for a Head Office — the client-facing gotcha this label exists
            to head off is counting it as one of the two permissioned LEVELS below). Placeholder
            and hint follow the same split so an admin never sees a Business-Segment-shaped
            example while filling in a Project. The stored field and its behaviour are unchanged —
            this is display only. */}
        <label style={s.label}>{family === "Project" ? "Project name" : "Segment name"}</label>
        <input
          style={fieldError.label ? { ...s.input, ...s.inputBad } : s.input}
          value={label}
          onChange={(e) => setLabel(e.target.value)}
          placeholder={family === "Project" ? "Group-Led Project" : "Upstream Operations"}
        />
        {fieldError.label ? <div style={s.fieldErr}>{fieldError.label}</div> : undefined}
        <div style={s.hint}>
          What uploaders pick from the {family === "Project" ? "Project" : "Segment"} dropdown.
          {key ? ` Its configuration key will be ${key}.` : ""}
        </div>

        <label style={s.label}>Appears under</label>
        {(["BusinessSegment", "Project"] as const).map((f) => (
          <label key={f} style={{ fontSize: 13, marginRight: 16 }}>
            <input
              type="radio"
              name="family"
              checked={family === f}
              onChange={() => setFamily(f)}
              style={{ marginRight: 6 }}
            />
            {/* ⚠ LABEL ONLY, AND THE STORED `Category` VALUE STAYS THE LITERAL `Project`.
                Every consumer tests `Category === "Project"` — the upload form's own tab
                (`Form.tsx`), and the detail-panel labels on the approver's screen and My
                Submissions. Edit that cell in CRS Config to match this wording and the comparison
                stops matching: the segment moves under the Segment tab and both panels revert to
                saying "Segment", with nothing erroring and nothing logged.
                Wording is the client's: "Group-Led Project", capitalised, settled 2026-09-04;
                "Business Segment" standardised to "Segment", client QA item #53, 2026-09-13. */}
            {f === "BusinessSegment" ? "Segment" : "Group-Led Project"}
          </label>
        ))}

        <span style={s.labelRow}>
          <label style={{ ...s.label, marginTop: 0 }}>Term set ID</label>
          {/* Client QA item #57, 2026-09-13: "Include a Refresh button beside Open Term Store
              Management" — the same pairing `StructureManager.tsx` already has beside its own
              below-Unit level's Term set ID field. THE CLASSIC, SITE-LEVEL PAGE — the modern one
              (`SiteAdmin.aspx#/termStoreAdminCenter`) is the TENANT admin centre and answers "Access
              denied" to a site collection administrator (found 2026-08-30). Built from `siteUrl`,
              never hardcoded. New tab: this screen holds an unsaved-changes guard, and the admin is
              going there precisely to fetch a value to paste back into the field below. */}
          <a
            style={s.labelLink}
            href={`${siteUrl}/_layouts/15/termstoremanager.aspx`}
            onClick={openInNewTab}
            target="_blank"
            rel="noopener noreferrer"
          >
            Open the Term Store
          </a>
          {/* Offered only where the verdict came from the server — see `setCheckIsRecheckable`.
              Disabled WHILE checking rather than hidden, or hiding it moves the Term Store link as
              the row reflows out from under the cursor of an admin reaching for it. */}
          {setCheckIsRecheckable(check) && (
            <button
              type="button"
              style={check.state === "checking" ? s.recheckOff : s.recheckBtn}
              disabled={check.state === "checking"}
              title="Check this term set again — use it after editing the set in the Term Store"
              onClick={() => setRecheck((n) => n + 1)}
            >
              {check.state === "checking" ? "…" : "↻ Re-check"}
            </button>
          )}
        </span>
        <input
          style={fieldError.termSet ? { ...s.input, ...s.inputBad } : s.input}
          value={termSetGuid}
          onChange={(e) => setTermSetGuid(e.target.value)}
          placeholder="023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf"
        />
        {/* Shown BEFORE the term-store verdict, and it is the more urgent of the two: a set that
            resolves perfectly well is still wrong if another segment owns it. */}
        {fieldError.termSet ? <div style={s.fieldErr}>{fieldError.termSet}</div> : undefined}
        {check.state !== "blank" && (
          <div style={{ ...s.hint, ...setCheckStyle(check), fontWeight: 600 }}>
            {setCheckMessage(check)}
          </div>
        )}

        <label style={s.label}>Top folder name</label>
        <input
          style={fieldError.folder ? { ...s.input, ...s.inputBad } : s.input}
          value={stagingFolder}
          onChange={(e) => setStagingFolder(e.target.value)}
          placeholder="UPOPS"
        />
        {fieldError.folder ? <div style={s.fieldErr}>{fieldError.folder}</div> : undefined}
        <div style={s.hint}>
          The one folder every document in this segment sits under, in both libraries. Short and
          upper-case by convention. It cannot be shared with another segment.
        </div>
      </div>

      <div style={s.card}>
        <p style={s.h}>Levels that carry permissions</p>
        <div style={s.hint}>
          Name them yourself — one folder level each, and the{" "}
          <strong>deepest one holds the access</strong>, which is the level your groups are granted
          on. There must be exactly as many of these as the term set has levels of terms.
          <br />
          Head offices use <strong>Department</strong> then <strong>Unit</strong>; Upstream Ops uses{" "}
          <strong>Region</strong> then <strong>Estate/Mill</strong>; SDGI uses <strong>Refinery</strong>{" "}
          then <strong>Department</strong>. <strong>Exactly two are required</strong> — the business
          segment itself is the Top folder name above, not one of these.
        </div>
        {/* Empty is the starting state since 2026-08-15, so it has to read as "your turn" rather than
            as a list that failed to load. */}
        {tiers.length === 0 && (
          <div style={{ ...s.tierRow, background: "#fbfbfb", border: "1px dashed #d8d8d8", color: "#767676", fontSize: 12.5 }}>
            No levels yet — add the first one below. It becomes the top folder inside the segment.
          </div>
        )}
        <div style={{ marginTop: 10 }}>
          {/* The Add control is closed once two are named — see `MAX_PERMISSIONED_TIERS`. Removing
              one re-opens it, so this is a cap rather than a lock. */}
          {tiers.map((t, i) => (
            <div key={`${t}-${i}`} style={s.tierRow}>
              <span style={s.tierName}>{t}</span>
              <span style={s.tierMeta}>
                column {columnNameFor(t) || "—"}
                {i === tiers.length - 1 ? " · holds the access" : ""}
              </span>
              <span>
                <button style={s.iconBtn} disabled={i === 0} onClick={() => moveTier(i, -1)}>
                  ↑
                </button>
                <button
                  style={s.iconBtn}
                  disabled={i === tiers.length - 1}
                  onClick={() => moveTier(i, 1)}
                >
                  ↓
                </button>
                <button style={s.danger} onClick={() => setTiers(tiers.filter((_x, j) => j !== i))}>
                  Remove
                </button>
              </span>
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
          <input
            style={{ ...s.input, flex: "1 1 200px" }}
            value={newTier}
            onChange={(e) => setNewTier(e.target.value)}
            placeholder="Add a level, e.g. Region"
          />
          <button
            style={newTier.trim() && !tiersFull ? s.ghost : s.off}
            disabled={!newTier.trim() || tiersFull}
            onClick={addTier}
          >
            Add level
          </button>
        </div>
        {/* ⚠ THE REASON SITS BESIDE THE GREYED BUTTON. An unexplained disabled control reads as a
            broken page, and here the admin has just typed a name into the box next to it — without
            this they would retype it, or reload. It also names the misreading that produces the
            attempt: the business segment is the Top folder name, not one of these levels. */}
        {tiersFull && (
          <div style={{ ...s.hint, marginTop: 6, color: "#7a4f00" }}>
            Two levels is the maximum — every segment in this system has exactly two, and the
            business segment itself is the <strong>Top folder name</strong> above rather than a level
            here. Remove one to change it.
          </div>
        )}
      </div>

      <div style={s.card}>
        <p style={s.h}>Levels below that (shared, no permissions)</p>
        <div style={s.hint}>
          These are created when someone uploads, and inherit the permissions above, so they need
          no groups. Every segment starts with the standard pair; change them per segment
          afterwards on the <strong>Folder levels</strong> tab.
        </div>
        <div style={{ marginTop: 10 }}>
          {below.map((b) => (
            <div key={b.column} style={{ ...s.tierRow, ...s.tierLock }}>
              <span style={s.tierName}>{b.label}</span>
              <span style={s.tierMeta}>inherits · created on first use</span>
            </div>
          ))}
        </div>
        <div style={s.path}>{pathPreview}</div>
      </div>

      {/* `validateNewSegment` already refuses a segment with no permissioned level, but that message
          arrives after a click. Now the list starts empty, "no levels" is the state a first-time user
          begins in, so the reason belongs beside the button rather than behind it. */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <button
          style={busy || tiers.length === 0 ? s.off : s.btn}
          disabled={busy || tiers.length === 0}
          onClick={() => {
            create().catch(() => undefined); // create() reports its own failures into `result`
          }}
        >
          {busy ? "Working…" : "Create segment"}
        </button>
        {!busy && tiers.length === 0 && (
          <span style={{ fontSize: 12, color: "#8a4b00" }}>
            Name at least one level that carries permissions first.
          </span>
        )}
        {progress && <span style={{ fontSize: 12, color: "#605e5c" }}>{progress}</span>}
      </div>
      </>
      )}
    </div>
  );
}

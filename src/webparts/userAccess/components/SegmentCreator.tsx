import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, sanitizeFolderSegment } from "../../../shared/formModel";
import { effectiveOnDemandTiers } from "../../../shared/folderChain";
import { EVENT } from "../../../shared/auditLog";
// libraryUrlSegment, NOT libraryTitle, for anything that builds a PATH: the two differ
// ("Approval Document" vs "/ApprovalDocument") and the title fails silently in a URL — gotcha #12.
import { cachedListTitle, libraryTitle, libraryUrlSegment, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { ensureColumn } from "../../../shared/spColumns";
import {
  SegmentCounts,
  canOfferFolderDelete,
  confirmationMatches,
  deletionSummary,
  needsTypedConfirmation,
  survivorLines,
  unknownCounts,
} from "../../../shared/segmentDeletion";
import {
  buildSegmentLevels,
  columnNameFor,
  columnsForDraft,
  depthVerdict,
  ExistingSegment,
  isGuid,
  modeKeyFor,
  NewSegmentDraft,
  nextSortOrder,
  normalizeGuid,
  validateNewSegment,
} from "../../../shared/newSegment";

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

const DOCUMENTS_LIST_TITLE = "Documents";

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

const s: Record<string, React.CSSProperties> = {
  msg: { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.5 },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { background: "#fff4e5", border: "1px solid #f0d9b5", color: "#7a4f00" },
  ok: { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  card: { border: "1px solid #e1e1e1", borderRadius: 8, padding: "14px 16px", marginBottom: 12, background: "#fff" },
  label: { display: "block", fontSize: 12, fontWeight: 600, color: "#323130", margin: "14px 0 4px" },
  input: { width: "100%", boxSizing: "border-box", padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4 },
  hint: { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  btn: { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "8px 16px", fontSize: 13, cursor: "pointer" },
  ghost: { background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  off: { background: "#f3f2f1", color: "#a19f9d", border: "1px solid #e1dfdd", borderRadius: 4, padding: "8px 16px", fontSize: 13, cursor: "not-allowed" },
  iconBtn: { background: "#fff", border: "1px solid #c8c8c8", borderRadius: 4, padding: "3px 8px", fontSize: 12, cursor: "pointer", marginRight: 4 },
  danger: { background: "#fff", color: "#a4262c", border: "1px solid #e6b3b5", borderRadius: 4, padding: "4px 9px", fontSize: 12, cursor: "pointer" },
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
}

export default function SegmentCreator({
  context,
  siteUrl,
  onDirtyChange,
}: SegmentCreatorProps): React.ReactElement {
  const [existing, setExisting] = useState<ExistingSegment[]>([]);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [loaded, setLoaded] = useState(false);

  const [label, setLabel] = useState("");
  const [family, setFamily] = useState<"BusinessSegment" | "Project">("BusinessSegment");
  const [termSetGuid, setTermSetGuid] = useState("");
  const [stagingFolder, setStagingFolder] = useState("");
  const [tiers, setTiers] = useState<string[]>(["Department", "Unit"]);
  const [below, setBelow] = useState<Level[]>([]);
  const [newTier, setNewTier] = useState("");

  // Deletion. `delCounts === undefined` means "still counting" — distinct from a count that
  // finished and came back unknown, which is a state the dialog has to render differently.
  const [deleting, setDeleting] = useState<ExistingSegment | undefined>(undefined);
  const [delCounts, setDelCounts] = useState<SegmentCounts | undefined>(undefined);
  const [delFolders, setDelFolders] = useState(false);
  const [delTyped, setDelTyped] = useState("");
  const [delLog, setDelLog] = useState<string[]>([]);

  const [check, setCheck] = useState<SetCheck>({ state: "blank" });
  const [busy, setBusy] = useState(false);
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
  }, [termSetGuid]);

  /* ── Dirty tracking ────────────────────────────────────────────────────────── */

  const dirty =
    label.trim() !== "" ||
    termSetGuid.trim() !== "" ||
    stagingFolder.trim() !== "" ||
    tiers.join("|") !== "Department|Unit";

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

  /* ── Create ────────────────────────────────────────────────────────────────── */

  const create = async (): Promise<void> => {
    setBusy(true);
    setResult(undefined);
    setProgress("");
    try {
      const d = draft();

      // 1. VALIDATE FIRST. A rejected draft must leave nothing behind — no column, no row.
      const errors = validateNewSegment(d, existing);
      if (errors.length > 0) {
        setResult({ ok: false, text: errors.join(" ") });
        return;
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
        for (const lib of [libraryTitle(), DOCUMENTS_LIST_TITLE]) {
          setProgress(`Creating ${col.internal} in ${lib}…`);
          if (await ensureColumn(context.spHttpClient, siteUrl, lib, col.internal, col.display)) {
            created.push(`${col.internal} (${lib})`);
          }
        }
      }

      // 4. The mode row, last.
      setProgress("Writing the configuration row…");
      const key = modeKeyFor(d.label);
      const folder = sanitizeFolderSegment(d.stagingFolder).trim();
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
      setTiers(["Department", "Unit"]);
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
     empty being deleted. */

  /** Every subfolder path beneath `root`, depth-first. Throws — the caller reports `unknown`. */
  const walkFolders = async (lib: string, root: string): Promise<string[]> => {
    const found: string[] = [];
    const queue = [`${siteUrl.replace(/^https?:\/\/[^/]+/, "")}/${lib}/${root}`];
    while (queue.length > 0 && found.length < 5000) {
      const here = queue.shift() as string;
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Folders?@f='${encodeURIComponent(here)}'&$select=ServerRelativeUrl,Name`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      // 404 = this branch does not exist in this library, which is not an error: a segment can
      // have folders in one library and not the other.
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
  };

  /** Files directly in one folder. Throws — the caller reports `unknown`. */
  const countFiles = async (folderUrl: string): Promise<number> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Files/$count?@f='${encodeURIComponent(folderUrl)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (res.status === 404) return 0;
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return Number(await res.text()) || 0;
  };

  /**
   * The Group Map rows carrying this segment.
   *
   * Matched on the term-set GUID, which is what a folder-scope row stores in `Segment`.
   */
  const loadSegmentGroupMapRows = async (seg: ExistingSegment): Promise<number[]> => {
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
  };

  /**
   * Count what a segment holds, across BOTH libraries, plus its Group Map rows.
   *
   * Any failure makes the whole count `unknown` rather than a partial number. A partial count is
   * worse than none here: it would read as authoritative and could show "0 documents" for a
   * segment whose second library simply did not answer.
   */
  const countSegment = async (seg: ExistingSegment): Promise<SegmentCounts> => {
    const root = (seg.stagingFolder ?? "").trim();
    if (!root) {
      return unknownCounts("this segment has no top folder recorded, so its folders cannot be found");
    }
    let folders = 0;
    let documents = 0;
    try {
      for (const lib of [libraryUrlSegment(), DOCUMENTS_LIST_TITLE]) {
        const rootUrl = `${siteUrl.replace(/^https?:\/\/[^/]+/, "")}/${lib}/${root}`;
        documents += await countFiles(rootUrl);
        const subs = await walkFolders(lib, root);
        folders += subs.length;
        for (const f of subs) documents += await countFiles(f);
      }
    } catch (e) {
      return unknownCounts(`the folders could not be read (${(e as Error).message})`);
    }

    // The Group Map read is separate and NOT fatal to the count: a segment can have unreadable
    // folders and readable rows, or the reverse. An unreadable Group Map is reported as unknown
    // too, because deleting rows we could not see is the same class of blind act.
    let groupMapRows = 0;
    try {
      const rows = await loadSegmentGroupMapRows(seg);
      groupMapRows = rows.length;
    } catch (e) {
      return unknownCounts(`the ${cachedListTitle(LIST_SUFFIX.groupMap)} list could not be read (${(e as Error).message})`);
    }
    return { state: "counted", folders, documents, groupMapRows };
  };

  const deleteItem = async (listTitle: string, itemId: number): Promise<void> => {
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
  };

  /** Recycle (not purge) a folder, so it is restorable for 93 days — see the dialog's promise. */
  const recycleFolder = async (lib: string, root: string): Promise<boolean> => {
    const url = `${siteUrl.replace(/^https?:\/\/[^/]+/, "")}/${lib}/${root}`;
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Recycle()?@f='${encodeURIComponent(url)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (res.status === 404) return false;
    if (!res.ok) throw new Error(`${lib}: HTTP ${res.status}`);
    return true;
  };

  const openDelete = (seg: ExistingSegment): void => {
    setDeleting(seg);
    setDelCounts(undefined);
    setDelFolders(false);
    setDelTyped("");
    setDelLog([]);
    countSegment(seg)
      .then(setDelCounts)
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
        rowIds = await loadSegmentGroupMapRows(seg);
      } catch (e) {
        lines.push(`Could not read the folder-access mappings (${(e as Error).message}) — none were removed.`);
        ok = false;
      }
      for (const id of rowIds) {
        try {
          await deleteItem(cachedListTitle(LIST_SUFFIX.groupMap), id);
          rowsRemoved++;
        } catch {
          ok = false;
        }
      }
      if (rowIds.length > 0) {
        lines.push(`Folder-access mappings removed: ${rowsRemoved} of ${rowIds.length}.`);
      }

      // 2. The mode row. If THIS fails, stop: a run that removed the grants but left the segment
      //    live is a segment whose uploaders have quietly lost their access.
      if (seg.itemId === undefined) throw new Error("this segment's configuration row has no id, so it cannot be deleted");
      await deleteItem(cachedListTitle(LIST_SUFFIX.config), seg.itemId);
      lines.push("The segment is no longer offered in the upload form, and reconciliation will not walk it.");

      // 3. Folders, only if asked. After the row, never before — see the spec's D6.
      if (alsoFolders) {
        const root = (seg.stagingFolder ?? "").trim();
        for (const lib of [libraryUrlSegment(), DOCUMENTS_LIST_TITLE]) {
          try {
            const went = await recycleFolder(lib, root);
            lines.push(went ? `${lib}/${root} moved to the recycle bin.` : `${lib}/${root} did not exist.`);
          } catch (e) {
            lines.push(`Could not delete ${lib}/${root} — ${(e as Error).message}`);
            ok = false;
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
              try { await deleteItem(cachedListTitle(LIST_SUFFIX.folderMap), r.Id); gone++; } catch { ok = false; }
            }
            if (mine.length > 0) lines.push(`Folder Map rows removed: ${gone} of ${mine.length}.`);
          }
        } catch {
          lines.push("The Folder Map rows could not be tidied up; reconciliation will report them.");
          ok = false;
        }
      }

      lines.push(...survivorLines(alsoFolders));
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
            ? `Counted before deleting: ${delCounts.folders} folder(s), ${delCounts.documents} document(s), ${delCounts.groupMapRows} mapping(s)`
            : `Counts were UNKNOWN before deleting (${delCounts.reason ?? "unreadable"})`,
          ...lines,
        ],
      }).catch(() => undefined);

      setDeleting(undefined);
      await reload();
    } catch (e) {
      setDelLog([...lines, `Stopped: ${(e as Error).message}`]);
      setResult({ ok: false, text: `"${seg.label}" was NOT deleted — ${(e as Error).message}` });
    } finally {
      setBusy(false);
    }
  };

  /* ── Render ────────────────────────────────────────────────────────────────── */

  const key = modeKeyFor(label);
  const folderPreview = sanitizeFolderSegment(stagingFolder).trim();
  const pathPreview =
    `/${folderPreview || "TOPFOLDER"}/` +
    [...tiers.map((t) => t || "?"), ...below.map((b) => b.label)].map((n) => `<${n}>`).join("/");

  if (!loaded) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading&hellip;</p>;

  const typedOk =
    delCounts !== undefined &&
    deleting !== undefined &&
    (!needsTypedConfirmation(delCounts, delFolders) || confirmationMatches(delTyped, deleting.label));

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
      {existing.length > 0 && (
        <div style={s.card}>
          <p style={s.h}>Segments on this site ({existing.length})</p>
          {existing.map((seg) => (
            <div key={seg.key} style={{ ...s.tierRow, background: "#fff", border: "1px solid #f0f0f0" }}>
              <span style={s.tierName}>{seg.label || seg.key}</span>
              <span style={s.tierMeta}>
                top folder <strong>{seg.stagingFolder || "—"}</strong>
              </span>
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
            Deleting a segment stops it being offered and removes its folder-access mappings. It
            does not delete any document, and never deletes a column.
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
                    {delCounts.documents === 1 ? "" : "s"} across both libraries, and{" "}
                    <strong>{delCounts.groupMapRows}</strong> folder-access mapping
                    {delCounts.groupMapRows === 1 ? "" : "s"}.
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
                  {survivorLines(delFolders && canOfferFolderDelete(delCounts)).map((line, i) => (
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
                    {busy ? "Deleting…" : "Delete segment"}
                  </button>
                  <button style={s.ghost} disabled={busy} onClick={() => setDeleting(undefined)}>
                    Cancel
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {result && (
        <div style={{ ...s.msg, ...(result.ok ? s.ok : s.err) }}>
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

      <div style={{ ...s.msg, ...s.warn }}>
        The segment&apos;s <strong>term set must already exist</strong> in the term store, with its
        full structure of terms. This page does not create terms — reconciliation builds one folder
        level per term level, so the term set is what decides the shape.
      </div>

      <div style={s.card}>
        <label style={s.label}>Segment name</label>
        <input
          style={s.input}
          value={label}
          onChange={(e) => setLabel(e.target.value)}
          placeholder="Upstream Operations"
        />
        <div style={s.hint}>
          What uploaders pick from the Segment dropdown.
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
            {f === "BusinessSegment" ? "Business Segment" : "Project"}
          </label>
        ))}

        <label style={s.label}>Term set ID</label>
        <input
          style={s.input}
          value={termSetGuid}
          onChange={(e) => setTermSetGuid(e.target.value)}
          placeholder="023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf"
        />
        {check.state !== "blank" && (
          <div style={{ ...s.hint, ...setCheckStyle(check), fontWeight: 600 }}>
            {setCheckMessage(check)}
          </div>
        )}

        <label style={s.label}>Top folder name</label>
        <input
          style={s.input}
          value={stagingFolder}
          onChange={(e) => setStagingFolder(e.target.value)}
          placeholder="UPOPS"
        />
        <div style={s.hint}>
          The one folder every document in this segment sits under, in both libraries. Short and
          upper-case by convention. It cannot be shared with another segment.
        </div>
      </div>

      <div style={s.card}>
        <p style={s.h}>Levels that carry permissions</p>
        <div style={s.hint}>
          One folder level each, and the <strong>deepest one holds the access</strong> — that is the
          level your groups are granted on. There must be exactly as many of these as the term set
          has levels of terms. Head offices use Department then Unit; Upstream Ops uses Region then
          Estate/Mill; I&amp;T uses a single level.
        </div>
        <div style={{ marginTop: 10 }}>
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
            style={newTier.trim() ? s.ghost : s.off}
            disabled={!newTier.trim()}
            onClick={addTier}
          >
            Add level
          </button>
        </div>
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

      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <button
          style={busy ? s.off : s.btn}
          disabled={busy}
          onClick={() => {
            create().catch(() => undefined); // create() reports its own failures into `result`
          }}
        >
          {busy ? "Working…" : "Create segment"}
        </button>
        {progress && <span style={{ fontSize: 12, color: "#605e5c" }}>{progress}</span>}
      </div>
    </div>
  );
}

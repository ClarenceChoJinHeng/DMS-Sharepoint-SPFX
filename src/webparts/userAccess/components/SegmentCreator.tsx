import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, sanitizeFolderSegment } from "../../../shared/formModel";
import { effectiveOnDemandTiers } from "../../../shared/folderChain";
import { cachedListTitle, libraryTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { ensureColumn } from "../../../shared/spColumns";
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

      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${config}')/items` +
          `?$select=Title,ModeLabel,StagingFolder,SortOrder&$filter=ConfigType eq 'mode'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) throw new Error(`the configuration list returned HTTP ${res.status}`);
      const rows = ((await res.json()).value ?? []) as Array<{
        Title?: string;
        ModeLabel?: string;
        StagingFolder?: string;
        SortOrder?: number;
      }>;
      setExisting(
        rows.map((r) => ({
          key: (r.Title ?? "").trim(),
          label: (r.ModeLabel ?? r.Title ?? "").trim(),
          stagingFolder: (r.StagingFolder ?? "").trim(),
          sortOrder: r.SortOrder,
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

  /* ── Render ────────────────────────────────────────────────────────────────── */

  const key = modeKeyFor(label);
  const folderPreview = sanitizeFolderSegment(stagingFolder).trim();
  const pathPreview =
    `/${folderPreview || "TOPFOLDER"}/` +
    [...tiers.map((t) => t || "?"), ...below.map((b) => b.label)].map((n) => `<${n}>`).join("/");

  if (!loaded) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading&hellip;</p>;

  return (
    <div>
      {loadError && (
        <div style={{ ...s.msg, ...s.err }}>
          Could not read the existing segments — {loadError}. Adding one now risks a duplicate name
          or a shared top folder, so fix that first.
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

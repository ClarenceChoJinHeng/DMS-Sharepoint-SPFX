import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, parseLevels, sanitizeFolderSegment } from "../../../shared/formModel";
import { splitChain, validateChain } from "../../../shared/folderChain";
import { cachedListTitle, libraryTitle, libraryUrlSegment, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";

/**
 * Folder Structure — add, reorder and remove the folder levels BENEATH Unit.
 *
 * Piece 2 of `2026-08-06-configurable-folder-structure-chain-design.md`, specced in
 * `2026-08-10-structure-manager-ui-design.md`. Piece 1 made the structure data-driven;
 * it lives as `Levels` JSON on a DMS Config `mode` row, and the client cannot edit JSON —
 * which is the whole reason the feature exists. This is the editor that closes that.
 *
 * THE JSON IS NEVER SHOWN. An admin edits a list of levels and sees the resulting path.
 *
 * Scope, deliberately: below-Unit levels only. The permissioned tiers (Segment, Department,
 * Unit) carry folder ACLs and are set at provisioning time; renaming one renames live ACL'd
 * folders and reordering one asks reconciliation to walk a tree that does not exist. Neither
 * is a config edit, so they render locked.
 */

const DOCUMENTS_LIST_TITLE = "Documents";
/** Documents' URL segment differs from its title, like the approval library's does. */
const DOCUMENTS_URL_SEGMENT = "Shared Documents";

/** A DMS Config `mode` row, reduced to what this screen edits. */
interface SegmentRow {
  id: number;
  key: string;           // Title, e.g. "mode_gho"
  label: string;         // ModeLabel, e.g. "Group Head Office"
  stagingFolder: string; // top folder name, e.g. "GHO"
  chain: Level[];
  /** Parsed-chain problem, if any — shown instead of letting them edit a broken row. */
  chainError?: string;
  /** Whether documents are already filed under this segment. Drives the warning. */
  hasDocuments?: boolean;
}

/** The in-progress "add a level" form. */
interface DraftTier {
  label: string;
  termSetGuid: string; // blank = values come from the level above
  position: number;    // index within the below-Unit list
}

const s: Record<string, React.CSSProperties> = {
  msg:       { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.5 },
  err:       { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn:      { background: "#fff4e5", border: "1px solid #f0d9b5", color: "#7a4f00" },
  ok:        { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  card:      { border: "1px solid #e1e1e1", borderRadius: 8, padding: "14px 16px", marginBottom: 12, background: "#fff" },
  segRow:    { display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 16, flexWrap: "wrap" },
  segName:   { fontSize: 15, fontWeight: 600, color: "#1b1b1b" },
  path:      { fontSize: 12, color: "#605e5c", fontFamily: "Consolas, monospace", marginTop: 4, wordBreak: "break-all" },
  badge:     { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em", padding: "2px 8px", borderRadius: 10 },
  badgeUsed: { background: "#fff4e5", color: "#7a4f00" },
  badgeFree: { background: "#f1f8f4", color: "#0f6c3f" },
  btn:       { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  ghost:     { background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  danger:    { background: "#fff", color: "#a4262c", border: "1px solid #e6b3b5", borderRadius: 4, padding: "4px 9px", fontSize: 12, cursor: "pointer" },
  iconBtn:   { background: "#fff", border: "1px solid #c8c8c8", borderRadius: 4, padding: "3px 8px", fontSize: 12, cursor: "pointer", marginRight: 4 },
  off:       { background: "#f3f2f1", color: "#a19f9d", border: "1px solid #e1dfdd", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "not-allowed" },
  tierRow:   { display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 6, background: "#fafafa", marginBottom: 6, fontSize: 13, flexWrap: "wrap" },
  tierLock:  { background: "#f3f2f1", color: "#605e5c" },
  tierName:  { fontWeight: 600, flex: "0 0 150px" },
  tierMeta:  { fontSize: 11, color: "#8a8886", flex: "1 1 160px" },
  label:     { display: "block", fontSize: 12, fontWeight: 600, color: "#323130", margin: "14px 0 4px" },
  input:     { width: "100%", boxSizing: "border-box", padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4 },
  hint:      { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  modalBg:   { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 },
  modal:     { background: "#fff", borderRadius: 8, padding: 24, maxWidth: 520, width: "100%", maxHeight: "80vh", overflowY: "auto" },
};

export interface StructureManagerProps {
  context: WebPartContext;
  siteUrl: string;
}

export default function StructureManager({ context, siteUrl }: StructureManagerProps): React.ReactElement {
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [segments, setSegments] = useState<SegmentRow[]>([]);
  const [editing, setEditing] = useState<string | undefined>(undefined);
  const [draft, setDraft] = useState<Level[]>([]);
  const [adding, setAdding] = useState<DraftTier | undefined>(undefined);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<{ text: string; ok: boolean } | undefined>(undefined);
  const [confirmSave, setConfirmSave] = useState(false);
  const [confirmText, setConfirmText] = useState("");

  const editingSegment = (): SegmentRow | undefined => segments.filter((x) => x.key === editing)[0];

  /* ---------- Load ------------------------------------------------------- */

  const loadSegments = async (): Promise<SegmentRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
        `/items?$select=Id,Title,ModeLabel,StagingFolder,Levels` +
        `&$filter=ConfigType eq 'mode'&$top=200`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`the configuration list returned HTTP ${res.status}`);
    const data = await res.json();
    const rows = ((data.value ?? []) as Array<{
      Id: number; Title?: string; ModeLabel?: string; StagingFolder?: string; Levels?: string;
    }>).map((r) => {
      const chain = parseLevels(r.Levels ?? "");
      const err = validateChain(chain);
      const row: SegmentRow = {
        id: r.Id,
        key: (r.Title ?? "").trim(),
        label: (r.ModeLabel ?? r.Title ?? "").trim(),
        stagingFolder: (r.StagingFolder ?? "").trim(),
        chain,
      };
      // A row whose chain does not validate is shown but NOT editable: saving would
      // rewrite a structure nobody has seen, and the fix may lie outside this screen.
      if (err) row.chainError = err.message;
      return row;
    });
    return rows.sort((a, b) => a.label.localeCompare(b.label));
  };

  /**
   * Whether ANY document is already filed under this segment's top folder.
   *
   * Existence, not a count — it is all the warning needs, and one CAML query with
   * RowLimit 1 answers it whatever the library size. A count would mean paging through
   * a library the Year × Document Type grid can push past the 5,000-item threshold.
   *
   * Files only (`FSObjType = 0`). Folders exist from the moment reconciliation runs, so
   * counting them would report every provisioned segment as "in use" before anyone had
   * uploaded anything.
   *
   * Checks the approval library too: a segment with nothing approved can still hold
   * pending files whose paths a structure change would leave behind.
   */
  const segmentHasDocuments = async (folder: string): Promise<boolean> => {
    if (!folder) return false;
    const webPath = new URL(siteUrl).pathname.replace(/\/$/, "");
    const libs: Array<[string, string]> = [
      [libraryTitle(), libraryUrlSegment()],
      [DOCUMENTS_LIST_TITLE, DOCUMENTS_URL_SEGMENT],
    ];
    for (const [title, segment] of libs) {
      const res: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/GetItems`,
        SPHttpClient.configurations.v1,
        {
          headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" },
          body: JSON.stringify({
            query: {
              ViewXml:
                `<View Scope="RecursiveAll"><Query><Where>` +
                `<Eq><FieldRef Name="FSObjType" /><Value Type="Integer">0</Value></Eq>` +
                `</Where></Query><RowLimit>1</RowLimit></View>`,
              FolderServerRelativeUrl: `${webPath}/${segment}/${folder}`,
            },
          }),
        },
      );
      if (!res.ok) continue;
      const data = await res.json();
      if (((data.value ?? []) as unknown[]).length > 0) return true;
    }
    return false;
  };

  useEffect(() => {
    let cancelled = false;
    // Names FIRST — every read below goes through cachedListTitle, and an unprimed cache
    // resolves to the legacy DMS titles, which 404 on a renamed site and would present
    // as "the config could not be read" rather than "the list is called something else".
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      .then(() => loadSegments())
      .then(async (rows) => {
        if (cancelled) return;
        setSegments(rows);
        setLoading(false);
        // Per-segment document checks are slower than the list read, so fill them in
        // afterwards rather than holding the whole page on them.
        for (const row of rows) {
          let used = true;
          try {
            used = await segmentHasDocuments(row.stagingFolder);
          } catch {
            // Cannot prove it is empty, so assume in use. A needless warning costs one
            // click; a missing one costs a split folder tree.
            used = true;
          }
          if (cancelled) return;
          setSegments((prev) => prev.map((p) => (p.key === row.key ? { ...p, hasDocuments: used } : p)));
        }
      })
      .catch((e: Error) => {
        if (cancelled) return;
        setLoadError(e.message);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  /**
   * The folder path a chain produces, with level names in brackets where a real value
   * goes. The only representation of the change an admin can reason about — which is why
   * it sits under every editor and the JSON sits nowhere.
   */
  const pathPreview = (seg: SegmentRow, chain: Level[]): string =>
    [seg.stagingFolder || seg.label, ...chain.map((l) => `[${l.label}]`)].join(" / ");

  /* ---------- Editing --------------------------------------------------- */

  const startEdit = (seg: SegmentRow): void => {
    setEditing(seg.key);
    setDraft(seg.chain.map((l) => ({ ...l })));
    setAdding(undefined);
    setResult(undefined);
  };

  const cancelEdit = (): void => {
    setEditing(undefined);
    setDraft([]);
    setAdding(undefined);
    setConfirmSave(false);
    setConfirmText("");
  };

  const permissionedCount = (chain: Level[]): number => splitChain(chain).permissioned.length;

  /** Move a below-Unit level. Bounded to the below-Unit region — the prefix cannot move. */
  const moveTier = (index: number, delta: number): void => {
    const first = permissionedCount(draft);
    const to = index + delta;
    if (to < first || to >= draft.length) return;
    const next = draft.slice();
    const item = next[index];
    next.splice(index, 1);
    next.splice(to, 0, item);
    setDraft(next);
  };

  const removeTier = (index: number): void => setDraft(draft.filter((_l, i) => i !== index));

  /**
   * Derive the column internal name from the level name — never typed by the admin.
   *
   * A SharePoint column's internal name is fixed at creation, permanently, from the title
   * it is created with. Letting an untrained admin type it offers a decision that cannot
   * be undone and whose consequence is invisible until a metadata write silently fails.
   * So strip everything SharePoint would encode, create under that, then set the display
   * title to what they typed.
   */
  const columnNameFor = (label: string): string =>
    sanitizeFolderSegment(label).replace(/[^A-Za-z0-9]/g, "");

  const addTier = (): void => {
    if (!adding) return;
    const label = adding.label.trim();
    const col = columnNameFor(label);
    if (!label || !col) return;
    const tier: Level = {
      label,
      column: col,
      labelCol: col,
      tidCol: `${col}Tid`,
      permissioned: false,
    };
    // Blank term set = values come from the level above (how each Unit gets its own
    // SubUnits). Presence of termSet is the discriminator; there is no extra flag.
    const set = adding.termSetGuid.trim();
    if (set) tier.termSet = set;
    const next = draft.slice();
    next.splice(permissionedCount(draft) + adding.position, 0, tier);
    setDraft(next);
    setAdding(undefined);
  };

  /* ---------- Save ------------------------------------------------------ */

  /** Create a text column if absent. Idempotent; returns whether it created one. */
  const ensureTextColumn = async (
    listTitle: string,
    internalName: string,
    displayName: string,
  ): Promise<boolean> => {
    const check: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields?$select=InternalName&$top=500`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!check.ok) throw new Error(`${listTitle}: could not read its columns (HTTP ${check.status})`);
    const data = await check.json();
    const exists = ((data.value ?? []) as Array<{ InternalName?: string }>)
      .filter((f) => (f.InternalName ?? "").toLowerCase() === internalName.toLowerCase()).length > 0;
    if (exists) return false;

    // CreateFieldAsXml, not the /fields collection: it is the only form that sets the
    // internal name (Name/StaticName) independently of the display name, which is what
    // keeps the internal name free of _x0020_ encoding.
    //
    // Options 8 = AddToAllContentTypes. Deliberately NOT 12: that also adds the column
    // to the default VIEW, silently reshaping every library view the client arranged.
    const xml =
      `<Field Type="Text" DisplayName="${internalName}" Name="${internalName}" ` +
      `StaticName="${internalName}" MaxLength="255" />`;
    const create: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/CreateFieldAsXml`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" },
        body: JSON.stringify({ parameters: { SchemaXml: xml, Options: 8 } }),
      },
    );
    if (!create.ok) {
      const body = await create.text().catch(() => "");
      throw new Error(
        `${listTitle}: could not create the "${internalName}" column (HTTP ${create.status}). ${body.slice(0, 180)}`,
      );
    }
    // Rename to the human title afterwards. The internal name is already frozen by the
    // create above, so this only changes what is displayed.
    if (displayName !== internalName) {
      await context.spHttpClient
        .post(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')` +
            `/fields/getbyinternalnameortitle('${encodeURIComponent(internalName)}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json",
              "X-HTTP-Method": "MERGE",
              "IF-MATCH": "*",
            },
            body: JSON.stringify({ Title: displayName }),
          },
        )
        .catch(() => undefined); // cosmetic — never fail a save over a display name
    }
    return true;
  };

  const saveStructure = async (): Promise<void> => {
    const seg = editingSegment();
    if (!seg) return;
    setBusy(true);
    setResult(undefined);
    try {
      // 1. VALIDATE FIRST, before a single column is created. A rejected chain must
      //    leave nothing behind.
      const err = validateChain(draft);
      if (err) throw new Error(err.message);

      // 2. Columns, in BOTH libraries. Order matters: a column with no chain is a
      //    harmless orphan, whereas a chain naming a missing column breaks every upload
      //    in the segment — and validateUpdateListItem returns HTTP 200 with
      //    HasException, so that failure is not even loud.
      const created: string[] = [];
      for (const tier of splitChain(draft).onDemand) {
        const cols: Array<[string, string]> = [];
        if (tier.labelCol) cols.push([tier.labelCol, tier.label]);
        if (tier.tidCol) cols.push([tier.tidCol, `${tier.label} ID`]);
        for (const [internal, display] of cols) {
          for (const lib of [libraryTitle(), DOCUMENTS_LIST_TITLE]) {
            if (await ensureTextColumn(lib, internal, display)) created.push(`${internal} (${lib})`);
          }
        }
      }

      // 3. Write the chain.
      const write: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items(${seg.id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify({ Levels: JSON.stringify(draft) }),
        },
      );
      if (!write.ok) {
        throw new Error(
          `The columns were created but the structure could not be saved (HTTP ${write.status}). ` +
            `Those columns are harmless and will be reused when you try again.`,
        );
      }

      setSegments((prev) =>
        prev.map((p) => (p.key === seg.key ? { ...p, chain: draft.map((l) => ({ ...l })) } : p)),
      );
      setResult({
        ok: true,
        text:
          `Structure saved for ${seg.label}.` +
          (created.length > 0 ? ` Created ${created.length} column(s): ${created.join(", ")}.` : "") +
          ` Uploads use the new shape as soon as people reload the form — no reconciliation run` +
          ` is needed, because folders below Unit are created on demand and inherit the unit's` +
          ` permissions.`,
      });
      cancelEdit();
    } catch (e) {
      setResult({ ok: false, text: (e as Error).message });
    } finally {
      setBusy(false);
      setConfirmSave(false);
      setConfirmText("");
    }
  };

  const onSaveClicked = (): void => {
    const seg = editingSegment();
    if (!seg) return;
    // Warn only where it matters. On an empty segment this is a free change, and a modal
    // there would train them to click through the one that counts.
    if (seg.hasDocuments) setConfirmSave(true);
    else saveStructure().catch(() => undefined);
  };

  /* ---------- Render ---------------------------------------------------- */

  if (loading) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading folder structure&hellip;</p>;

  if (loadError !== undefined) {
    return (
      <div style={{ ...s.msg, ...s.err }}>
        Could not read the configuration: {loadError}. Nothing is safe to edit without it — this
        screen will not guess at a structure.
      </div>
    );
  }

  const seg = editingSegment();

  if (seg) {
    const first = permissionedCount(draft);
    const onDemand = draft.slice(first);
    return (
      <>
        {result !== undefined && <div style={{ ...s.msg, ...(result.ok ? s.ok : s.err) }}>{result.text}</div>}

        <div style={s.card}>
          <div style={s.segName}>{seg.label}</div>
          <div style={s.path}>{pathPreview(seg, draft)}</div>
        </div>

        <p style={s.label}>Fixed levels — these carry folder permissions and cannot be changed here</p>
        {draft.slice(0, first).map((l, i) => (
          <div style={{ ...s.tierRow, ...s.tierLock }} key={`p${i}`}>
            <span style={{ flex: "0 0 16px" }} aria-hidden="true">&#128274;</span>
            <span style={s.tierName}>{l.label}</span>
            <span style={s.tierMeta}>
              column: {l.labelCol ?? l.column}
              {l.tidCol !== undefined ? ` + ${l.tidCol}` : ""}
            </span>
          </div>
        ))}

        <p style={s.label}>Folders below Unit — these inherit the unit&rsquo;s permissions</p>
        {onDemand.length === 0 && (
          <p style={{ fontSize: 12, color: "#8a8886", margin: "0 0 8px", lineHeight: 1.5 }}>
            None configured, so uploads use the built-in <strong>Year</strong> then{" "}
            <strong>Document Type</strong> folders. Adding a level here replaces that with the list
            you build — so include Year and Document Type unless you mean to drop them.
          </p>
        )}
        {onDemand.map((l, k) => {
          const i = first + k;
          return (
            <div style={s.tierRow} key={`d${i}`}>
              <span style={s.tierName}>{l.label}</span>
              <span style={s.tierMeta}>
                {l.termSet !== undefined ? "own list of values" : "values come from the level above"}
                {" · column: "}
                {l.labelCol ?? l.column}
              </span>
              <span style={{ flex: "0 0 auto" }}>
                <button style={s.iconBtn} disabled={k === 0} onClick={() => moveTier(i, -1)} title="Move up">
                  &#8593;
                </button>
                <button
                  style={s.iconBtn}
                  disabled={k === onDemand.length - 1}
                  onClick={() => moveTier(i, 1)}
                  title="Move down"
                >
                  &#8595;
                </button>
                <button style={s.danger} onClick={() => removeTier(i)}>Remove</button>
              </span>
            </div>
          );
        })}

        {adding === undefined ? (
          <button
            style={s.ghost}
            onClick={() => setAdding({ label: "", termSetGuid: "", position: onDemand.length })}
          >
            + Add folder level
          </button>
        ) : (
          <div style={{ ...s.card, marginTop: 12 }}>
            <label style={s.label} htmlFor="sm-label">Folder level name</label>
            <input
              id="sm-label"
              style={s.input}
              value={adding.label}
              onChange={(e) => setAdding({ ...adding, label: e.target.value })}
              placeholder="e.g. SubUnit"
            />
            <p style={s.hint}>
              Shown above the dropdown on the upload form. Its column will be{" "}
              <strong>{columnNameFor(adding.label) || "—"}</strong>, created automatically in both
              libraries.
            </p>

            <label style={s.label} htmlFor="sm-set">Where its values come from</label>
            <input
              id="sm-set"
              style={s.input}
              value={adding.termSetGuid}
              onChange={(e) => setAdding({ ...adding, termSetGuid: e.target.value })}
              placeholder="Paste a term set ID, or leave blank"
            />
            <p style={s.hint}>
              Leave <strong>blank</strong> if the values sit under each Unit in the term store — every
              unit then offers its own. Paste a <strong>term set ID</strong> if all units pick from
              one shared list.
            </p>

            <label style={s.label} htmlFor="sm-pos">Position</label>
            <select
              id="sm-pos"
              style={s.input}
              value={String(adding.position)}
              onChange={(e) => setAdding({ ...adding, position: Number(e.target.value) })}
            >
              {Array.from({ length: onDemand.length + 1 }, (_x, p) => (
                <option key={p} value={String(p)}>
                  {onDemand.length === 0
                    ? "First level below Unit"
                    : p === 0
                      ? `Before ${onDemand[0].label}`
                      : p === onDemand.length
                        ? `After ${onDemand[onDemand.length - 1].label}`
                        : `Between ${onDemand[p - 1].label} and ${onDemand[p].label}`}
                </option>
              ))}
            </select>
            <p style={s.hint}>
              Position decides how many folders exist. Near the top, one folder per unit; at the
              bottom, one for every Year and Document Type combination.
            </p>

            <div style={{ marginTop: 14 }}>
              <button
                style={adding.label.trim() && columnNameFor(adding.label) ? s.btn : s.off}
                disabled={!adding.label.trim() || !columnNameFor(adding.label)}
                onClick={addTier}
              >
                Add
              </button>{" "}
              <button style={s.ghost} onClick={() => setAdding(undefined)}>Cancel</button>
            </div>
          </div>
        )}

        <div style={{ marginTop: 24, borderTop: "1px solid #edebe9", paddingTop: 16 }}>
          <button style={busy ? s.off : s.btn} disabled={busy} onClick={onSaveClicked}>
            {busy ? "Saving…" : "Save structure"}
          </button>{" "}
          <button style={s.ghost} disabled={busy} onClick={cancelEdit}>Cancel</button>
        </div>

        {confirmSave && (
          <div style={s.modalBg} role="dialog" aria-modal="true">
            <div style={s.modal}>
              <h3 style={{ margin: "0 0 12px", fontSize: 17 }}>This changes where new documents are filed</h3>
              <p style={{ fontSize: 13, lineHeight: 1.6 }}>
                <strong>{seg.label}</strong> already has documents filed.
              </p>
              <p style={{ fontSize: 13, lineHeight: 1.6 }}>
                Those documents stay exactly where they are. New uploads will use the new shape, so
                both folder layouts will exist side by side — and nothing moves the older documents
                automatically.
              </p>
              <p style={{ fontSize: 13, lineHeight: 1.6 }}>
                This is safe on a segment nobody is using yet. On one already in use, plan it
                deliberately: moving existing documents is a separate job.
              </p>
              <label style={s.label} htmlFor="sm-confirm">Type CHANGE to confirm</label>
              <input
                id="sm-confirm"
                style={s.input}
                value={confirmText}
                onChange={(e) => setConfirmText(e.target.value)}
              />
              <div style={{ marginTop: 16, textAlign: "right" }}>
                <button
                  style={s.ghost}
                  onClick={() => {
                    setConfirmSave(false);
                    setConfirmText("");
                  }}
                >
                  Cancel
                </button>{" "}
                <button
                  style={confirmText.trim().toUpperCase() === "CHANGE" ? s.btn : s.off}
                  disabled={confirmText.trim().toUpperCase() !== "CHANGE"}
                  onClick={() => {
                    saveStructure().catch(() => undefined);
                  }}
                >
                  Confirm
                </button>
              </div>
            </div>
          </div>
        )}
      </>
    );
  }

  return (
    <>
      {result !== undefined && <div style={{ ...s.msg, ...(result.ok ? s.ok : s.err) }}>{result.text}</div>}
      {segments.length === 0 && (
        <div style={{ ...s.msg, ...s.warn }}>
          No business segments are configured, so there is no structure to edit. Segments exist as{" "}
          <strong>mode</strong> rows in the configuration list.
        </div>
      )}
      {segments.map((row) => (
        <div style={s.card} key={row.key}>
          <div style={s.segRow}>
            <div style={{ flex: "1 1 300px" }}>
              <div style={s.segName}>{row.label}</div>
              <div style={s.path}>{pathPreview(row, row.chain)}</div>
            </div>
            <div style={{ flex: "0 0 auto", display: "flex", alignItems: "center", gap: 10 }}>
              <span
                style={{
                  ...s.badge,
                  ...(row.hasDocuments === undefined ? {} : row.hasDocuments ? s.badgeUsed : s.badgeFree),
                }}
              >
                {row.hasDocuments === undefined ? "checking…" : row.hasDocuments ? "in use" : "empty"}
              </span>
              <button
                style={row.chainError === undefined ? s.ghost : s.off}
                disabled={row.chainError !== undefined}
                onClick={() => startEdit(row)}
              >
                Edit structure
              </button>
            </div>
          </div>
          {row.chainError !== undefined && (
            <div style={{ ...s.msg, ...s.err, marginTop: 10, marginBottom: 0 }}>
              This segment&rsquo;s structure has a problem, so it cannot be edited here:{" "}
              {row.chainError}
            </div>
          )}
        </div>
      ))}
      <p style={{ fontSize: 12, color: "#8a8886", marginTop: 16, lineHeight: 1.6 }}>
        <strong>In use</strong> means documents are already filed under that segment, so adding a
        level leaves two folder shapes side by side. <strong>Empty</strong> means the structure can
        be changed freely.
      </p>
    </>
  );
}

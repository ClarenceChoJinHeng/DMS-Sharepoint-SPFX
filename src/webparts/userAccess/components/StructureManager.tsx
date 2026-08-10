import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, parseLevels, sanitizeFolderSegment } from "../../../shared/formModel";
import { effectiveOnDemandTiers, splitChain, validateChain } from "../../../shared/folderChain";
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

/**
 * What we know about the term set ID the admin typed.
 *
 * A wrong ID here is the worst failure this screen can produce, because nothing reports
 * it: the tier saves, reconciliation ignores below-Unit tiers entirely, and the only
 * symptom is an upload form with one permanently empty dropdown that blocks the upload.
 * The admin who typed it is not the person who hits it. So resolve the ID against the
 * term store while they are still looking at the field.
 *
 * `unknown` is deliberately NOT treated as bad. A 500 or a dropped connection says
 * nothing about the ID, and refusing to save on it would block a correct one.
 */
type SetCheck =
  | { state: "blank" }
  | { state: "malformed" }
  | { state: "checking" }
  | { state: "found"; name: string; count: number }
  | { state: "notfound" }
  | { state: "unknown"; status: number };

/** Accepts the braced/whitespaced forms people paste out of SharePoint UI and URLs. */
function normalizeGuid(raw: string): string {
  return raw.trim().replace(/^[{(]|[)}]$/g, "").trim().toLowerCase();
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

/**
 * Derive the column internal name from the level name — never typed by the admin.
 *
 * A SharePoint column's internal name is fixed at creation, permanently, from the title
 * it is created with. Letting an untrained admin type it offers a decision that cannot
 * be undone and whose consequence is invisible until a metadata write silently fails.
 * So strip everything SharePoint would encode, create under that, then set the display
 * title to what they typed.
 */
function columnNameFor(label: string): string {
  return sanitizeFolderSegment(label).replace(/[^A-Za-z0-9]/g, "");
}

/** The verdict, in the admin's language. Never shows the GUID back — they can see it. */
function setCheckMessage(c: SetCheck): string {
  switch (c.state) {
    case "checking":
      return "Checking…";
    case "malformed":
      return "That is not a term set ID. Copy the ID from the term store — it looks like " +
        "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf.";
    case "notfound":
      return "No term set with that ID exists on this site. If you copied it from another " +
        "site, or copied a term instead of the term set, it will not work here.";
    case "unknown":
      return `Could not check that ID right now${c.status ? ` (HTTP ${c.status})` : ""}. ` +
        "You can still add the level, but confirm the ID is right — a wrong one leaves the " +
        "dropdown empty and blocks uploads.";
    case "found":
      return c.count === 0
        ? `Found "${c.name}", but it has no terms yet. Add the terms before anyone uploads, ` +
          "or the dropdown will be empty and block them."
        : `Found "${c.name}" — ${c.count} ${c.count === 1 ? "value" : "values"}.`;
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

/**
 * Block Add only where the ID is KNOWN to be wrong.
 *
 * A `found` set with no terms is allowed through: authoring the tier before the terms is a
 * legitimate order of work, and the warning says what is still owed. `unknown` is allowed
 * for the reason on `SetCheck`. `checking` blocks for the few hundred ms it lasts, so a
 * fast click cannot outrun the verdict.
 */
function canAddTier(adding: DraftTier, check: SetCheck): boolean {
  const label = adding.label.trim();
  if (!label || !columnNameFor(label)) return false;
  return check.state !== "malformed" && check.state !== "notfound" && check.state !== "checking";
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
  // The Year / Document Type term sets, so a segment still running on the built-in pair can
  // be shown the levels it is EFFECTIVELY using rather than an empty list.
  const [legacySets, setLegacySets] = useState<{ year: string; docType: string }>({ year: "", docType: "" });
  const [setCheck, setSetCheck] = useState<SetCheck>({ state: "blank" });

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
   * The Year / Document Type term-set GUIDs from the `setting` rows the upload form reads.
   * Needed only to seed the built-in pair when a segment has no below-Unit levels yet.
   */
  const loadLegacyTermSets = async (): Promise<{ year: string; docType: string }> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
        `/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'&$top=200`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return { year: "", docType: "" };
    const data = await res.json();
    const rows = (data.value ?? []) as Array<{ Title?: string; SettingValue?: string }>;
    const get = (key: string): string =>
      (rows.filter((r) => (r.Title ?? "").trim() === key)[0]?.SettingValue ?? "").trim();
    return { year: get("termSet_yearPeriod"), docType: get("termSet_documentType") };
  };

  /**
   * Resolve a typed term set ID against the term store: does it exist, what is it called,
   * and how many top-level terms does it hold.
   *
   * The name is the part that actually catches mistakes. A GUID either resolving or not
   * only rules out typos; showing "Year / Period" back when they meant SubUnit catches the
   * commonest error by far — pasting the ID of the wrong set, or of a TERM instead of a
   * set. A term's GUID is well-formed and 404s here, which is exactly the outcome we want.
   *
   * The count is read separately because a set can exist and be empty, and an empty set
   * produces the same permanently-blank dropdown as a wrong ID.
   */
  const checkTermSet = async (guid: string): Promise<SetCheck> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${guid}?$select=id,localizedNames`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    // 404 is the answer, not an error: no set with that ID is readable from this site.
    // Term sets are per-site here, so one copied from another site's config lands here too.
    if (res.status === 404) return { state: "notfound" };
    if (!res.ok) return { state: "unknown", status: res.status };

    const set = (await res.json()) as { localizedNames?: Array<{ name?: string }> };
    const name = (set.localizedNames ?? [])[0]?.name ?? "";

    let count = 0;
    try {
      const kids: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${guid}/children?$select=id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (kids.ok) count = (((await kids.json()).value ?? []) as unknown[]).length;
    } catch {
      count = 0; // reported as "0 terms" below, which is a warning and not a block
    }
    return { state: "found", name, count };
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
      .then(() => loadLegacyTermSets())
      .then((sets) => {
        if (!cancelled) setLegacySets(sets);
      })
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

  // Check the term set ID as it is typed. Debounced, and `stale` discards a slow reply
  // that lands after the field has moved on — otherwise a verdict about an earlier,
  // half-typed GUID would sit under the field looking like a verdict about this one.
  const typedSet = adding === undefined ? "" : adding.termSetGuid;
  useEffect(() => {
    const guid = normalizeGuid(typedSet);
    if (!guid) {
      setSetCheck({ state: "blank" });
      return undefined;
    }
    if (!GUID_RE.test(guid)) {
      setSetCheck({ state: "malformed" });
      return undefined;
    }
    setSetCheck({ state: "checking" });
    let stale = false;
    const timer = setTimeout(() => {
      checkTermSet(guid)
        .then((r) => {
          if (!stale) setSetCheck(r);
        })
        .catch(() => {
          if (!stale) setSetCheck({ state: "unknown", status: 0 });
        });
    }, 400);
    return () => {
      stale = true;
      clearTimeout(timer);
    };
  }, [typedSet]);

  /**
   * The folder path a chain produces, with level names in brackets where a real value
   * goes. The only representation of the change an admin can reason about — which is why
   * it sits under every editor and the JSON sits nowhere.
   */
  const pathPreview = (seg: SegmentRow, chain: Level[]): string =>
    [seg.stagingFolder || seg.label, ...chain.map((l) => `[${l.label}]`)].join(" / ");

  /* ---------- Editing --------------------------------------------------- */

  /**
   * Open the editor on the levels the segment is EFFECTIVELY using, not just the ones
   * written down.
   *
   * A segment with nothing below Unit runs on the built-in Year → Document Type pair, and
   * showing an empty list would be a lie with teeth: adding one level would silently
   * replace that pair, dropping Year and Document Type from every future path. Seeding
   * the draft with them makes adding a level an insertion into a list the admin can see,
   * and saving writes all of it explicitly.
   *
   * The seeded pair comes from `effectiveOnDemandTiers`, so it carries the same column
   * names the form has always written and — importantly — NO `tidCol`, because Year and
   * Document Type are managed-metadata columns that neither have nor want a companion.
   */
  const startEdit = (seg: SegmentRow): void => {
    const { permissioned } = splitChain(seg.chain);
    const below = effectiveOnDemandTiers(seg.chain, legacySets.year, legacySets.docType);
    setEditing(seg.key);
    setDraft([...permissioned, ...below].map((l) => ({ ...l })));
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
    // Store the normalized GUID, never the raw paste: braces and stray whitespace survive
    // into the URL and 404 the term-store call at upload time, long after this screen.
    const set = normalizeGuid(adding.termSetGuid);
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
        // ensureTextColumn skips a column that already exists, whatever its type — so the
        // seeded Year / Document Type pair touches nothing, and only genuinely new levels
        // get columns created. A tier with no tidCol (managed metadata) never gets one
        // invented for it here.
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
          ` permissions.` +
          // New columns are created on the items but NOT added to any view, because adding them
          // there would reshape every view the client arranged — for every level anyone ever adds.
          // Invisible columns look exactly like the save having failed, so say it here instead.
          (created.length > 0
            ? ` The new columns are not shown in any library view yet — that is deliberate, so` +
              ` existing views keep the layout you set. Add them where you want them using the` +
              ` view's "Show or hide columns".`
            : ""),
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
            None — uploads would stop at the Unit folder. Add at least one level.
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
            {setCheck.state !== "blank" && (
              <p style={{ ...s.hint, marginTop: 6, fontSize: 12, ...setCheckStyle(setCheck) }}>
                {setCheckMessage(setCheck)}
              </p>
            )}

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
                style={canAddTier(adding, setCheck) ? s.btn : s.off}
                disabled={!canAddTier(adding, setCheck)}
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

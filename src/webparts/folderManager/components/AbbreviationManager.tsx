import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, parseLevels } from "../../../shared/formModel";
import { splitChain } from "../../../shared/folderChain";
import { abbrevListTitle } from "../../../shared/folderAbbreviation";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import {
  AbbrevRowDraft,
  changedRows,
  folderNameFor,
  hasBlockingProblem,
  renamingRows,
  RowProblem,
  validateRows,
} from "../../../shared/abbreviationDraft";

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
  warn: { background: "#fff4e5", border: "1px solid #f0d9b5", color: "#7a4f00" },
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
    gridTemplateColumns: "minmax(0, 1fr) 150px max-content minmax(0, 170px)",
    alignItems: "center",
    columnGap: 12,
    rowGap: 4,
    padding: "8px 10px",
    borderBottom: "1px solid #f3f2f1",
  },
  tierHead: { display: "flex", alignItems: "center", gap: 12, padding: "10px 10px 6px", fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: "#605e5c", flexWrap: "wrap" },
  // Spans every track: a collision message names another term and its parent, and truncating it into
  // one column would hide the half that says which other row to look at.
  problem: { fontSize: 11, lineHeight: 1.5, marginTop: 3, gridColumn: "1 / -1" },
  hint: { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  pathHint: { fontSize: 11, color: "#8a8886", fontFamily: "Consolas, monospace" },
};

export interface AbbreviationManagerProps {
  context: WebPartContext;
  siteUrl: string;
  onDirtyChange?: (dirty: boolean) => void;
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
}

export default function AbbreviationManager({
  context,
  siteUrl,
  onDirtyChange,
  onMissingChange,
}: AbbreviationManagerProps): React.ReactElement {
  const [segments, setSegments] = useState<SegmentOption[]>([]);
  const [chosen, setChosen] = useState<string>("");
  const [rows, setRows] = useState<AbbrevRowDraft[]>([]);
  const [loading, setLoading] = useState(true);
  const [treeLoading, setTreeLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<{ ok: boolean; text: string } | undefined>(undefined);
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
      const opts = raw
        .filter((r) => (r.TermSetGuid ?? "").trim() !== "")
        .map((r) => {
          const chain: Level[] = parseLevels(r.Levels ?? "");
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
          };
        })
        .sort((a, b) => a.label.localeCompare(b.label));
      setSegments(opts);
      if (opts.length > 0) setChosen(opts[0].key);
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

      // Walk only as deep as there are permissioned levels. Anything deeper belongs to a below-Unit
      // tier, which names its folder from the term label and needs no code — walking it would list
      // terms the page must not offer to name.
      const maxDepth = seg.levelNames.length;
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
          level: seg.levelNames[n.depth - 1] ?? `Level ${n.depth}`,
          parentGuid: n.parentId,
          abbreviation: found ? found.abbreviation : "",
          original: found ? found.abbreviation : "",
        };
      });
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

  const sameAsName = (termGuid: string): void =>
    setRows(rows.map((r) => (r.termGuid === termGuid ? { ...r, abbreviation: r.label } : r)));

  /**
   * Fill every EMPTY row of a level from its term name.
   *
   * Empty only, deliberately: overwriting codes the client already chose would silently discard
   * authored data and, worse, rename live folders on the next reconciliation run.
   */
  const sameAsNameForLevel = (level: string): void =>
    setRows(
      rows.map((r) =>
        r.level === level && r.abbreviation.trim() === "" ? { ...r, abbreviation: r.label } : r,
      ),
    );

  /* ── Save ──────────────────────────────────────────────────────────────────── */

  const save = async (): Promise<void> => {
    const seg = segment();
    if (!seg || blocked) return;
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
    } catch (e) {
      setResult({ ok: false, text: (e as Error).message });
    } finally {
      setBusy(false);
    }
  };

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

  /* ── Render ────────────────────────────────────────────────────────────────── */

  if (loading) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading segments&hellip;</p>;

  return (
    <div>
      {result && <div style={{ ...s.msg, ...(result.ok ? s.ok : s.err) }}>{result.text}</div>}

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
        {seg && (
          <div style={s.hint}>
            Top folder <strong>{seg.stagingFolder || "(not set)"}</strong> — that one is set on the
            segment itself, not here.
          </div>
        )}
      </div>

      {seg && seg.belowNames.length > 0 && (
        <div style={{ ...s.msg, ...s.info }}>
          <strong>{seg.belowNames.join(", ")}</strong> need no abbreviation — those folders are named
          from the term itself (<span style={s.pathHint}>2026</span>,{" "}
          <span style={s.pathHint}>Tax Return</span>) and are created when someone uploads.
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
          {missing > 0 && (
            <div style={{ ...s.msg, ...s.warn }}>
              {missing} term{missing === 1 ? "" : "s"} still {missing === 1 ? "has" : "have"} no
              abbreviation. Reconciliation skips those — no folder is created, and nobody in that unit
              can upload.
            </div>
          )}

          {seg.levelNames.map((levelName) => {
            const levelRows = rows.filter((r) => r.level === levelName);
            if (levelRows.length === 0) return null;
            const emptyHere = levelRows.filter((r) => r.abbreviation.trim() === "").length;
            return (
              <div key={levelName} style={{ marginBottom: 18 }}>
                <div style={s.tierHead}>
                  <span>
                    {levelName} ({levelRows.length})
                  </span>
                  {emptyHere > 0 && (
                    <button style={s.ghost} onClick={() => sameAsNameForLevel(levelName)}>
                      Use term names for the {emptyHere} empty one{emptyHere === 1 ? "" : "s"}
                    </button>
                  )}
                </div>
                {levelRows.map((r) => {
                  const p = problems[r.termGuid];
                  const name = folderNameFor(r.abbreviation);
                  return (
                    <div key={r.termGuid} style={s.row}>
                      <span style={{ fontSize: 13 }}>{r.label}</span>
                      <input
                        style={{ ...s.input, ...(p && p.error ? s.inputBad : {}) }}
                        value={r.abbreviation}
                        placeholder="code"
                        onChange={(e) => setCode(r.termGuid, e.target.value)}
                      />
                      <button style={s.ghost} onClick={() => sameAsName(r.termGuid)}>
                        Same as term name
                      </button>
                      <span style={s.pathHint}>{name ? `/${name}` : ""}</span>
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

          <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 8 }}>
            <button
              style={!dirty || blocked || busy ? s.off : s.btn}
              disabled={!dirty || blocked || busy}
              onClick={() => {
                save().catch(() => undefined); // save() reports its own failures into `result`
              }}
            >
              {busy
                ? "Saving…"
                : dirty
                  ? `Save ${pending.length} change${pending.length === 1 ? "" : "s"}`
                  : "Save"}
            </button>
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

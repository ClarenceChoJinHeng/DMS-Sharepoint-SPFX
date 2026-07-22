import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from "@microsoft/sp-http";
import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  roleFromGroupName,
  GroupMapRole,
  GroupMapDraft,
  GroupMapWriteRow,
} from "../../../shared/groupMapModel";

type Props = { context: WebPartContext; siteUrl: string };
type GroupPick = { id: string; displayName: string };
type ModePick = { label: string; termSetGuid: string };
type TermLite = { id: string; label: string };
type ExistingRow = GroupMapWriteRow & { itemId: number };

const ROLES: GroupMapRole[] = ["MEMBER", "UPL", "APR", "GLOBAL"];
const GROUP_MAP_LIST = "DMS Group Map";
const CONFIG_LIST = "DMS Config";

const s: Record<string, React.CSSProperties> = {
  wrap:       { fontSize: 13, color: "#242424", lineHeight: 1.5 },
  intro:      { fontSize: 13, color: "#444", margin: "0 0 16px" },
  card:       { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  label:      { display: "block", fontWeight: 600, fontSize: 12, margin: "12px 0 4px" },
  input:      { width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  select:     { width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff" },
  ddwrap:     { position: "relative" },
  dd:         { position: "absolute", top: "100%", left: 0, right: 0, zIndex: 20, background: "#fff", border: "1px solid #c7c7c7", borderRadius: 4, maxHeight: 220, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,.12)" },
  ddItem:     { padding: "7px 10px", cursor: "pointer", borderBottom: "1px solid #f0f0f0" },
  roleRow:    { display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 },
  roleBtn:    { padding: "5px 12px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  roleActive: { background: "#0f6c3f", color: "#fff", borderColor: "#0f6c3f" },
  pickedChip: { display: "inline-flex", alignItems: "center", gap: 8, padding: "5px 10px", background: "#eef6f0", border: "1px solid #b7dcc4", borderRadius: 4, fontSize: 12 },
  chipX:      { border: "none", background: "transparent", cursor: "pointer", color: "#0f6c3f", fontWeight: 700 },
  seglvl:     { marginTop: 6, fontSize: 12, color: "#0f6c3f", background: "transparent", border: "none", cursor: "pointer", textDecoration: "underline", padding: 0 },
  preview:    { marginTop: 14, padding: "10px 12px", background: "#fff", border: "1px dashed #b7dcc4", borderRadius: 4, fontSize: 12 },
  addBtn:     { marginTop: 12, padding: "7px 18px", fontSize: 13, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  addBtnOff:  { marginTop: 12, padding: "7px 18px", fontSize: 13, background: "#c7c7c7", color: "#fff", border: "none", borderRadius: 4, cursor: "not-allowed" },
  table:      { width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 8 },
  th:         { textAlign: "left", padding: "6px 8px", borderBottom: "2px solid #e1e1e1", fontWeight: 600, color: "#555" },
  td:         { padding: "6px 8px", borderBottom: "1px solid #f0f0f0", verticalAlign: "top" },
  delBtn:     { padding: "3px 10px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  ghost:      { padding: "3px 10px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  toast:      { position: "fixed", bottom: 20, right: 20, padding: "10px 16px", borderRadius: 6, color: "#fff", fontSize: 13, zIndex: 50, boxShadow: "0 4px 14px rgba(0,0,0,.18)" },
  mono:       { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
};

export default function GroupMapBuilder({ context, siteUrl }: Props): React.ReactElement {
  const [modes, setModes]       = useState<ModePick[]>([]);
  const [existing, setExisting] = useState<ExistingRow[]>([]);
  const [busy, setBusy]         = useState(false);
  const [toast, setToast]       = useState<{ message: string; error: boolean } | undefined>(undefined);

  // Draft selections
  const [group, setGroup]         = useState<GroupPick | undefined>(undefined);
  const [role, setRole]           = useState<GroupMapRole | "">("");
  const [mode, setMode]           = useState<ModePick | undefined>(undefined);
  const [cascade, setCascade]     = useState<TermLite[][]>([]); // options per level
  const [chosen, setChosen]       = useState<TermLite[]>([]);   // picked term per level
  const [tierGuid, setTierGuid]   = useState<string>("");

  // Group search box
  const [query, setQuery]         = useState("");
  const [results, setResults]     = useState<GroupPick[]>([]);
  const [searching, setSearching] = useState(false);

  const [confirmDel, setConfirmDel] = useState<number | undefined>(undefined);

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(undefined), 5000);
  };

  /* ── Data access ───────────────────────────────────────────────────────── */

  const searchGroups = async (q: string): Promise<GroupPick[]> => {
    const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");
    let req = client
      .api("/groups")
      .header("ConsistencyLevel", "eventual")
      .count(true)
      .select("id,displayName")
      .top(25);
    if (q.trim().length >= 1) req = req.search(`"displayName:${q.trim()}"`);
    const res = await req.get();
    const groups = (res as { value?: Array<{ id: string; displayName: string }> }).value ?? [];
    return groups.map((g) => ({ id: g.id, displayName: g.displayName }));
  };

  const loadModes = async (): Promise<ModePick[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(CONFIG_LIST)}')/items?$select=ModeLabel,TermSetGuid,Levels&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{ ModeLabel?: string; TermSetGuid?: string; Levels?: string }>)
      .filter((r) => (r.TermSetGuid ?? "").trim() && (r.Levels ?? "").trim())
      .map((r) => ({ label: (r.ModeLabel ?? "").trim() || (r.TermSetGuid ?? "").trim(), termSetGuid: (r.TermSetGuid ?? "").trim() }));
  };

  const loadTerms = async (url: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      url, SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{ id: string; labels: Array<{ name: string }> }>)
      .map((t) => ({ id: t.id, label: t.labels[0].name }));
  };

  const loadTops = (termSetGuid: string): Promise<TermLite[]> =>
    loadTerms(`${siteUrl}/_api/v2.1/termStore/sets/${termSetGuid}/children`);

  const loadChildren = (termSetGuid: string, parentId: string): Promise<TermLite[]> =>
    loadTerms(`${siteUrl}/_api/v2.1/termStore/sets/${termSetGuid}/terms/${parentId}/children`);

  const loadExisting = async (): Promise<ExistingRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST)}')/items?$select=Id,GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{ Id: number; GroupId?: string; GroupName?: string; Segment?: string; UnitTermGuid?: string; Role?: string }>)
      .map((r) => ({
        itemId: r.Id,
        GroupId: r.GroupId ?? "",
        GroupName: r.GroupName ?? "",
        Segment: r.Segment ?? "",
        UnitTermGuid: r.UnitTermGuid ?? "",
        Role: (r.Role ?? "").toUpperCase() as GroupMapRole,
      }));
  };

  const postRow = async (row: GroupMapWriteRow): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST)}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata" },
        // Title is mandatory on a default SP list; set it so the create never 400s.
        body: JSON.stringify({ Title: row.GroupName || row.GroupId, ...row }),
      },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  };

  const deleteRow = async (itemId: number): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST)}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "IF-MATCH": "*", "X-HTTP-Method": "DELETE" } },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  };

  /* ── Mount ─────────────────────────────────────────────────────────────── */

  useEffect(() => {
    loadModes().then(setModes).catch(() => setModes([]));
    loadExisting().then(setExisting).catch(() => setExisting([]));
  }, []);

  /* ── Group search (debounced) ──────────────────────────────────────────── */

  useEffect(() => {
    if (group) return; // already picked
    const q = query.trim();
    if (q.length < 1) { setResults([]); return; }
    let cancelled = false;
    setSearching(true);
    const t = setTimeout(() => {
      searchGroups(q)
        .then((r) => { if (!cancelled) setResults(r); })
        .then(() => { if (!cancelled) setSearching(false); })
        .catch(() => { if (!cancelled) { setResults([]); setSearching(false); } });
    }, 350);
    return () => { cancelled = true; clearTimeout(t); };
  }, [query, group]);

  /* ── Selection handlers ────────────────────────────────────────────────── */

  const pickGroup = (g: GroupPick): void => {
    setGroup(g);
    setQuery(g.displayName);
    setResults([]);
    // Pre-select the role implied by the name suffix (_UPL/_APR); admin can override.
    setRole(roleFromGroupName(g.displayName));
  };

  const clearGroup = (): void => {
    setGroup(undefined);
    setQuery("");
    setResults([]);
  };

  const pickRole = (r: GroupMapRole): void => {
    setRole(r);
    if (r === "GLOBAL") { setMode(undefined); setCascade([]); setChosen([]); setTierGuid(""); }
  };

  const pickMode = async (termSetGuid: string): Promise<void> => {
    const m = modes.find((x) => x.termSetGuid === termSetGuid);
    setMode(m);
    setChosen([]);
    setTierGuid("");
    setCascade([]);
    if (m) {
      const tops = await loadTops(m.termSetGuid).catch(() => [] as TermLite[]);
      setCascade([tops]);
    }
  };

  const pickTerm = async (levelIdx: number, termId: string): Promise<void> => {
    if (!mode) return;
    const term = (cascade[levelIdx] ?? []).find((t) => t.id === termId);
    if (!term) return;
    const newChosen = chosen.slice(0, levelIdx);
    newChosen[levelIdx] = term;
    setChosen(newChosen);
    setTierGuid(term.id); // assign at this tier by default
    const kids = await loadChildren(mode.termSetGuid, term.id).catch(() => [] as TermLite[]);
    const newCascade = cascade.slice(0, levelIdx + 1);
    if (kids.length) newCascade[levelIdx + 1] = kids;
    setCascade(newCascade);
  };

  const assignAtSegment = (): void => {
    if (!mode) return;
    setChosen([]);
    setTierGuid(mode.termSetGuid);
    setCascade((c) => c.slice(0, 1)); // keep top-level options for re-pick
  };

  /* ── Derived ───────────────────────────────────────────────────────────── */

  const draft: GroupMapDraft = {
    groupId: group?.id ?? "",
    groupName: group?.displayName ?? "",
    role: (role || undefined) as GroupMapRole,
    segmentGuid: mode?.termSetGuid,
    tierGuid,
  };
  const draftErrors = validateDraft(draft);
  const canAdd = draftErrors.length === 0 && !busy;

  const tierLabel = (): string => {
    if (role === "GLOBAL") return "—";
    if (!tierGuid) return "(not chosen)";
    if (mode && tierGuid === mode.termSetGuid) return "(segment level)";
    const hit = chosen.find((t) => t.id === tierGuid);
    return hit ? hit.label : tierGuid;
  };

  const segmentLabelFor = (guid: string): string => {
    const m = modes.find((x) => x.termSetGuid === guid);
    return m ? m.label : guid;
  };

  const onAdd = async (): Promise<void> => {
    const row = buildGroupMapRow(draft);
    if (isDuplicateRow(existing, row)) { showToast("This exact mapping already exists.", true); return; }
    setBusy(true);
    try {
      await postRow(row);
      setExisting(await loadExisting());
      // reset the draft (keep the picked group for quick multi-role mapping)
      setRole(""); setMode(undefined); setCascade([]); setChosen([]); setTierGuid("");
      showToast("Row added — re-run Folder Reconciliation to apply permissions.", false);
    } catch (e) {
      showToast(`Add failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  const onDelete = async (itemId: number): Promise<void> => {
    setBusy(true);
    try {
      await deleteRow(itemId);
      setExisting(await loadExisting());
      setConfirmDel(undefined);
      showToast("Row deleted — re-run Folder Reconciliation to apply the change.", false);
    } catch (e) {
      showToast(`Delete failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  /* ── Render ────────────────────────────────────────────────────────────── */

  return (
    <div style={s.wrap}>
      <p style={s.intro}>
        Map an <strong>existing</strong> security group to a segment, tier, and role. This writes a clean
        row into the <strong>DMS Group Map</strong> list — no hand-typed GUIDs. It does <strong>not</strong>{" "}
        create the group or apply permissions: after adding rows, <strong>re-run Folder Reconciliation</strong>{" "}
        to grant access (MEMBER → Read, UPL → Contribute, APR → Design).
      </p>

      <div style={s.card}>
        {/* Group */}
        <label style={s.label}>Group</label>
        {group ? (
          <span style={s.pickedChip}>
            {group.displayName}
            <button style={s.chipX} disabled={busy} title="Change" onClick={clearGroup}>✕</button>
          </span>
        ) : (
          <div style={s.ddwrap}>
            <input
              style={s.input}
              placeholder="Type to search Entra groups…"
              value={query}
              disabled={busy}
              onChange={(e) => setQuery(e.target.value)}
            />
            {(searching || results.length > 0) && (
              <div style={s.dd}>
                {searching && <div style={s.ddItem}>Searching…</div>}
                {!searching && results.map((g) => (
                  <div key={g.id} style={s.ddItem} onClick={() => pickGroup(g)}>{g.displayName}</div>
                ))}
                {!searching && results.length === 0 && query.trim() && <div style={s.ddItem}>No groups found.</div>}
              </div>
            )}
          </div>
        )}

        {/* Role */}
        <label style={s.label}>Role</label>
        <div style={s.roleRow}>
          {ROLES.map((r) => (
            <button
              key={r}
              disabled={busy}
              style={{ ...s.roleBtn, ...(role === r ? s.roleActive : {}) }}
              onClick={() => pickRole(r)}
            >
              {r}
            </button>
          ))}
        </div>

        {/* Segment + cascade (hidden for GLOBAL) */}
        {role && role !== "GLOBAL" && (
          <>
            <label style={s.label}>Segment</label>
            <select
              style={s.select}
              disabled={busy}
              value={mode?.termSetGuid ?? ""}
              onChange={(e) => { pickMode(e.target.value).catch(() => undefined); }}
            >
              <option value="">— select a segment —</option>
              {modes.map((m) => <option key={m.termSetGuid} value={m.termSetGuid}>{m.label}</option>)}
            </select>

            {mode && (
              <>
                <label style={s.label}>Tier (where this group applies)</label>
                {cascade.map((opts, i) => (
                  <select
                    key={i}
                    style={{ ...s.select, marginBottom: 6 }}
                    disabled={busy}
                    value={chosen[i]?.id ?? ""}
                    onChange={(e) => { pickTerm(i, e.target.value).catch(() => undefined); }}
                  >
                    <option value="">— select —</option>
                    {opts.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
                  </select>
                ))}
                <button style={s.seglvl} disabled={busy} onClick={assignAtSegment}>
                  Assign at segment level (whole {mode.label})
                </button>
              </>
            )}
          </>
        )}

        {/* Preview */}
        {group && role && (
          <div style={s.preview}>
            <strong>Will write:</strong> {group.displayName} · <em>{role}</em>
            {role !== "GLOBAL" && <> · segment <strong>{mode?.label ?? "(none)"}</strong> · tier <strong>{tierLabel()}</strong></>}
          </div>
        )}

        <button style={canAdd ? s.addBtn : s.addBtnOff} disabled={!canAdd} onClick={() => { onAdd().catch(() => undefined); }}>
          Add mapping
        </button>
      </div>

      {/* Existing rows */}
      <h3 style={{ fontSize: 14, margin: "0 0 6px" }}>Existing mappings ({existing.length})</h3>
      <table style={s.table}>
        <thead>
          <tr>
            <th style={s.th}>Group</th>
            <th style={s.th}>Segment</th>
            <th style={s.th}>Tier</th>
            <th style={s.th}>Role</th>
            <th style={s.th}></th>
          </tr>
        </thead>
        <tbody>
          {existing.length === 0 && (
            <tr><td style={s.td} colSpan={5}>No mappings yet.</td></tr>
          )}
          {existing.map((r) => (
            <tr key={r.itemId}>
              <td style={s.td}>{r.GroupName || <span style={s.mono}>{r.GroupId}</span>}</td>
              <td style={s.td}>{r.Segment ? segmentLabelFor(r.Segment) : "—"}</td>
              <td style={s.td}>
                {!r.UnitTermGuid ? "—" : r.UnitTermGuid === r.Segment ? "(segment level)" : <span style={s.mono}>{r.UnitTermGuid}</span>}
              </td>
              <td style={s.td}>{r.Role}</td>
              <td style={s.td}>
                {confirmDel === r.itemId ? (
                  <span style={{ display: "inline-flex", gap: 6 }}>
                    <button style={s.delBtn} disabled={busy} onClick={() => { onDelete(r.itemId).catch(() => undefined); }}>Yes, delete</button>
                    <button style={s.ghost} disabled={busy} onClick={() => setConfirmDel(undefined)}>Cancel</button>
                  </span>
                ) : (
                  <button style={s.delBtn} disabled={busy} onClick={() => setConfirmDel(r.itemId)}>Delete</button>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#a4262c" : "#0f6c3f" }}>{toast.message}</div>
      )}
    </div>
  );
}

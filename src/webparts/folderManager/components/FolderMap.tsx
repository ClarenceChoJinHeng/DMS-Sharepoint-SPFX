import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IFolderManagerProps } from "./IFolderManagerProps";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";

// Title is RESOLVED — the client renames the list to "CRS Folder Map" at import.
const LIST_NAME = (): string => cachedListTitle(LIST_SUFFIX.folderMap);
const LIST_ENC  = (): string => encodeURIComponent(LIST_NAME());

// LIST_TYPE is deliberately NOT resolved and must stay "DMS".
//
// SharePoint derives a list's item entity type from the name it was CREATED with; renaming the
// title never moves it. Verified live 2026-08-05 against the client's fully CRS-renamed site,
// which still reports SP.Data.DMS_x0020_Folder_x0020_MapListItem. Deriving this from the current
// title would break every write on exactly the sites the rename was meant to support.
const LIST_TYPE = "SP.Data.DMS_x0020_Folder_x0020_MapListItem";

type TermOption = { id: string; label: string };
type ModeInfo  = { key: string; label: string; termSetGuid: string };
type MapEntry  = { id: number; mode: string; termLabel: string; folderName: string };

const s: Record<string, React.CSSProperties> = {
  wrap:       { maxWidth: 720, margin: "32px auto", padding: "0 24px 48px", fontFamily: "'Segoe UI', sans-serif" },
  h2:         { fontSize: 22, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle:   { fontSize: 13, color: "#666", margin: "0 0 24px" },
  segWrap:    { display: "flex", marginBottom: 24 },
  seg:        { display: "flex", border: "1px solid #0f6c3f", borderRadius: 8, overflow: "hidden" },
  segBtn:     { padding: "8px 22px", fontSize: 13, fontFamily: "'Segoe UI', sans-serif", fontWeight: 600, cursor: "pointer", background: "#fff", color: "#0f6c3f", border: "none", borderRight: "1px solid #0f6c3f" },
  segActive:  { background: "#0f6c3f", color: "#fff" },
  note:       { fontSize: 12, color: "#666", background: "#f7f7f7", border: "1px solid #e0e0e0", borderRadius: 4, padding: "8px 12px", marginBottom: 20 },
  errNote:    { fontSize: 12, color: "#a4262c", background: "#fdf3f3", border: "1px solid #f4c2c2", borderRadius: 4, padding: "10px 14px", marginBottom: 20 },
  table:      { width: "100%", borderCollapse: "collapse" as const, marginBottom: 24 },
  th:         { fontSize: 11, fontWeight: 700, textTransform: "uppercase" as const, letterSpacing: ".06em", color: "#555", borderBottom: "2px solid #e0e0e0", padding: "6px 10px", textAlign: "left" as const },
  td:         { padding: "9px 10px", fontSize: 13, borderBottom: "1px solid #f0f0f0", verticalAlign: "middle" as const },
  arrow:      { color: "#bbb", fontSize: 16, textAlign: "center" as const, width: 32, padding: "9px 4px" },
  termBadge:  { display: "inline-block", background: "#e8f5ee", color: "#0f6c3f", borderRadius: 4, padding: "2px 8px", fontSize: 12, fontWeight: 600 },
  deleteBtn:  { background: "none", border: "none", color: "#a4262c", cursor: "pointer", fontSize: 12, fontFamily: "'Segoe UI', sans-serif", padding: "4px 8px" },
  addSection: { marginTop: 8 },
  addLabel:   { fontSize: 11, fontWeight: 700, textTransform: "uppercase" as const, letterSpacing: ".06em", color: "#0f6c3f", margin: "0 0 10px" },
  addRow:     { display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" as const },
  select:     { padding: "7px 10px", border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 13, background: "#fff", minWidth: 190 },
  input:      { padding: "7px 10px", border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 13, width: 200 },
  addBtn:     { padding: "7px 20px", background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 13, fontWeight: 600, cursor: "pointer" },
  addBtnDis:  { opacity: .45, cursor: "default" },
  emptyMsg:   { fontSize: 13, color: "#999", padding: "12px 0" },
  allMapped:  { fontSize: 13, color: "#0f6c3f", padding: "4px 0 12px" },
  toast:      { position: "fixed" as const, top: 24, right: 24, color: "#fff", padding: "14px 44px 14px 16px", borderRadius: 6, fontSize: 13, zIndex: 9999, minWidth: 280, maxWidth: 420, boxShadow: "0 4px 16px rgba(0,0,0,.18)" },
  toastClose: { position: "absolute" as const, top: 10, right: 12, background: "none", border: "none", cursor: "pointer", color: "#fff", fontSize: 16, opacity: .7, lineHeight: "1" },
};

export default function FolderMap({ context }: IFolderManagerProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;

  const [modes,           setModes]           = useState<ModeInfo[]>([]);
  const [selectedModeKey, setSelectedModeKey] = useState<string>("");
  const [termOptions,     setTermOptions]     = useState<TermOption[]>([]);
  const [entries,         setEntries]         = useState<MapEntry[]>([]);
  const [newTermId,       setNewTermId]       = useState<string>("");
  const [newFolderName,   setNewFolderName]   = useState<string>("");
  const [loading,         setLoading]         = useState(true);
  const [saving,          setSaving]          = useState(false);
  const [listMissing,     setListMissing]     = useState(false);
  const [toast,           setToast]           = useState<{ message: string; error: boolean } | null>(null);

  const showToast = (message: string, error = false): void => {
    setToast({ message, error });
    setTimeout(() => setToast(null), 5000);
  };

  /* -- Term Store ---------------------------------------------------------- */

  const loadTermSet = async (termSetGuid: string): Promise<TermOption[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetGuid}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value ?? []).map(
      (t: { id: string; labels: Array<{ name: string }> }) => ({
        id:    t.id,
        label: t.labels[0]?.name ?? t.id,
      }),
    );
  };

  /* -- DMS Config ---------------------------------------------------------- */

  const loadModes = async (): Promise<ModeInfo[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items?$select=Title,ModeLabel,TermSetGuid&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value ?? []).map(
      (item: { Title: string; ModeLabel: string; TermSetGuid: string }) => ({
        key:         item.Title,
        label:       item.ModeLabel,
        termSetGuid: item.TermSetGuid,
      }),
    );
  };

  /* -- Folder Map list ----------------------------------------------------- */

  const loadEntries = async (): Promise<MapEntry[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${LIST_ENC()}')/items?$select=Id,Title,TermLabel,FolderName&$orderby=Title,TermLabel`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) {
      if (res.status === 404) setListMissing(true);
      return [];
    }
    setListMissing(false);
    const data = await res.json();
    return (data.value ?? []).map(
      (item: { Id: number; Title: string; TermLabel: string; FolderName: string }) => ({
        id:         item.Id,
        mode:       item.Title,
        termLabel:  item.TermLabel,
        folderName: item.FolderName,
      }),
    );
  };

  const addEntry = async (): Promise<void> => {
    const term = termOptions.find((t) => t.id === newTermId);
    if (!selectedModeKey || !term || !newFolderName.trim()) {
      showToast("Select a term and enter a folder name.", true);
      return;
    }
    const duplicate = entries.some(
      (e) => e.mode === selectedModeKey && e.termLabel === term.label,
    );
    if (duplicate) {
      showToast(`A mapping for "${term.label}" already exists in this mode.`, true);
      return;
    }

    setSaving(true);
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${LIST_ENC()}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          __metadata: { type: LIST_TYPE },
          Title:      selectedModeKey,
          TermLabel:  term.label,
          FolderName: newFolderName.trim(),
        }),
      },
    );
    setSaving(false);

    if (!res.ok) {
      showToast("Failed to save. Verify the list exists and you have Contribute access.", true);
      return;
    }
    showToast("Mapping saved.");
    setNewTermId("");
    setNewFolderName("");
    const refreshed = await loadEntries();
    setEntries(refreshed);
  };

  const deleteEntry = async (id: number, label: string): Promise<void> => {
    setSaving(true);
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${LIST_ENC()}')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "X-HTTP-Method": "DELETE",
          "IF-MATCH":      "*",
        },
      },
    );
    setSaving(false);

    if (!res.ok) {
      showToast(`Failed to remove "${label}".`, true);
      return;
    }
    showToast(`Mapping for "${label}" removed.`);
    setEntries((prev) => prev.filter((e) => e.id !== id));
  };

  /* -- Init ---------------------------------------------------------------- */

  useEffect(() => {
    const init = async (): Promise<void> => {
      // Names FIRST: every read below resolves through the cache, and an unprimed cache falls
      // back to the legacy DMS titles, which 404 on a CRS-renamed site.
      await primeNames(context.spHttpClient, siteUrl);
      const loadedModes = await loadModes();
      setModes(loadedModes);

      const firstKey  = loadedModes[0]?.key ?? "";
      const firstMode = loadedModes[0];
      setSelectedModeKey(firstKey);

      const [loadedEntries, loadedTerms] = await Promise.all([
        loadEntries(),
        firstMode
          ? loadTermSet(firstMode.termSetGuid)
          : Promise.resolve([] as TermOption[]),
      ]);
      setEntries(loadedEntries);
      setTermOptions(loadedTerms);
      setLoading(false);
    };
    init().catch(() => {
      setLoading(false);
      showToast("Could not load data. Check your permissions.", true);
    });
  }, []);

  const handleModeSelect = async (key: string): Promise<void> => {
    setSelectedModeKey(key);
    setNewTermId("");
    const mode = modes.find((m) => m.key === key);
    if (!mode) { setTermOptions([]); return; }
    const terms = await loadTermSet(mode.termSetGuid).catch(() => [] as TermOption[]);
    setTermOptions(terms);
  };

  /* -- Render -------------------------------------------------------------- */

  const modeEntries = entries.filter((e) => e.mode === selectedModeKey);
  const unusedTerms = termOptions.filter(
    (t) => !entries.some((e) => e.mode === selectedModeKey && e.termLabel === t.label),
  );
  const canAdd = !!newTermId && !!newFolderName.trim() && !saving;

  return (
    <section style={s.wrap}>
      <h2 style={s.h2}>Folder Map</h2>
      <p style={s.subtitle}>
        Map each term label to its actual SharePoint folder name.
        If no mapping exists, the term label is used as the folder name (fallback).
      </p>

      {listMissing && (
        <div style={s.errNote}>
          <strong>List not found.</strong> Create a SharePoint list named <strong>&ldquo;{LIST_NAME()}&rdquo;</strong> with
          two extra columns: <code>TermLabel</code> (Single line of text) and{" "}
          <code>FolderName</code> (Single line of text). Then refresh this page.
        </div>
      )}

      {loading ? (
        <p style={{ fontSize: 13, color: "#666" }}>Loading…</p>
      ) : (
        <>
          {/* Mode selector */}
          <div style={s.segWrap}>
            <div style={s.seg}>
              {modes.map((m, i) => (
                <button
                  key={m.key}
                  onClick={() => handleModeSelect(m.key)}
                  style={{
                    ...s.segBtn,
                    ...(i === modes.length - 1 ? { borderRight: "none" } : {}),
                    ...(selectedModeKey === m.key ? s.segActive : {}),
                  }}
                >
                  {m.label}
                </button>
              ))}
            </div>
          </div>

          <div style={s.note}>
            Term labels are pulled directly from the Term Store. The folder name must match exactly
            what is in SharePoint (case-sensitive).
          </div>

          {/* Existing mappings */}
          {modeEntries.length === 0 ? (
            <p style={s.emptyMsg}>
              No mappings yet for this mode — the term label will be used as the folder name.
            </p>
          ) : (
            <table style={s.table}>
              <thead>
                <tr>
                  <th style={s.th}>Term Label (Term Store)</th>
                  <th style={{ ...s.th, ...s.arrow }} />
                  <th style={s.th}>Folder Name (SharePoint)</th>
                  <th style={{ ...s.th, width: 80 }} />
                </tr>
              </thead>
              <tbody>
                {modeEntries.map((entry) => (
                  <tr key={entry.id}>
                    <td style={s.td}>
                      <span style={s.termBadge}>{entry.termLabel}</span>
                    </td>
                    <td style={{ ...s.td, ...s.arrow }}>→</td>
                    <td style={s.td}>{entry.folderName}</td>
                    <td style={s.td}>
                      <button
                        onClick={() => deleteEntry(entry.id, entry.termLabel)}
                        disabled={saving}
                        style={{ ...s.deleteBtn, ...(saving ? { opacity: .5 } : {}) }}
                      >
                        Remove
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}

          {/* Add new mapping */}
          {unusedTerms.length > 0 ? (
            <div style={s.addSection}>
              <p style={s.addLabel}>Add Mapping</p>
              <div style={s.addRow}>
                <select
                  value={newTermId}
                  onChange={(e) => setNewTermId(e.target.value)}
                  style={s.select}
                >
                  <option value="">— Select term —</option>
                  {unusedTerms.map((t) => (
                    <option key={t.id} value={t.id}>{t.label}</option>
                  ))}
                </select>
                <span style={{ color: "#bbb", fontSize: 16 }}>→</span>
                <input
                  type="text"
                  placeholder="Exact folder name in SharePoint"
                  value={newFolderName}
                  onChange={(e) => setNewFolderName(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && canAdd) void addEntry(); }}
                  style={s.input}
                />
                <button
                  onClick={addEntry}
                  disabled={!canAdd}
                  style={{ ...s.addBtn, ...(!canAdd ? s.addBtnDis : {}) }}
                >
                  {saving ? "Saving…" : "Add"}
                </button>
              </div>
            </div>
          ) : termOptions.length > 0 ? (
            <p style={s.allMapped}>All terms for this mode are mapped.</p>
          ) : null}
        </>
      )}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#d13438" : "#0f6c3f" }}>
          {toast.message}
          <button onClick={() => setToast(null)} style={s.toastClose}>✕</button>
        </div>
      )}
    </section>
  );
}

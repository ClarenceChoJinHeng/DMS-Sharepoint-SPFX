import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IReconciliationProps } from "./IReconciliationProps";
import {
  loadMappedTermGuids,
  resolveFolderByPath,
  writeFolderMapping,
} from "../../../shared/dmsFolderMap";

/** One selectable folder-bearing term, with everything needed to map it. */
interface Candidate {
  termGuid: string;
  title: string; // "Account > AI Engineer"
  expectedPath: string; // computed server-relative path
  section: string; // "Departments" | "Projects"
  status: "" | "saving" | "saved" | "notfound" | "error";
  message: string;
}

interface Mode {
  key: string;
  termSetGuid: string;
  stagingFolder: string; // "Departments" | "Projects"
}

// Fallback modes — mirror Form.tsx DEFAULT_MODES. Reconciliation only needs
// the term set + staging folder per mode.
const DEFAULT_MODES: Mode[] = [
  { key: "department", termSetGuid: "eaba82e5-3e5f-4719-9a76-091f034ad407", stagingFolder: "Departments" },
  { key: "project", termSetGuid: "94ce322b-4515-4fda-8f50-35709f1f521d", stagingFolder: "Projects" },
];

type TermLite = { id: string; label: string };

const Reconciliation: React.FC<IReconciliationProps> = ({ context }) => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const serverRelative = context.pageContext.web.serverRelativeUrl;

  const [rows, setRows] = useState<Candidate[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [stagingLibrary, setStagingLibrary] = useState("Staging");

  const loadChildren = async (termSetId: string, parentId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${parentId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({
      id: t.id,
      label: t.labels[0].name,
    }));
  };

  const loadTops = async (termSetId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error(`Term set ${termSetId} returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({
      id: t.id,
      label: t.labels[0].name,
    }));
  };

  // Try to read stagingLibrary from DMS Config; fall back to "Staging".
  const loadStagingLibrary = async (): Promise<string> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,SettingValue&$filter=Title eq 'stagingLibrary'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return "Staging";
      const data = await res.json();
      return (data.value ?? [])[0]?.SettingValue || "Staging";
    } catch {
      return "Staging";
    }
  };

  useEffect(() => {
    (async () => {
      try {
        const lib = await loadStagingLibrary();
        setStagingLibrary(lib);
        const mapped = await loadMappedTermGuids(context.spHttpClient, siteUrl);
        const candidates: Candidate[] = [];

        for (const mode of DEFAULT_MODES) {
          const tops = await loadTops(mode.termSetGuid).catch(() => [] as TermLite[]);
          for (const top of tops) {
            // Walk every descendant; each is a folder a file could land in.
            const walk = async (parentId: string, ancestors: string[]): Promise<void> => {
              const children = await loadChildren(mode.termSetGuid, parentId);
              for (const child of children) {
                const chain = [...ancestors, child.label];
                if (!mapped.has(child.id)) {
                  candidates.push({
                    termGuid: child.id,
                    title: `${top.label} > ${chain.join(" > ")}`,
                    expectedPath: `${serverRelative}/${lib}/${mode.stagingFolder}/${top.label}/${chain.join("/")}`,
                    section: mode.stagingFolder,
                    status: "",
                    message: "",
                  });
                }
                await walk(child.id, chain);
              }
            };
            await walk(top.id, []);
          }
        }
        setRows(candidates);
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        setLoading(false);
      }
    })().catch((e: unknown) => {
      setError(String(e));
      setLoading(false);
    });
  }, []);

  const updateRow = (i: number, patch: Partial<Candidate>): void => {
    setRows((prev) => prev.map((r, idx) => (idx === i ? { ...r, ...patch } : r)));
  };

  const saveRow = async (i: number): Promise<void> => {
    const row = rows[i];
    updateRow(i, { status: "saving", message: "" });
    const resolved = await resolveFolderByPath(context.spHttpClient, siteUrl, row.expectedPath);
    if (!resolved) {
      updateRow(i, { status: "notfound", message: "Folder not found — check the path and try again." });
      return;
    }
    try {
      await writeFolderMapping(context.spHttpClient, siteUrl, {
        termGuid: row.termGuid,
        folderUniqueId: resolved.uniqueId,
        title: row.title,
        folderUrl: resolved.serverRelativeUrl,
        section: row.section,
      });
      updateRow(i, { status: "saved", message: `Mapped → ${resolved.uniqueId}` });
    } catch (e) {
      updateRow(i, { status: "error", message: e instanceof Error ? e.message : String(e) });
    }
  };

  const autoMatchAll = async (): Promise<void> => {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].status !== "saved") {
        await saveRow(i);
      }
    }
  };

  if (loading) return <div style={{ padding: 16 }}>Loading terms and existing mappings…</div>;
  if (error) return <div style={{ padding: 16, color: "#a00" }}>Error: {error}</div>;

  const remaining = rows.filter((r) => r.status !== "saved").length;

  return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>Folder Reconciliation</h2>
      <p>
        Library: <strong>{stagingLibrary}</strong> &nbsp;|&nbsp; Unmapped terms:{" "}
        <strong>{remaining}</strong> of {rows.length}
      </p>
      <button onClick={() => { autoMatchAll().catch(console.error); }} disabled={remaining === 0}>
        Auto-match all (resolve every expected path &amp; save)
      </button>
      <table style={{ width: "100%", marginTop: 12, borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ textAlign: "left", borderBottom: "2px solid #ccc" }}>
            <th style={{ padding: 6 }}>Term</th>
            <th style={{ padding: 6 }}>Folder path (edit if it doesn&apos;t match)</th>
            <th style={{ padding: 6 }}>Action</th>
            <th style={{ padding: 6 }}>Status</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={r.termGuid} style={{ borderBottom: "1px solid #eee", opacity: r.status === "saved" ? 0.5 : 1 }}>
              <td style={{ padding: 6 }}>{r.title}</td>
              <td style={{ padding: 6 }}>
                <input
                  style={{ width: "100%" }}
                  value={r.expectedPath}
                  disabled={r.status === "saved"}
                  onChange={(e) => updateRow(i, { expectedPath: e.target.value })}
                />
              </td>
              <td style={{ padding: 6 }}>
                <button onClick={() => { saveRow(i).catch(console.error); }} disabled={r.status === "saving" || r.status === "saved"}>
                  {r.status === "saving" ? "Saving…" : "Resolve & Save"}
                </button>
              </td>
              <td style={{ padding: 6, color: r.status === "saved" ? "#0a0" : r.status === "notfound" || r.status === "error" ? "#a00" : "#555" }}>
                {r.status === "saved" ? "✔ " : ""}{r.message}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default Reconciliation;

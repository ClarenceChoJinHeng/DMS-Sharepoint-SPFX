import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICrossSiteBrowserProps } from "./ICrossSiteBrowserProps";
import {
  parseSiteMapRows,
  siteOrigin,
  libraryRootPath,
  SegmentTarget,
} from "../../../shared/siteMap";

// Segment Documents browser (multi-site Phase 2, view side).
// Reads the "DMS Site Map" list to learn which segment lives on which site, then
// browses that segment's Documents library from here and opens files in place.
// Falls back to the property-pane target when the list is absent (test/manual use).

const SITE_MAP_LIST = "DMS Site Map";

interface FolderEntry {
  name: string;
  serverRelativeUrl: string;
  itemCount: number;
}
interface FileEntry {
  name: string;
  serverRelativeUrl: string;
  modified: string;
  size: number;
}
interface Crumb {
  name: string;
  path: string;
}

const isSystemFolder = (name: string): boolean =>
  name === "Forms" || name.startsWith("_");

const CrossSiteBrowser: React.FC<ICrossSiteBrowserProps> = ({
  context,
  targetSiteUrl,
  libraryName,
}) => {
  const dmsSiteUrl = context.pageContext.web.absoluteUrl;

  const [targets, setTargets] = useState<SegmentTarget[]>([]);
  const [selected, setSelected] = useState(0);
  const [configLoading, setConfigLoading] = useState(true);
  const [configMsg, setConfigMsg] = useState("");

  const [folders, setFolders] = useState<FolderEntry[]>([]);
  const [files, setFiles] = useState<FileEntry[]>([]);
  const [crumbs, setCrumbs] = useState<Crumb[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const activeTarget: SegmentTarget | undefined =
    targets[Math.min(selected, targets.length - 1)];

  const load = useCallback(
    async (siteUrl: string, path: string): Promise<void> => {
      setLoading(true);
      setError("");
      try {
        // Alias form (@f) — never inline the encoded path as a quoted literal.
        const foldersUrl =
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Folders` +
          `?@f='${encodeURIComponent(path)}'&$select=Name,ServerRelativeUrl,ItemCount&$orderby=Name`;
        const filesUrl =
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Files` +
          `?@f='${encodeURIComponent(path)}'&$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=Name`;

        const [foldersRes, filesRes]: SPHttpClientResponse[] = await Promise.all([
          context.spHttpClient.get(foldersUrl, SPHttpClient.configurations.v1, {
            headers: { Accept: "application/json;odata=nometadata" },
          }),
          context.spHttpClient.get(filesUrl, SPHttpClient.configurations.v1, {
            headers: { Accept: "application/json;odata=nometadata" },
          }),
        ]);

        if (!foldersRes.ok) {
          throw new Error(
            `Read failed (HTTP ${foldersRes.status}). Check the segment site URL, its Documents library, and that you have access there.`,
          );
        }
        const foldersJson = await foldersRes.json();
        const nextFolders: FolderEntry[] = (foldersJson.value ?? [])
          .filter((f: { Name: string }) => !isSystemFolder(f.Name))
          .map((f: { Name: string; ServerRelativeUrl: string; ItemCount: number }) => ({
            name: f.Name,
            serverRelativeUrl: f.ServerRelativeUrl,
            itemCount: f.ItemCount ?? 0,
          }));

        let nextFiles: FileEntry[] = [];
        if (filesRes.ok) {
          const filesJson = await filesRes.json();
          nextFiles = (filesJson.value ?? []).map(
            (f: {
              Name: string;
              ServerRelativeUrl: string;
              TimeLastModified: string;
              Length: string;
            }) => ({
              name: f.Name,
              serverRelativeUrl: f.ServerRelativeUrl,
              modified: f.TimeLastModified,
              size: Number(f.Length ?? 0),
            }),
          );
        }

        setFolders(nextFolders);
        setFiles(nextFiles);
      } catch (e) {
        setFolders([]);
        setFiles([]);
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        setLoading(false);
      }
    },
    [context],
  );

  // 1) Load the DMS Site Map config once; fall back to the property-pane target.
  useEffect(() => {
    (async (): Promise<void> => {
      setConfigLoading(true);
      const fallback = (msg: string): void => {
        const url = (targetSiteUrl ?? "").trim().replace(/\/+$/, "");
        if (url) {
          setTargets([
            {
              segment: "(manual target)",
              siteUrl: url,
              documentsLibrary: (libraryName ?? "").trim() || "Documents",
            },
          ]);
          setConfigMsg("");
        } else {
          setTargets([]);
          setConfigMsg(msg);
        }
      };
      try {
        const url =
          `${dmsSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(SITE_MAP_LIST)}')` +
          `/items?$select=Title,SiteUrl,DocumentsLibrary&$top=500`;
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!res.ok) {
          fallback(
            `The "${SITE_MAP_LIST}" list was not found (HTTP ${res.status}). Create it (columns: Title, SiteUrl, DocumentsLibrary) or set a manual target in the web part properties.`,
          );
          return;
        }
        const data = await res.json();
        const rows = (data.value ?? []).map(
          (r: { Title?: string; SiteUrl?: string; DocumentsLibrary?: string }) => ({
            title: r.Title ?? "",
            siteUrl: r.SiteUrl ?? "",
            documentsLibrary: r.DocumentsLibrary ?? "",
          }),
        );
        const parsed = parseSiteMapRows(rows);
        if (parsed.length > 0) {
          setTargets(parsed);
          setConfigMsg("");
        } else {
          fallback(
            `The "${SITE_MAP_LIST}" list has no usable rows. Add a row (Title, SiteUrl, DocumentsLibrary) or set a manual target in the web part properties.`,
          );
        }
      } catch (e) {
        fallback(e instanceof Error ? e.message : String(e));
      } finally {
        setConfigLoading(false);
      }
    })().catch(() => undefined);
  }, [targetSiteUrl, libraryName]);

  // 2) When the selected target changes, reset to its Documents root and load.
  useEffect(() => {
    if (!activeTarget) return;
    const root = libraryRootPath(activeTarget.siteUrl, activeTarget.documentsLibrary);
    setCrumbs([{ name: activeTarget.documentsLibrary, path: root }]);
    load(activeTarget.siteUrl, root).catch(() => undefined);
  }, [targets, selected]);

  const openFolder = (folder: FolderEntry): void => {
    if (!activeTarget) return;
    setCrumbs((prev) => [...prev, { name: folder.name, path: folder.serverRelativeUrl }]);
    load(activeTarget.siteUrl, folder.serverRelativeUrl).catch(() => undefined);
  };

  const goToCrumb = (index: number): void => {
    if (!activeTarget) return;
    const target = crumbs[index];
    setCrumbs((prev) => prev.slice(0, index + 1));
    load(activeTarget.siteUrl, target.path).catch(() => undefined);
  };

  const fmtSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const origin = activeTarget ? siteOrigin(activeTarget.siteUrl) : "";

  if (configLoading) {
    return (
      <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
        Loading segment configuration…
      </div>
    );
  }

  if (!activeTarget) {
    return (
      <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
        <h2 style={{ marginBottom: 4 }}>Segment Documents</h2>
        <p style={{ color: "#a00" }}>{configMsg}</p>
      </div>
    );
  }

  return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
      <h2 style={{ marginBottom: 4 }}>Segment Documents</h2>

      {/* Segment picker — only when more than one segment site is configured. */}
      {targets.length > 1 && (
        <div style={{ margin: "8px 0" }}>
          <label style={{ fontSize: 13, marginRight: 6 }}>Segment:</label>
          <select
            value={selected}
            onChange={(e) => setSelected(Number(e.target.value))}
          >
            {targets.map((t, i) => (
              <option key={t.siteUrl} value={i}>
                {t.segment}
              </option>
            ))}
          </select>
        </div>
      )}

      <p style={{ marginTop: 0, fontSize: 12, color: "#666" }}>
        Reading <strong>{activeTarget.documentsLibrary}</strong> from{" "}
        <code>{activeTarget.siteUrl}</code>
      </p>

      {/* Breadcrumb */}
      <div style={{ margin: "8px 0", fontSize: 14 }}>
        {crumbs.map((c, i) => (
          <span key={c.path}>
            {i > 0 && <span style={{ color: "#999" }}> / </span>}
            <button
              onClick={() => goToCrumb(i)}
              disabled={i === crumbs.length - 1}
              style={{
                border: "none",
                background: "none",
                color: i === crumbs.length - 1 ? "#333" : "#0078d4",
                cursor: i === crumbs.length - 1 ? "default" : "pointer",
                padding: 0,
                font: "inherit",
              }}
            >
              {c.name}
            </button>
          </span>
        ))}
      </div>

      {loading && <div style={{ padding: 8 }}>Loading…</div>}
      {error && <div style={{ padding: 8, color: "#a00" }}>Error: {error}</div>}

      {!loading && !error && (
        <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 8 }}>
          <thead>
            <tr style={{ textAlign: "left", borderBottom: "2px solid #ccc" }}>
              <th style={{ padding: 6 }}>Name</th>
              <th style={{ padding: 6 }}>Type</th>
              <th style={{ padding: 6 }}>Modified</th>
              <th style={{ padding: 6 }}>Size</th>
            </tr>
          </thead>
          <tbody>
            {folders.map((f) => (
              <tr key={f.serverRelativeUrl} style={{ borderBottom: "1px solid #eee" }}>
                <td style={{ padding: 6 }}>
                  <button
                    onClick={() => openFolder(f)}
                    style={{
                      border: "none",
                      background: "none",
                      color: "#0078d4",
                      cursor: "pointer",
                      padding: 0,
                      font: "inherit",
                    }}
                  >
                    📁 {f.name}
                  </button>
                </td>
                <td style={{ padding: 6, color: "#666" }}>Folder</td>
                <td style={{ padding: 6 }} />
                <td style={{ padding: 6, color: "#666" }}>{f.itemCount} item(s)</td>
              </tr>
            ))}
            {files.map((f) => (
              <tr key={f.serverRelativeUrl} style={{ borderBottom: "1px solid #eee" }}>
                <td style={{ padding: 6 }}>
                  <a
                    href={`${origin}${encodeURI(f.serverRelativeUrl)}`}
                    target="_blank"
                    rel="noreferrer"
                  >
                    📄 {f.name}
                  </a>
                </td>
                <td style={{ padding: 6, color: "#666" }}>File</td>
                <td style={{ padding: 6, color: "#666" }}>
                  {f.modified ? new Date(f.modified).toLocaleDateString() : ""}
                </td>
                <td style={{ padding: 6, color: "#666" }}>{fmtSize(f.size)}</td>
              </tr>
            ))}
            {folders.length === 0 && files.length === 0 && (
              <tr>
                <td colSpan={4} style={{ padding: 12, color: "#999" }}>
                  This folder is empty.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      )}
    </div>
  );
};

export default CrossSiteBrowser;

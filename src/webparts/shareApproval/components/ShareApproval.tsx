// Admin-only queue for DMS Share Requests. Lists Pending rows; Approve grants the
// recipient the chosen role on the item (resolved by UniqueId) and sets Status.
// Each card shows the item's LOCATION breadcrumb and, for folders, an expandable
// tree of its contents (lazy-loaded by UniqueId) so the approver sees what is being
// shared without opening the item.
import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IShareApprovalProps } from "./IShareApprovalProps";
import { itemAbsoluteUrl, itemBreadcrumb } from "../../../shared/shareGuard";

const LIST = "DMS Share Requests";

const ICON_FOLDER = "📁"; // 📁
const ICON_FILE = "📄"; // 📄
const CHEV_OPEN = "▾"; // ▾
const CHEV_SHUT = "▸"; // ▸

interface Row {
  id: number; title: string; itemUrl: string; itemUniqueId: string;
  itemType: "File" | "Folder"; recipientId: number; recipientTitle: string;
  accessLevel: "Read" | "Edit"; reason: string; status: string;
}

interface ChildNode { name: string; uniqueId: string; type: "File" | "Folder"; }

type FetchChildren = (folderId: string) => Promise<ChildNode[]>;

// Recursive, lazy folder tree. Loads a folder's children on mount; each subfolder
// expands to mount another FolderTree only when clicked.
const FolderTree: React.FC<{ folderId: string; fetchChildren: FetchChildren; depth?: number }> = (
  { folderId, fetchChildren, depth = 0 },
) => {
  const [kids, setKids] = useState<ChildNode[] | undefined>(undefined);
  const [err, setErr] = useState("");
  const [open, setOpen] = useState<Record<string, boolean>>({});

  useEffect(() => {
    let alive = true;
    fetchChildren(folderId)
      .then((k) => { if (alive) setKids(k); })
      .catch((e) => { if (alive) setErr((e as Error).message); });
    return () => { alive = false; };
  }, [folderId, fetchChildren]);

  const rowIndent = depth * 16 + 18;
  if (err) return <div style={{ marginLeft: rowIndent, fontSize: 12, color: "#a00" }}>Couldn&rsquo;t load: {err}</div>;
  if (kids === undefined) return <div style={{ marginLeft: rowIndent, fontSize: 12, color: "#999" }}>Loading&hellip;</div>;
  if (kids.length === 0) return <div style={{ marginLeft: rowIndent, fontSize: 12, color: "#999" }}>(empty)</div>;

  return (
    <div>
      {kids.map((k) =>
        k.type === "Folder" ? (
          <div key={k.uniqueId}>
            <div
              style={{ marginLeft: rowIndent, fontSize: 13, cursor: "pointer", userSelect: "none" }}
              onClick={() => setOpen((o) => ({ ...o, [k.uniqueId]: !o[k.uniqueId] }))}
            >
              <span style={{ display: "inline-block", width: 12 }}>{open[k.uniqueId] ? CHEV_OPEN : CHEV_SHUT}</span>
              {ICON_FOLDER} {k.name}
            </div>
            {open[k.uniqueId] && <FolderTree folderId={k.uniqueId} fetchChildren={fetchChildren} depth={depth + 1} />}
          </div>
        ) : (
          <div key={k.uniqueId} style={{ marginLeft: rowIndent + 12, fontSize: 13 }}>
            {ICON_FILE} {k.name}
          </div>
        ),
      )}
    </div>
  );
};

const badgeStyle: React.CSSProperties = {
  fontSize: 11, padding: "1px 8px", borderRadius: 10, background: "#eef2f7",
  color: "#33475b", border: "1px solid #d9e1ec",
};

const ShareApproval: React.FC<IShareApprovalProps> = ({ context }) => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const webServerRelativeUrl = context.pageContext.web.serverRelativeUrl;
  const origin = new URL(siteUrl).origin;
  const [rows, setRows] = useState<Row[]>([]);
  const [busy, setBusy] = useState<number | null>(null);
  const [msg, setMsg] = useState("");
  const [contentsOpen, setContentsOpen] = useState<Record<number, boolean>>({});

  const load = useCallback(async (): Promise<void> => {
    const url = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items`
      + `?$select=Id,Title,ItemUrl,ItemUniqueId,ItemType,AccessLevel,Reason,Status,RecipientId,Recipient/Title`
      + `&$expand=Recipient&$filter=Status eq 'Pending'&$orderby=Id desc&$top=200`;
    const res: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) { setMsg(`Could not load requests (HTTP ${res.status}).`); return; }
    const data = await res.json();
    setRows((data.value ?? []).map((r: Record<string, unknown>) => ({
      id: r.Id as number, title: r.Title as string, itemUrl: r.ItemUrl as string, itemUniqueId: r.ItemUniqueId as string,
      itemType: r.ItemType as "File" | "Folder", recipientId: r.RecipientId as number, recipientTitle: (r.Recipient as {Title?: string})?.Title ?? "",
      accessLevel: r.AccessLevel as "Read" | "Edit", reason: r.Reason as string, status: r.Status as string,
    })) as Row[]);
  }, [context, siteUrl]);

  useEffect(() => { load().catch(() => undefined); }, [load]);

  // Lazy child loader used by the folder tree. Navigates by UniqueId (GetFolderById)
  // so deep/long server-relative paths never trigger the encoded-path HTTP 400.
  const fetchChildren = useCallback<FetchChildren>(async (folderId) => {
    const build = (verb: string): string =>
      `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/${verb}?$select=Name,UniqueId&$orderby=Name&$top=500`;
    const opts = { headers: { Accept: "application/json;odata=nometadata" } };
    const [foldersRes, filesRes] = await Promise.all([
      context.spHttpClient.get(build("Folders"), SPHttpClient.configurations.v1, opts),
      context.spHttpClient.get(build("Files"), SPHttpClient.configurations.v1, opts),
    ]);
    if (!foldersRes.ok && !filesRes.ok) throw new Error(`HTTP ${foldersRes.status}/${filesRes.status}`);
    const folders = foldersRes.ok ? ((await foldersRes.json()).value ?? []) : [];
    const files = filesRes.ok ? ((await filesRes.json()).value ?? []) : [];
    return [
      ...folders.map((f: Record<string, unknown>) => ({ name: f.Name as string, uniqueId: f.UniqueId as string, type: "Folder" as const })),
      ...files.map((f: Record<string, unknown>) => ({ name: f.Name as string, uniqueId: f.UniqueId as string, type: "File" as const })),
    ];
  }, [context, siteUrl]);

  const roleDefId = async (name: string): Promise<number | undefined> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/roledefinitions?$select=Id,Name&$filter=Name eq '${name}'`,
      SPHttpClient.configurations.v1, { headers: { Accept: "application/json;odata=nometadata" } });
    if (!res.ok) return undefined;
    return ((await res.json()).value ?? [])[0]?.Id as number | undefined;
  };

  // item path segment for the REST verb: files use GetFileById, folders GetFolderById
  const itemListItemPath = (r: Row): string =>
    r.itemType === "Folder"
      ? `GetFolderById(guid'${r.itemUniqueId}')/ListItemAllFields`
      : `GetFileById(guid'${r.itemUniqueId}')/ListItemAllFields`;

  const setStatus = async (id: number, status: string): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items(${id})`,
      SPHttpClient.configurations.v1, {
        headers: { Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "IF-MATCH": "*", "X-HTTP-Method": "MERGE" },
        body: JSON.stringify({ Status: status, DecidedOn: new Date().toISOString() }),
      });
    if (!res.ok) throw new Error(`status update HTTP ${res.status}`);
  };

  const approve = async (r: Row): Promise<void> => {
    setBusy(r.id); setMsg("");
    try {
      const rdid = await roleDefId(r.accessLevel === "Edit" ? "Edit" : "Read");
      if (rdid === undefined) throw new Error(`No "${r.accessLevel}" role definition on site.`);
      const base = `${siteUrl}/_api/web/${itemListItemPath(r)}`;
      // Preserve existing access, add the recipient: break with copyRoleAssignments=true.
      await context.spHttpClient.post(
        `${base}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`,
        SPHttpClient.configurations.v1, { headers: { Accept: "application/json;odata=nometadata" } });
      await context.spHttpClient.post(
        `${base}/roleassignments/addroleassignment(principalid=${r.recipientId},roledefid=${rdid})`,
        SPHttpClient.configurations.v1, { headers: { Accept: "application/json;odata=nometadata" } });
      await setStatus(r.id, "Approved");
      setMsg(`Approved — ${r.recipientTitle} now has ${r.accessLevel} on ${r.title}.`);
      await load();
    } catch (e) { setMsg(`Approve failed: ${(e as Error).message}`); }
    finally { setBusy(null); }
  };

  const reject = async (r: Row): Promise<void> => {
    setBusy(r.id); try { await setStatus(r.id, "Rejected"); await load(); }
    finally { setBusy(null); }
  };

  return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>Share requests — pending approval</h2>
      {msg && <p style={{ color: "#0f6c3f" }}>{msg}</p>}
      {rows.length === 0 && <p style={{ color: "#999" }}>No pending requests.</p>}
      {rows.map((r) => {
        const crumbs = itemBreadcrumb(r.itemUrl, webServerRelativeUrl);
        return (
          <div key={r.id} style={{ border: "1px solid #e0e0e0", borderRadius: 6, padding: 12, marginBottom: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <span style={{ fontSize: 15, fontWeight: 600 }}>
                {r.itemType === "Folder" ? ICON_FOLDER : ICON_FILE} {r.title}
              </span>
              <span style={badgeStyle}>{r.itemType}</span>
              <span style={badgeStyle}>{r.accessLevel}</span>
            </div>
            <div style={{ fontSize: 12, color: "#666", marginTop: 3 }}>
              {crumbs.length ? crumbs.join("  ›  ") : r.itemUrl}
            </div>
            <div style={{ fontSize: 13, marginTop: 4 }}>
              <strong>Recipient:</strong> {r.recipientTitle || "—"}
              {r.reason ? <span>{"  ·  "}<strong>Reason:</strong> {r.reason}</span> : null}
            </div>

            {r.itemType === "Folder" && (
              <div style={{ marginTop: 8 }}>
                <div
                  style={{ fontSize: 13, cursor: "pointer", color: "#0078d4", userSelect: "none" }}
                  onClick={() => setContentsOpen((s) => ({ ...s, [r.id]: !s[r.id] }))}
                >
                  {contentsOpen[r.id] ? CHEV_OPEN : CHEV_SHUT} Contents
                </div>
                {contentsOpen[r.id] && (
                  <div style={{ marginTop: 4 }}>
                    <FolderTree folderId={r.itemUniqueId} fetchChildren={fetchChildren} />
                  </div>
                )}
              </div>
            )}

            <div style={{ marginTop: 10, display: "flex", gap: 8, alignItems: "center" }}>
              <button disabled={busy === r.id} onClick={() => { approve(r).catch(() => undefined); }}>Approve</button>
              <button disabled={busy === r.id} onClick={() => { reject(r).catch(() => undefined); }}>Reject</button>
              <a href={itemAbsoluteUrl(origin, r.itemUrl)} target="_blank" rel="noreferrer" style={{ fontSize: 12, marginLeft: "auto" }}>open item</a>
            </div>
          </div>
        );
      })}
    </div>
  );
};
export default ShareApproval;

// Admin-only queue for DMS Share Requests. Lists Pending rows; Approve grants the
// recipient the chosen role on the item (resolved by UniqueId) and sets Status.
import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IShareApprovalProps } from "./IShareApprovalProps";
import { itemAbsoluteUrl } from "../../../shared/shareGuard";

const LIST = "DMS Share Requests";

interface Row {
  id: number; title: string; itemUrl: string; itemUniqueId: string;
  itemType: "File" | "Folder"; recipientId: number; recipientTitle: string;
  accessLevel: "Read" | "Edit"; reason: string; status: string;
}

const ShareApproval: React.FC<IShareApprovalProps> = ({ context }) => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const origin = new URL(siteUrl).origin;
  const [rows, setRows] = useState<Row[]>([]);
  const [busy, setBusy] = useState<number | null>(null);
  const [msg, setMsg] = useState("");

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
      {rows.map(r => (
        <div key={r.id} style={{ border: "1px solid #e0e0e0", borderRadius: 6, padding: 12, marginBottom: 8 }}>
          <strong>{r.title}</strong> &rarr; {r.recipientTitle} ({r.accessLevel})
          <div style={{ fontSize: 12, color: "#666" }}>{r.reason}</div>
          <a href={itemAbsoluteUrl(origin, r.itemUrl)} target="_blank" rel="noreferrer" style={{ fontSize: 12 }}>open item</a>
          <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
            <button disabled={busy === r.id} onClick={() => { approve(r).catch(() => undefined); }}>Approve</button>
            <button disabled={busy === r.id} onClick={() => { reject(r).catch(() => undefined); }}>Reject</button>
          </div>
        </div>
      ))}
    </div>
  );
};
export default ShareApproval;

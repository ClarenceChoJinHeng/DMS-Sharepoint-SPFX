/**
 * CRS Requests — the approver's queue.
 *
 * Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md
 *
 * A PIC cannot delete or share in `Documents`; they raise a request from My Submissions and a Head of
 * Unit decides it here.
 *
 * THE APPROVAL EXECUTES IN THE APPROVER'S OWN SESSION. When they press Approve, THEIR browser recycles
 * the file or grants the access, because they hold the rights the requester lacks. No service account,
 * no Power Automate, nothing acting on anyone's behalf — so the audit row names the person who
 * actually did it, and an approval can never exceed the approver's own rights. It fails loudly instead.
 *
 * The rules live in `shared/requests.ts`, under test. This file is the screen and the requests.
 */
import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { IRequestsProps } from "./IRequestsProps";
import { cachedListTitle, LIST_SUFFIX, libraryTitle } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { EVENT } from "../../../shared/auditLog";
import { normalizeRoleValue } from "../../../shared/groupMapModel";
import {
  RequestRow,
  RequestStatus,
  applyDecision,
  canDecide,
  counts,
  decisionSummary,
  isExternal,
  queueFor,
} from "../../../shared/requests";

const GET = { Accept: "application/json;odata=nometadata" };

const s: Record<string, React.CSSProperties> = {
  wrap:    { fontFamily: '"Segoe UI", system-ui, sans-serif', color: "#242424", fontSize: 13, lineHeight: 1.5 },
  h2:      { fontSize: 22, fontWeight: 600, margin: "0 0 6px" },
  sub:     { fontSize: 13, color: "#5f5f5f", margin: "0 0 20px", lineHeight: 1.55 },
  card:    { border: "1px solid #e1e1e1", borderRadius: 8, background: "#fff", padding: 16, marginBottom: 16 },
  head:    { fontSize: 15, fontWeight: 600, margin: "0 0 10px" },
  row:     { border: "1px solid #eceaea", borderRadius: 8, padding: "12px 14px", marginBottom: 10, background: "#fafafa" },
  rowTop:  { display: "flex", gap: 10, alignItems: "baseline", flexWrap: "wrap" },
  name:    { fontWeight: 600, flex: "1 1 240px", wordBreak: "break-word" },
  pill:    { fontSize: 11, fontWeight: 600, borderRadius: 10, padding: "2px 8px" },
  meta:    { fontSize: 11.5, color: "#6b7a71", marginTop: 4 },
  reason:  { marginTop: 8, fontSize: 12.5, background: "#fff", border: "1px solid #eceaea", borderRadius: 6, padding: "8px 10px" },
  actions: { display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" },
  approve: { padding: "6px 16px", fontSize: 12.5, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  reject:  { padding: "6px 16px", fontSize: 12.5, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  off:     { padding: "6px 16px", fontSize: 12.5, background: "#e6e6e6", color: "#9a9a9a", border: "none", borderRadius: 4, cursor: "not-allowed" },
  ghost:   { padding: "6px 14px", fontSize: 12.5, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  warn:    { border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 8, padding: "12px 14px", fontSize: 12.5, color: "#8a4b00", marginBottom: 16 },
  err:     { border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 8, padding: "12px 14px", fontSize: 12.5, color: "#a4262c", marginBottom: 16 },
  ok:      { border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 8, padding: "12px 14px", fontSize: 12.5, color: "#1c4d33", marginBottom: 16 },
  quiet:   { fontSize: 12.5, color: "#767676" },
  input:   { width: "100%", maxWidth: 460, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  modalBg: { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 },
  modal:   { background: "#fff", borderRadius: 8, padding: 20, width: "min(560px, 94vw)", maxHeight: "86vh", overflowY: "auto" },
};

const PILL: Record<RequestStatus, React.CSSProperties> = {
  Pending:  { background: "#fff2df", color: "#8a4b00" },
  Approved: { background: "#0f6c3f", color: "#fff" },
  Rejected: { background: "#f0eff0", color: "#767676" },
  Failed:   { background: "#fdf3f4", color: "#a4262c" },
};

/** The columns the list must have. Internal names are space-free, as with the audit log. */
const COLUMNS: Array<{ name: string; type: number; note?: boolean }> = [
  { name: "RequestType", type: 2 },
  { name: "Status", type: 2 },
  { name: "ItemUniqueId", type: 2 },
  { name: "ItemName", type: 2 },
  { name: "ItemUrl", type: 2 },
  { name: "Segment", type: 2 },
  { name: "Unit", type: 2 },
  { name: "UnitTermGuid", type: 2 },
  { name: "RequestedBy", type: 2 },
  { name: "RequestedAt", type: 4 },
  { name: "Reason", type: 3, note: true },
  { name: "ShareWith", type: 2 },
  { name: "SharePermission", type: 2 },
  { name: "ExpiresAt", type: 4 },
  { name: "DecidedBy", type: 2 },
  { name: "DecidedAt", type: 4 },
  { name: "DecisionNote", type: 3, note: true },
];

type Load<T> =
  | { state: "loading" }
  | { state: "error"; message: string }
  | { state: "ready"; value: T };

export default function Requests({ context }: IRequestsProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const me = (context.pageContext.user.email ?? "").toLowerCase();

  const [rows, setRows] = useState<Load<RequestRow[]>>({ state: "loading" });
  const [approverUnits, setApproverUnits] = useState<string[]>([]);
  const [tenantDomains, setTenantDomains] = useState<string[]>([]);
  const [listMissing, setListMissing] = useState(false);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<string | undefined>(undefined);
  const [deciding, setDeciding] = useState<{ row: RequestRow; approve: boolean } | undefined>(undefined);
  const [note, setNote] = useState("");

  const listUrl = (): string =>
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')`;

  /* JSON light with NO `__metadata`, and BOTH header halves saying nometadata. Verbose on one half
     gives "The property '__metadata' does not exist on type 'SP.List'"; verbose on both gives
     "Parsing JSON Light feeds or entries in requests without entity set is not supported", because
     SPFx attaches its own OData version header. `odata-version: ""` is needed for the same reason:
     SPFx injects 4.0, under which SharePoint cannot infer the entity set for a JSON-light entry.
     All learned on the audit log — do NOT "restore" ListItemEntityTypeFullName here. */
  const writeHeaders = {
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata",
    "odata-version": "",
  };

  const post = async (url: string, body?: unknown): Promise<SPHttpClientResponse> =>
    context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: writeHeaders,
      body: body === undefined ? undefined : JSON.stringify(body),
    });

  /* ── Load ─────────────────────────────────────────────────────────────── */

  const load = async (): Promise<void> => {
    await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);

    // Which units does this person approve for? From their group memberships matched against the
    // Group Map, keyed on the TERM GUID — a renamed unit must not silently empty someone's queue.
    const units: string[] = [];
    try {
      // The CURRENT USER's groups, not every group on the site — the same endpoint the upload form
      // uses. `fetchAllSiteGroups` would list all of them and make every unit look like this
      // person's to approve.
      const mineRes = await context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser/groups?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      const myIds: string[] = mineRes.ok
        ? ((await mineRes.json()).value ?? []).map((g: { Id?: number }) => String(g.Id ?? ""))
        : [];
      const mapRes = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items` +
          `?$select=GroupId,Role,UnitTermGuid,Segment&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (mapRes.ok) {
        const data = await mapRes.json();
        for (const r of (data.value ?? []) as Array<{ GroupId?: string; Role?: string; UnitTermGuid?: string }>) {
          // APR is the approver role. A HoU also holds DEL and SHARE, but APR is what says "this
          // person decides for this unit".
          /* APR ALONE since 2026-08-17: APRHC is retired, and `normalizeRoleValue` aliases a stored
             "APRHC" to "APR", so a legacy row still routes here. Any approver who is a Head of Unit
             now reaches HC through APR.

             normalizeRoleValue, NOT a raw toUpperCase — which is what this was, and it was a latent
             bug of the same class formModel.ts documents. The Group Map's Role is often the LONG form
             ("APPROVER", as "UPLOADER" was found live on 2026-08-07), and a raw compare against "APR"
             skips such a row. The consequence is silent and total: that unit's approver queue stays
             empty for ever while requests pile up behind it, with nothing on screen to say so. */
          const role = normalizeRoleValue(r.Role ?? "");
          if (role !== "APR") continue;
          if (myIds.indexOf(String(r.GroupId ?? "")) === -1) continue;
          const guid = (r.UnitTermGuid ?? "").trim();
          if (guid && units.indexOf(guid) === -1) units.push(guid);
        }
      }
    } catch {
      // An unreadable Group Map means an empty queue, not a broken page — said on screen below.
    }
    setApproverUnits(units);

    // The tenant's own domains, so `isExternal` can tell inside from outside. Defaults to the
    // signed-in user's own domain, which is not a guess — they are signed in to this tenant. A config
    // row adds more (multi-domain tenants) with no redeploy.
    const domains: string[] = [];
    const at = me.lastIndexOf("@");
    if (at > -1) domains.push(me.slice(at + 1));
    try {
      const cfg = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
          `?$select=Title,SettingValue&$filter=Title eq 'tenantDomains'&$top=5`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (cfg.ok) {
        const data = await cfg.json();
        for (const r of (data.value ?? []) as Array<{ SettingValue?: string }>) {
          for (const d of (r.SettingValue ?? "").split(/[,;\s]+/)) {
            const clean = d.trim().toLowerCase();
            if (clean && domains.indexOf(clean) === -1) domains.push(clean);
          }
        }
      }
    } catch {
      /* the signed-in user's own domain stands on its own */
    }
    setTenantDomains(domains);

    try {
      const res = await context.spHttpClient.get(
        `${listUrl()}/items?$select=Id,RequestType,Status,ItemUniqueId,ItemName,ItemUrl,Segment,Unit,` +
          `UnitTermGuid,RequestedBy,RequestedAt,Reason,ShareWith,SharePermission,ExpiresAt,DecidedBy,` +
          `DecidedAt,DecisionNote&$top=2000&$orderby=Id desc`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (res.status === 404) {
        setListMissing(true);
        setRows({ state: "ready", value: [] });
        return;
      }
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setListMissing(false);
      setRows({ state: "ready", value: ((data.value ?? []) as Record<string, string>[]).map(fromListItem) });
    } catch (e) {
      // "Could not read" is never rendered as "there are none" — an approver told their queue is
      // empty stops looking, and a request sits unanswered.
      setRows({ state: "error", message: (e as Error).message });
    }
  };

  useEffect(() => {
    load().catch((e) => setRows({ state: "error", message: (e as Error).message }));
  }, [siteUrl]);

  /* ── Provisioning ─────────────────────────────────────────────────────── */

  const provision = async (): Promise<void> => {
    setBusy(true);
    setNotice(undefined);
    try {
      const title = cachedListTitle(LIST_SUFFIX.requests);
      const made = await post(`${siteUrl}/_api/web/lists`, {
        Title: title,
        BaseTemplate: 100,
        Description: "Deletion and share requests raised by uploaders and decided by the Head of Unit.",
      });
      if (!made.ok) throw new Error(`creating the list failed — HTTP ${made.status}`);

      for (const c of COLUMNS) {
        const field: Record<string, unknown> = { Title: c.name, FieldTypeKind: c.type };
        if (c.note) field.NumberOfLines = 6;
        // A failed column is reported, never swallowed: a list missing ShareWith would accept a share
        // request and silently drop who it was for.
        const r = await post(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/fields`, field);
        if (!r.ok) throw new Error(`adding the "${c.name}" column failed — HTTP ${r.status}`);
      }
      setNotice(
        `Created "${title}". Its permissions are NOT set automatically — an administrator should ` +
          "restrict who may edit it, the same as the audit log.",
      );
      await load();
    } catch (e) {
      setNotice(`Could not finish: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  /* ── Carrying out a decision ──────────────────────────────────────────── */

  /**
   * Recycle, never delete outright — restorable for 93 days, which is what makes approving a deletion
   * reasonable at all. Resolved by UniqueId, so a rename or move since the request was raised does not
   * matter.
   */
  const performDeletion = async (row: RequestRow): Promise<string | undefined> => {
    const res = await post(`${siteUrl}/_api/web/GetFileById(guid'${row.itemUniqueId}')/recycle()`);
    if (res.ok) return undefined;
    if (res.status === 404) return "That document no longer exists — it may already have been deleted.";
    if (res.status === 403) return "You do not have permission to delete that document.";
    return `The document could not be deleted (HTTP ${res.status}).`;
  };

  /**
   * One endpoint for internal and external alike.
   *
   * `SP.Web.ShareObject` is what SharePoint's own Share dialog calls, so it honours the tenant and
   * site sharing settings rather than working around them. If external sharing is off there, this
   * fails and says so — the correct outcome, and far better than a grant that half-works.
   */
  const performShare = async (row: RequestRow): Promise<string | undefined> => {
    const people = (row.shareWith ?? []).map((e) => ({ Key: e }));
    if (people.length === 0) return "No recipients were recorded on this request.";
    if (!row.itemUrl) return "This request has no document address recorded.";

    // SharePoint's own role values: 1073741826 = View, 1073741827 = Edit.
    const roleValue = row.sharePermission === "Edit" ? "role:1073741827" : "role:1073741826";
    const res = await post(`${siteUrl}/_api/SP.Web.ShareObject`, {
      url: `${window.location.origin}${row.itemUrl}`,
      peoplePickerInput: JSON.stringify(people),
      roleValue,
      groupId: 0,
      propagateAcl: false,
      sendEmail: true,
      includeAnonymousLinkInEmail: false,
      emailSubject: `A document has been shared with you: ${row.itemName}`,
      emailBody: row.reason ?? "",
      useSimplifiedRoles: true,
    });
    if (!res.ok) {
      if (res.status === 403) return "You do not have permission to share that document.";
      return `The document could not be shared (HTTP ${res.status}).`;
    }
    // HTTP 200 does NOT mean it worked — the per-recipient result is in the body, exactly as with
    // validateUpdateListItem (gotcha #4). A refusal by tenant policy arrives here, not as a status.
    try {
      const body = await res.json();
      const results = (body?.value ?? []) as Array<{ Status?: boolean; Message?: string; User?: string }>;
      const failed = results.filter((r) => r && r.Status === false);
      if (failed.length > 0) {
        return failed.map((f) => `${f.User ?? "recipient"}: ${f.Message ?? "refused"}`).join("; ");
      }
    } catch {
      /* an unreadable body after a 200 counts as success — the grant is what matters */
    }
    return undefined;
  };

  const decide = async (row: RequestRow, approve: boolean, text: string): Promise<void> => {
    setBusy(true);
    setNotice(undefined);
    try {
      let failure: string | undefined;
      if (approve) {
        failure = row.type === "Deletion" ? await performDeletion(row) : await performShare(row);
      }
      const decided = applyDecision(row, {
        approve, by: me, at: new Date().toISOString(), note: text, failure,
      });

      const upd = await context.spHttpClient.post(
        `${listUrl()}/items(${row.id})`,
        SPHttpClient.configurations.v1,
        {
          headers: { ...writeHeaders, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
          body: JSON.stringify({
            Status: decided.status,
            DecidedBy: decided.decidedBy,
            DecidedAt: decided.decidedAt,
            DecisionNote: decided.decisionNote,
          }),
        },
      );
      // The row could not be updated but the ACTION already happened. Said loudly: silently leaving
      // it Pending invites a second approver to do the same thing again.
      const rowWritten = upd.ok;

      writeAudit(context.spHttpClient, siteUrl, {
        event: approve ? EVENT.requestApproved : EVENT.requestRejected,
        outcome: failure || !rowWritten ? "Failed" : "Success",
        source: "Requests",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: me,
        library: libraryTitle(),
        itemName: row.itemName,
        itemUniqueId: row.itemUniqueId,
        summary: `${row.type} request ${decided.status.toLowerCase()} — ${row.itemName}`,
        details: [
          `Requested by ${row.requestedBy} on ${row.requestedAt}`,
          `Reason: ${row.reason}`,
          row.type === "Share" ? `Recipients: ${(row.shareWith ?? []).join(", ")}` : "",
          failure ? `The action failed: ${failure}` : "",
          rowWritten ? "" : "The request row could not be updated, so it may still show as pending.",
        ].filter((d) => d.length > 0),
      }).catch(() => undefined);

      setNotice(
        failure
          ? `Recorded as failed — ${failure}`
          : !rowWritten
            ? "The action was carried out, but the request could not be updated. Refresh before deciding it again."
            : approve
              ? `Approved. ${decisionSummary(row)}`
              : "Rejected. The document is untouched.",
      );
      setDeciding(undefined);
      setNote("");
      await load();
    } catch (e) {
      setNotice(`Could not finish: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  /* ── Render ───────────────────────────────────────────────────────────── */

  if (rows.state === "loading") return <p style={s.wrap}>Loading&hellip;</p>;

  const all = rows.state === "ready" ? rows.value : [];
  const queue = queueFor(all, approverUnits);
  const mine = all.filter((r) => (r.requestedBy ?? "").toLowerCase() === me);
  const tally = counts(all);

  return (
    <section style={s.wrap}>
      <h2 style={s.h2}>Requests</h2>
      <p style={s.sub}>
        Deletion and share requests for documents in the approved library. Uploaders raise them from{" "}
        <strong>My Submissions</strong>; the Head of Unit decides them here. Approving carries the
        action out immediately, using your own permissions.
      </p>

      {notice && <div style={s.ok}>{notice}</div>}

      {rows.state === "error" && (
        <div style={s.err}>
          Could not read the requests ({rows.message}). This is <strong>not</strong> the same as there
          being none — do not treat the queue as empty until it clears.
        </div>
      )}

      {listMissing && (
        <div style={s.warn}>
          <p style={{ margin: "0 0 8px", fontWeight: 600 }}>The requests list does not exist yet</p>
          Nobody can raise a request until it does.
          <div style={s.actions}>
            <button style={busy ? s.off : s.approve} disabled={busy} onClick={() => { provision().catch(() => undefined); }}>
              {busy ? "Creating…" : `Create "${cachedListTitle(LIST_SUFFIX.requests)}"`}
            </button>
          </div>
        </div>
      )}

      {!listMissing && approverUnits.length === 0 && (
        <div style={s.warn}>
          You are not recorded as the approver for any unit, so nothing is waiting on you. If that is
          wrong, an administrator adds an <strong>APR</strong> mapping for your group on the{" "}
          <strong>Folder Access</strong> page.
        </div>
      )}

      {approverUnits.length > 0 && (
        <div style={s.card}>
          <p style={s.head}>Waiting for you ({queue.length})</p>
          {queue.length === 0 ? (
            <p style={s.quiet}>Nothing is waiting for a decision.</p>
          ) : (
            queue.map((r) => (
              <div key={r.id} style={s.row}>
                <div style={s.rowTop}>
                  <span style={s.name}>{r.itemName}</span>
                  <span style={{ ...s.pill, ...PILL[r.status] }}>{r.type}</span>
                </div>
                <div style={s.meta}>
                  {r.requestedBy} · {r.unit} · {(r.requestedAt ?? "").slice(0, 10)}
                </div>
                {r.type === "Share" && (
                  <div style={s.meta}>
                    {(r.shareWith ?? []).join(", ")} · {r.sharePermission ?? "View"}
                    {r.expiresAt ? ` · until ${r.expiresAt.slice(0, 10)}` : " · no expiry"}
                    {/* Named on the ROW, not only in the dialog: this is the fact that decides the
                        answer, and an approver should see it before reaching for a button. */}
                    {(r.shareWith ?? []).some((e) => isExternal(e, tenantDomains)) && (
                      <strong style={{ color: "#8a4b00" }}> · outside the organisation</strong>
                    )}
                  </div>
                )}
                <div style={s.reason}>{r.reason || <em style={s.quiet}>No reason given.</em>}</div>
                <div style={s.actions}>
                  <button style={busy ? s.off : s.approve} disabled={busy}
                    onClick={() => { setNote(""); setDeciding({ row: r, approve: true }); }}>
                    Approve
                  </button>
                  <button style={busy ? s.off : s.reject} disabled={busy}
                    onClick={() => { setNote(""); setDeciding({ row: r, approve: false }); }}>
                    Reject
                  </button>
                </div>
              </div>
            ))
          )}
        </div>
      )}

      <div style={s.card}>
        <p style={s.head}>Your requests ({mine.length})</p>
        {mine.length === 0 ? (
          <p style={s.quiet}>You have not raised any requests.</p>
        ) : (
          mine.map((r) => (
            <div key={r.id} style={s.row}>
              <div style={s.rowTop}>
                <span style={s.name}>{r.itemName}</span>
                <span style={{ ...s.pill, ...PILL[r.status] }}>{r.status}</span>
              </div>
              <div style={s.meta}>
                {r.type} · {(r.requestedAt ?? "").slice(0, 10)}
                {r.decidedBy ? ` · decided by ${r.decidedBy}` : ""}
              </div>
              {r.decisionNote && <div style={s.reason}>{r.decisionNote}</div>}
            </div>
          ))
        )}
      </div>

      <p style={s.quiet}>
        {tally.Pending} pending · {tally.Approved} approved · {tally.Rejected} rejected · {tally.Failed} failed
      </p>

      {deciding && (
        <div style={s.modalBg} onClick={() => { if (!busy) setDeciding(undefined); }}>
          <div style={s.modal} onClick={(e) => e.stopPropagation()}>
            <p style={{ ...s.head, fontSize: 16 }}>
              {deciding.approve ? "Approve" : "Reject"} this request?
            </p>
            {/* Says what approving actually DOES. "Approve" alone does not tell you whether a document
                is about to be recycled or handed to someone outside the company. */}
            <p style={{ fontSize: 13 }}>
              {deciding.approve
                ? decisionSummary(deciding.row)
                : `${deciding.row.itemName} will not be touched, and ${deciding.row.requestedBy} will see your note.`}
            </p>
            {deciding.approve &&
              deciding.row.type === "Share" &&
              (deciding.row.shareWith ?? []).some((e) => isExternal(e, tenantDomains)) && (
                <div style={s.warn}>
                  This sends the document <strong>outside the organisation</strong>. It stays
                  accessible until the share is removed
                  {deciding.row.expiresAt ? ` or ${deciding.row.expiresAt.slice(0, 10)} passes` : ""}.
                </div>
              )}
            <label style={{ display: "block", marginTop: 12, fontSize: 12, fontWeight: 600 }}>
              Note {deciding.approve ? "(optional)" : "— say why; the requester sees this"}
            </label>
            <input style={s.input} value={note} onChange={(e) => setNote(e.target.value)} />
            <div style={s.actions}>
              <button
                style={busy || !canDecide(deciding.row, approverUnits) ? s.off : deciding.approve ? s.approve : s.reject}
                disabled={busy || !canDecide(deciding.row, approverUnits)}
                onClick={() => { decide(deciding.row, deciding.approve, note).catch(() => undefined); }}
              >
                {busy ? "Working…" : deciding.approve ? "Approve and carry it out" : "Reject"}
              </button>
              <button style={s.ghost} disabled={busy} onClick={() => setDeciding(undefined)}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

/** A list item as this screen needs it. Field names are the list's, not the model's. */
function fromListItem(r: Record<string, string>): RequestRow {
  return {
    id: Number(r.Id),
    type: (r.RequestType === "Share" ? "Share" : "Deletion") as RequestRow["type"],
    // Anything unrecognised reads as Pending, so a row written by a newer version is still decidable
    // rather than invisible.
    status: (["Pending", "Approved", "Rejected", "Failed"].indexOf(r.Status) > -1
      ? r.Status
      : "Pending") as RequestStatus,
    itemUniqueId: r.ItemUniqueId ?? "",
    itemName: r.ItemName ?? "",
    itemUrl: r.ItemUrl ?? "",
    segment: r.Segment ?? "",
    unit: r.Unit ?? "",
    unitTermGuid: r.UnitTermGuid ?? "",
    requestedBy: r.RequestedBy ?? "",
    requestedAt: r.RequestedAt ?? "",
    reason: r.Reason ?? "",
    shareWith: (r.ShareWith ?? "").split(/[,;\s]+/).filter((x) => x.length > 0),
    sharePermission: (r.SharePermission === "Edit" ? "Edit" : "View") as RequestRow["sharePermission"],
    expiresAt: r.ExpiresAt ?? "",
    decidedBy: r.DecidedBy ?? "",
    decidedAt: r.DecidedAt ?? "",
    decisionNote: r.DecisionNote ?? "",
  };
}

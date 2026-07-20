import * as React from "react";
import { useState, useMemo } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IRequestShareProps } from "./IRequestShareProps";
import { parseShareTarget, buildRequestPayload, AccessLevel } from "../../../shared/shareGuard";

const LIST = "DMS Share Requests";

const RequestShare: React.FC<IRequestShareProps> = ({ context }) => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  // PeoplePicker wants the narrower IPeoplePickerContext shape (absoluteUrl + the two
  // clients), not the full WebPartContext — WebPartContext nests absoluteUrl under
  // pageContext.web, so build the shape it actually needs instead of forcing a cast.
  // The `unknown` hop below is needed because @pnp/spfx-controls-react ships its own
  // nested copies of @microsoft/sp-http / sp-http-msgraph; the runtime objects are the
  // same SPFx-provided instances, but TS sees them as structurally-incompatible types
  // due to private fields on the duplicated declarations.
  const peoplePickerContext = {
    absoluteUrl: siteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  } as unknown as IPeoplePickerContext;
  // Item context is passed on the query string by the "Request to Share" button.
  const target = useMemo(() => parseShareTarget(new URLSearchParams(window.location.search)), []);
  const [recipientLogin, setRecipientLogin] = useState("");
  const [accessLevel, setAccessLevel] = useState<AccessLevel>("Read");
  const [reason, setReason] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [done, setDone] = useState(false);

  const ensureUser = async (loginName: string): Promise<number> => {
    const res = await context.spHttpClient.post(`${siteUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" },
        body: JSON.stringify({ logonName: loginName }) });
    if (!res.ok) throw new Error(`ensureuser HTTP ${res.status}`);
    return (await res.json()).Id as number;
  };

  const submit = async (): Promise<void> => {
    setErr("");
    if (!target) { setErr("No item context — open this via the Request to Share button."); return; }
    if (!recipientLogin) { setErr("Pick a recipient."); return; }
    setBusy(true);
    try {
      const rid = await ensureUser(recipientLogin);
      const body = buildRequestPayload(target, rid, reason, accessLevel);
      const res = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata" },
          body: JSON.stringify(body) });
      if (!res.ok) throw new Error(`submit HTTP ${res.status} ${(await res.text()).slice(0, 200)}`);
      setDone(true);
    } catch (e) { setErr((e as Error).message); }
    finally { setBusy(false); }
  };

  if (!target) return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
      <p style={{ color: "#a00" }}>No item selected. Open this page via the &ldquo;Request to Share&rdquo; button in a document library.</p>
    </div>);
  if (done) return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
      <p style={{ color: "#0f6c3f" }}>Request submitted for approval. You&rsquo;ll be notified when it&rsquo;s decided.</p>
    </div>);

  const itemName = target.itemUrl.replace(/\/+$/, "").split("/").pop();

  return (
    <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif", maxWidth: 520 }}>
      <h2>Request to share</h2>
      <p style={{ fontSize: 13, color: "#444" }}>
        Item: <strong>{itemName}</strong> ({target.itemType}, {target.library})
      </p>
      <PeoplePicker
        context={peoplePickerContext}
        titleText="Recipient"
        personSelectionLimit={1}
        principalTypes={[PrincipalType.User]}
        resolveDelay={300}
        onChange={(items) => setRecipientLogin((items?.[0] as { loginName?: string })?.loginName ?? "")}
      />
      <label style={{ display: "block", marginTop: 12 }}>Access level{" "}
        <select value={accessLevel} onChange={(e) => setAccessLevel(e.target.value as AccessLevel)} style={{ marginLeft: 8 }}>
          <option value="Read">Read</option>
          <option value="Edit">Edit</option>
        </select>
      </label>
      <label style={{ display: "block", marginTop: 12 }}>Reason
        <textarea value={reason} onChange={(e) => setReason(e.target.value)} rows={3} style={{ display: "block", width: "100%" }} />
      </label>
      {err && <p style={{ color: "#a00" }}>{err}</p>}
      <button disabled={busy} onClick={() => { submit().catch(() => undefined); }} style={{ marginTop: 12 }}>
        {busy ? "Submitting…" : "Submit request"}
      </button>
    </div>
  );
};
export default RequestShare;

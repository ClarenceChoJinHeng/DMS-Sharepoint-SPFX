import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { LIST_SUFFIX, cachedListTitle } from "../../../shared/naming";
import {
  UPLOAD_PAUSE_SETTING,
  uploadsArePaused,
  pauseSettingValue,
} from "../../../shared/uploadPause";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

/**
 * The site-wide upload pause, as a screen in the "Change the folder structure" flow.
 *
 * Spec: docs/superpowers/specs/2026-08-19-upload-pause-design.md
 *
 * Mounted TWICE in that flow — once at the start to pause, once at the end to resume — from ONE
 * component. Two copies would drift, and the resume half is the one nobody would test.
 *
 * ⚠ RESUMING WRITES `no`, NEVER A BLANK CELL. A blank cell cannot be told apart from a row nobody
 * has ever set, and the entire point of the closing step is that an admin can confirm they
 * remembered to turn uploads back on.
 */
export interface IUploadPauseToggleProps {
  context: WebPartContext;
  siteUrl: string;
  /** `pause` opens the flow; `resume` closes it and is worded as the thing still outstanding. */
  mode: "pause" | "resume";
}

type Load = "loading" | "ok" | "error";

export const UploadPauseToggle: React.FC<IUploadPauseToggleProps> = ({ context, siteUrl, mode }) => {
  const [paused, setPaused] = useState<boolean | undefined>(undefined);
  const [load, setLoad] = useState<Load>("loading");
  const [busy, setBusy] = useState<boolean>(false);
  const [note, setNote] = useState<string | undefined>(undefined);

  const listUrl = (): string =>
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')`;

  const filter = `$filter=ConfigType eq 'setting' and Title eq '${UPLOAD_PAUSE_SETTING}'&$top=1`;

  const read = async (): Promise<void> => {
    setLoad("loading");
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${listUrl()}/items?$select=Id,SettingValue&${filter}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) { setLoad("error"); return; }
      const rows = ((await res.json()).value ?? []) as Array<{ Id: number; SettingValue?: string }>;
      // No row at all is a VALID state meaning "not paused" — the setting has simply never been
      // used on this site. Distinct from a failed read, which says nothing and shows an error.
      setPaused(rows.length > 0 ? uploadsArePaused(rows[0].SettingValue) : false);
      setLoad("ok");
    } catch {
      setLoad("error");
    }
  };

  useEffect(() => { read().catch(() => setLoad("error")); }, []);

  const write = async (next: boolean): Promise<void> => {
    setBusy(true);
    setNote(undefined);
    try {
      const find: SPHttpClientResponse = await context.spHttpClient.get(
        `${listUrl()}/items?$select=Id&${filter}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!find.ok) throw new Error(`could not read the config list (HTTP ${find.status})`);
      const rows = ((await find.json()).value ?? []) as Array<{ Id: number }>;
      const existing = rows.length > 0;
      const body = JSON.stringify(
        existing
          ? { SettingValue: pauseSettingValue(next) }
          : { Title: UPLOAD_PAUSE_SETTING, ConfigType: "setting", SettingValue: pauseSettingValue(next) },
      );
      // JSON light: no __metadata and no entity type, and `odata-version` blanked because SPFx
      // injects 4.0, under which SharePoint cannot infer the entity set for a light payload. Same
      // contract as the audit-log writer.
      const headers: Record<string, string> = {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "odata-version": "",
      };
      if (existing) {
        headers["X-HTTP-Method"] = "MERGE";
        headers["IF-MATCH"] = "*";
      }
      const url = existing ? `${listUrl()}/items(${rows[0].Id})` : `${listUrl()}/items`;
      const res: SPHttpClientResponse = await context.spHttpClient.post(
        url, SPHttpClient.configurations.v1, { headers, body },
      );
      if (!res.ok) throw new Error(`HTTP ${res.status} — ${(await res.text()).slice(0, 200)}`);
      setPaused(next);
      setNote(next
        ? "Uploads are paused. Nobody can file a document until you turn this back on."
        : "Uploads are on again.");
    } catch (e) {
      // NAMED, never a bare failure. Pausing is what makes the migration safe, so an admin who
      // believes uploads stopped when they did not will migrate on top of live traffic.
      setNote(`Could not change the setting — ${(e as Error).message}. Uploads are UNCHANGED.`);
    } finally {
      setBusy(false);
    }
  };

  const box: React.CSSProperties = {
    padding: "12px 14px", borderRadius: 4, fontSize: 13, lineHeight: 1.5, marginBottom: 12,
  };
  const btn = (enabled: boolean, on: string, off: string): React.CSSProperties => ({
    padding: "7px 16px", fontSize: 13, borderRadius: 4, border: "none", color: "#fff",
    background: enabled ? on : off,
    cursor: enabled ? "pointer" : "not-allowed",
  });

  const canPause = load === "ok" && !busy && paused === false;
  const canResume = load === "ok" && !busy && paused === true;

  return (
    <div>
      <h3 style={{ margin: "0 0 4px", fontSize: 18 }}>
        {mode === "pause" ? "Pause uploads" : "Turn uploads back on"}
      </h3>
      <p style={{ margin: "0 0 14px", color: "#666", lineHeight: 1.5 }}>
        {mode === "pause"
          ? "While this is on, the Upload Form and Bulk Upload refuse new documents across the whole site. Do the structure change, then turn it back on at the last step."
          : "The structure change is finished. Until you turn this off, nobody can file a document."}
      </p>

      {/* A read that FAILED is not "uploads are on", and the difference matters more here than
          almost anywhere else: an admin who believes uploads are paused when the setting could not
          even be read will migrate while people are still filing documents. */}
      {load === "error" ? (
        <div style={{ ...box, background: "#fdeeee", border: "1px solid #e0a0a0", color: "#7a2020" }}>
          <strong>Could not read the current setting.</strong> Uploads may or may not be paused —
          this screen cannot tell. Check the <code>{UPLOAD_PAUSE_SETTING}</code> row on CRS Config
          before you continue.
        </div>
      ) : load === "loading" ? (
        <div style={{ ...box, background: "#f4f4f4", border: "1px solid #ddd", color: "#555" }}>
          Reading the current setting…
        </div>
      ) : paused ? (
        <div style={{ ...box, ...NOTICE_ATTENTION }}>
          <strong>Uploads are PAUSED right now.</strong> The Upload Form and Bulk Upload are refusing
          new documents site-wide.
        </div>
      ) : (
        <div style={{ ...box, background: "#eef7f1", border: "1px solid #cfe4d8", color: "#0f6c3f" }}>
          <strong>Uploads are ON.</strong> People can file documents normally.
        </div>
      )}

      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <button
          disabled={!canPause}
          onClick={() => { write(true).catch(() => undefined); }}
          style={btn(canPause, "#9a6b00", "#c9b98f")}
        >
          Pause uploads
        </button>
        <button
          disabled={!canResume}
          onClick={() => { write(false).catch(() => undefined); }}
          style={btn(canResume, "#0f6c3f", "#b6c6bd")}
        >
          Turn uploads back on
        </button>
        <button
          disabled={busy}
          onClick={() => { read().catch(() => setLoad("error")); }}
          style={{ padding: "7px 14px", fontSize: 13, borderRadius: 4, border: "1px solid #ccc", background: "#fff", cursor: "pointer" }}
        >
          Refresh
        </button>
      </div>

      {note ? (
        <p style={{ marginTop: 12, fontSize: 13, color: note.indexOf("Could not") === 0 ? "#7a2020" : "#0f6c3f" }}>
          {note}
        </p>
      ) : undefined}

      <p style={{ marginTop: 16, fontSize: 12, color: "#777", lineHeight: 1.6 }}>
        <strong>This does not reach a page somebody already had open.</strong> A browser tab opened
        before you paused still holds its own copy of the form. It is re-checked at the moment of
        upload, so the document is refused rather than misfiled — but the safest time to change the
        structure is still outside working hours.
      </p>
    </div>
  );
};

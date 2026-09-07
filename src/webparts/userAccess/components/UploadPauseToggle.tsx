import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { LIST_SUFFIX, cachedListTitle } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import {
  UPLOAD_PAUSE_SETTING,
  uploadsArePaused,
  pauseSettingValue,
} from "../../../shared/uploadPause";
// `NOTICE_ATTENTION` was the paused banner's palette. The status switch states OFF in its own red,
// so the shared attention style is no longer used here — re-import it if a banner ever returns.

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
  /**
   * The live pause state, reported up whenever this screen learns it.
   *
   * ⚠ WITHOUT THIS THE GUIDED FLOW CONTRADICTS ITS OWN PANEL. `FolderAdmin` reads the setting once,
   * in its facts effect, and since 1.0.452.0 `blocksNext` holds Next while it reads `false`. Flip the
   * toggle and nothing told it — so the gate went on saying "Uploads are still switched on" beneath a
   * panel that had just paused them, and only a full page reload cleared it (reported on site
   * 2026-09-07). The toggle's own Refresh button does not help: it re-reads this panel and not the
   * fact the gate consults.
   *
   * Called after a successful READ as well as a successful write, because the mount-time read is
   * itself newer than the facts effect's when this screen is re-entered.
   *
   * Never called on a failure. A failed read says nothing about the setting, and reporting a guess
   * would put the gate back to claiming uploads are on when nobody knows.
   *
   * Optional so the standalone mount, and any future one, needs nothing.
   */
  onChanged?: (paused: boolean) => void;
}

type Load = "loading" | "ok" | "error";

export const UploadPauseToggle: React.FC<IUploadPauseToggleProps> = ({ context, siteUrl, mode, onChanged }) => {
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
      /* ⚠ AWAITED HERE, INSIDE THE READER — NOT LEFT TO A MOUNT EFFECT SOMEWHERE ELSE.
         `listUrl()` calls `cachedListTitle`, which answers the LEGACY `DMS Config` until priming
         settles. On a CRS site that 404s, so this screen showed "Could not read the current setting"
         intermittently — a pure race, which is why clearing the cache changed whether it appeared
         (reported on site 2026-09-07). Same defect as the reconciliation segment picker in
         1.0.207.0, and the same rule: every screen that builds a list URL must prime in its own
         reader. Cheap after the first call — `primeNames` caches. */
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${listUrl()}/items?$select=Id,SettingValue&${filter}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) { setLoad("error"); return; }
      const rows = ((await res.json()).value ?? []) as Array<{ Id: number; SettingValue?: string }>;
      // No row at all is a VALID state meaning "not paused" — the setting has simply never been
      // used on this site. Distinct from a failed read, which says nothing and shows an error.
      const now = rows.length > 0 ? uploadsArePaused(rows[0].SettingValue) : false;
      setPaused(now);
      setLoad("ok");
      // Only on a read that actually answered — see `onChanged`. The two early exits above report
      // nothing, so the flow's fact stays unknown and the gate stays open.
      if (onChanged) onChanged(now);
    } catch {
      setLoad("error");
    }
  };

  useEffect(() => { read().catch(() => setLoad("error")); }, []);

  const write = async (next: boolean): Promise<void> => {
    setBusy(true);
    setNote(undefined);
    try {
      // Primed here too, for the same reason as `read` — and it matters more: a write against the
      // legacy title fails, and the admin is told "Uploads are UNCHANGED" for what is a naming race.
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
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
      // After the write SUCCEEDED, never before it. The catch below reports nothing, so a failed
      // write leaves the flow's fact as it was rather than claiming a change that did not land.
      if (onChanged) onChanged(next);
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
  // The two coloured action buttons this styled were replaced by the SharePoint status switch
  // (2026-09-06), which carries its own colours.

  const canPause = load === "ok" && !busy && paused === false;
  const canResume = load === "ok" && !busy && paused === true;

  return (
    <div>
      {/* ⚠ NO HEADING AND NO DESCRIPTION HERE. This component is mounted ONLY as a step of a guided
          flow (`FolderAdmin` is its single importer), and that step already prints the same title
          and the same sentence immediately above — the client's screenshot showed each of them
          twice, one under the other.
          The wording now lives in ONE place: the `pauseUploads` / `resumeUploads` steps in
          `folderFlows.ts`. If this is ever mounted somewhere with no step around it, it needs its
          heading back, or it opens as an unlabelled switch. */}

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
      ) : (
        /* ── SharePoint status (client's design, 2026-09-06) ──────────────────────────────────
           ONE control instead of two buttons. The pair could express a state the site was not in -
           "Pause uploads" was clickable-looking while already paused - and the toggle cannot.

           ⚠ IT SHOWS THE STATE AND SETS THE OPPOSITE, which is the ordinary meaning of a switch but
           worth stating: ON means uploads are working, and pressing it PAUSES them. `canPause` /
           `canResume` still gate it, so a failed read (handled above) never reaches here and the
           switch is never live over a state nobody could confirm.

           ⚠ `paused === false` RATHER THAN `!paused`. The value is three-state - true, false, and
           not-yet-read - and `!undefined` is true, which would draw a green "uploads are on" switch
           over a setting that had not been read at all. */
        <div style={{ ...box, background: "#fff", border: "1px solid #e1e1e1", color: "#323130" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16 }}>
            <strong style={{ fontSize: 14 }}>SharePoint status</strong>
            <button
              type="button"
              role="switch"
              aria-checked={paused === false}
              aria-label={paused === false ? "Uploads are on — turn them off" : "Uploads are off — turn them on"}
              disabled={paused === false ? !canPause : !canResume}
              onClick={() => { write(paused === false).catch(() => undefined); }}
              style={{
                position: "relative",
                width: 46,
                height: 24,
                flexShrink: 0,
                borderRadius: 12,
                border: "none",
                padding: 0,
                cursor: busy ? "not-allowed" : "pointer",
                background: paused === false ? "#0f6c3f" : "#d13438",
                opacity: busy ? 0.6 : 1,
              }}
            >
              <span
                aria-hidden="true"
                style={{
                  position: "absolute",
                  top: 3,
                  left: paused === false ? 25 : 3,
                  width: 18,
                  height: 18,
                  borderRadius: "50%",
                  background: "#fff",
                  transition: "left .15s ease",
                }}
              />
            </button>
          </div>
          <div style={{ marginTop: 8 }}>
            <strong style={{ color: paused === false ? "#0f6c3f" : "#a4262c" }}>
              {paused === false ? "ON" : "OFF"}
            </strong>
            <div style={{ marginTop: 2, lineHeight: 1.5 }}>
              {paused === false
                ? "Users can upload documents normally."
                : "Uploads are paused. The Upload Form and Bulk Upload refuse new documents until this is turned back on."}
            </div>
          </div>
          <button
            disabled={busy}
            onClick={() => { read().catch(() => setLoad("error")); }}
            style={{ marginTop: 12, padding: "7px 14px", fontSize: 13, borderRadius: 4, border: "1px solid #ccc", background: "#fff", cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6 }}
          >
            <span aria-hidden="true">&#8635;</span>
            Refresh
          </button>
        </div>
      )}

      {note ? (
        <p style={{ marginTop: 12, fontSize: 13, color: note.indexOf("Could not") === 0 ? "#7a2020" : "#0f6c3f" }}>
          {note}
        </p>
      ) : undefined}

      {/* ⚠ THE "already-open page" NOTE WAS REMOVED AT THE CLIENT'S REQUEST (2026-09-06), AND WHAT
          IT DESCRIBED IS STILL TRUE. A browser tab opened before the pause holds its own copy of the
          form: the pause is re-checked at the moment of upload, so such a document is REFUSED rather
          than misfiled — the guard works, it simply cannot reach a page that never asks the server
          again. The operational answer is unchanged and is the client's own: make structure changes
          outside working hours. If anyone asks why somebody could still see an upload form after the
          pause, this is the answer. */}
    </div>
  );
};

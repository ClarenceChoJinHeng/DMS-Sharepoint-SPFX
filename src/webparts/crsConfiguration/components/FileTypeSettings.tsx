import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AuditOutcome, EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import {
  FALLBACK_FILE_TYPES,
  normalizeFileTypes,
  readAllowedFileTypesField,
} from "../../../shared/allowedFileTypes";
import {
  AdditionVerdict,
  badgeFor,
  blockedButTicked,
  classifyAddition,
  describeExtension,
  isBlockedType,
  panelState,
} from "../../../shared/fileTypeSettings";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

/**
 * File type settings — which extensions may be uploaded.
 *
 * Spec: docs/superpowers/specs/2026-08-13-file-type-settings-page-design.md
 *
 * Edits the `AllowedFileTypes` multi-select Choice column on `CRS Config`, the single source of truth
 * since 2026-07-30 — until now editable only by ticking boxes in a list view.
 *
 * TWO WRITES, TWO PERMISSIONS, and they look identical on screen: a toggle updates the ITEM's value
 * (item edit), while adding a type updates the FIELD's Choices (Manage Lists). They fail differently and
 * say so differently.
 *
 * This page is NOT a security boundary — the list filters a file picker and validates in JS. Whoever can
 * edit `CRS Config` controls upload policy whether or not they can see this page. Spec §3.
 */

const CONFIG_ROW_TITLE = "allowedExtensions";
const COLUMN_INTERNAL_NAME = "AllowedFileTypes";

const s: Record<string, React.CSSProperties> = {
  // Capped and CENTRED, like every other admin screen (AccessShell, FolderManager). The cap keeps the
  // Description column readable on a wide monitor; without `margin auto` it pinned hard left inside a
  // full-width section, which reads as a rendering fault rather than a deliberate measure.
  /* The Upload Form's page shell - see the note in AuditLog.tsx. Centred but unpadded before. */
  wrap: { fontFamily: "'Segoe UI', sans-serif", color: "#1b1b1b", maxWidth: 1100, margin: "32px auto", padding: "0 24px 48px" },
  h2: { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle: { fontSize: 13, color: "#605e5c", margin: "0 0 20px", lineHeight: 1.5 },
  card: { border: "1px solid #e1dfdd", borderRadius: 8, padding: "16px 18px", marginBottom: 16 },
  cardTitle: {
    fontSize: 12, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase",
    color: "#0f6c3f", margin: "0 0 4px",
  },
  hint: { fontSize: 12, color: "#605e5c", margin: "0 0 14px", lineHeight: 1.5 },
  msg: { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.55 },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { ...NOTICE_ATTENTION },
  ok: { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  info: { background: "#f3f2f1", border: "1px solid #e1dfdd", color: "#323130" },
  head: {
    display: "grid", gridTemplateColumns: "84px minmax(0,1fr) 92px 84px", columnGap: 12,
    padding: "0 10px 8px", fontSize: 12, fontWeight: 600, color: "#605e5c",
    borderBottom: "1px solid #edebe9",
  },
  row: {
    display: "grid", gridTemplateColumns: "84px minmax(0,1fr) 92px 84px", columnGap: 12, rowGap: 4,
    alignItems: "center", padding: "10px", borderBottom: "1px solid #f3f2f1",
  },
  /** Quiet by default — removing a type is rarer than toggling one, so it must not compete. */
  removeLink: {
    background: "none", border: "none", padding: 0, fontSize: 12, color: "#a4262c",
    cursor: "pointer", textDecoration: "underline",
  },
  removeOff: {
    background: "none", border: "none", padding: 0, fontSize: 12, color: "#c8c6c4",
    cursor: "not-allowed",
  },
  danger: {
    background: "#a4262c", color: "#fff", border: "none", borderRadius: 4, padding: "8px 16px",
    fontSize: 13, cursor: "pointer",
  },
  badge: {
    display: "inline-block", fontSize: 10, fontWeight: 700, letterSpacing: ".04em",
    background: "#f3f2f1", border: "1px solid #e1dfdd", borderRadius: 4, padding: "3px 6px",
    color: "#605e5c", minWidth: 34, textAlign: "center",
  },
  ext: { fontSize: 13, fontWeight: 600, fontFamily: "Consolas, monospace" },
  desc: { fontSize: 13, color: "#323130" },
  locked: { fontSize: 11, color: "#a4262c", fontWeight: 600 },
  input: {
    padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4, width: 180,
    boxSizing: "border-box",
  },
  btn: {
    background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "8px 16px",
    fontSize: 13, cursor: "pointer",
  },
  ghost: {
    background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4,
    padding: "7px 14px", fontSize: 13, cursor: "pointer",
  },
  off: {
    background: "#f3f2f1", color: "#a19f9d", border: "1px solid #e1dfdd", borderRadius: 4,
    padding: "8px 16px", fontSize: 13, cursor: "not-allowed",
  },
  overlay: {
    position: "fixed", top: 0, right: 0, bottom: 0, left: 0, background: "rgba(0,0,0,.35)",
    display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16,
  },
  modal: {
    background: "#fff", borderRadius: 8, padding: "20px 22px", maxWidth: 460, width: "100%",
    boxShadow: "0 8px 30px rgba(0,0,0,.2)",
  },
};

/** A switch built from a button — no control library, and it announces itself to a screen reader. */
const Toggle = (props: {
  on: boolean;
  disabled?: boolean;
  busy?: boolean;
  label: string;
  onClick: () => void;
}): React.ReactElement => (
  <button
    role="switch"
    aria-checked={props.on}
    aria-label={props.label}
    disabled={props.disabled === true || props.busy === true}
    onClick={props.onClick}
    style={{
      width: 46, height: 24, borderRadius: 12, position: "relative", padding: 0,
      cursor: props.disabled === true ? "not-allowed" : "pointer",
      border: `1px solid ${props.on ? "#0f6c3f" : "#c8c8c8"}`,
      background: props.disabled === true ? "#f3f2f1" : props.on ? "#0f6c3f" : "#fff",
      opacity: props.busy === true ? 0.6 : 1,
    }}
  >
    <span
      style={{
        position: "absolute", top: 2, left: props.on ? 24 : 2, width: 18, height: 18,
        borderRadius: "50%", background: props.on ? "#fff" : "#a19f9d",
      }}
    />
  </button>
);

export interface FileTypeSettingsProps {
  context: WebPartContext;
  siteUrl: string;
}

interface Note {
  tone: "ok" | "warn" | "err" | "info";
  text: string;
}

export default function FileTypeSettings({
  context,
  siteUrl,
}: FileTypeSettingsProps): React.ReactElement {
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [columnPresent, setColumnPresent] = useState(false);
  const [fieldId, setFieldId] = useState<string>("");
  const [choices, setChoices] = useState<string[]>([]);
  const [ticked, setTicked] = useState<string[]>([]);
  const [itemId, setItemId] = useState<number | undefined>(undefined);
  const [busyExt, setBusyExt] = useState<string | undefined>(undefined);
  const [note, setNote] = useState<Note | undefined>(undefined);
  const [draft, setDraft] = useState("");
  const [adding, setAdding] = useState(false);
  const [confirmLast, setConfirmLast] = useState<string | undefined>(undefined);
  const [suggest, setSuggest] = useState<{ typed: string; better: string } | undefined>(undefined);
  const [highlight, setHighlight] = useState<string | undefined>(undefined);
  /** Set when a change was applied but not recorded. Never blocks — see logPolicy. */
  const [auditWarn, setAuditWarn] = useState(false);

  const configList = (): string => encodeURIComponent(cachedListTitle(LIST_SUFFIX.config));

  /**
   * Record a policy change in the audit log.
   *
   * Always AFTER the change has been made, and never able to fail it: the write itself never throws,
   * and a false return only raises a non-blocking banner. Refusing a legitimate policy change because
   * the log was unreachable would be worse than an incomplete log — but an admin is still told,
   * because an audit gap somebody knows about is worth far more than one nobody does.
   */
  /**
   * The type awaiting a remove confirmation, or undefined.
   *
   * Removing is a TWO-WRITE operation like adding, and the order is the whole safety — see
   * `commitRemove`. Kept as the extension string rather than a boolean so the dialog can name it.
   */
  const [removingExt, setRemovingExt] = useState<string | undefined>(undefined);
  const [removeBusy, setRemoveBusy] = useState(false);

  const logPolicy = async (
    summary: string,
    details: string[],
    outcome: AuditOutcome = "Success",
  ): Promise<void> => {
    const ok = await writeAudit(context.spHttpClient, siteUrl, {
      event: EVENT.policyChanged,
      outcome,
      source: "CrsConfiguration",
      at: new Date(),
      actorName: context.pageContext.user.displayName,
      actorEmail: context.pageContext.user.email,
      summary,
      details,
    });
    if (!ok) setAuditWarn(true);
  };

  /* ── Load ──────────────────────────────────────────────────────────────────── */

  const load = async (): Promise<void> => {
    // NAMES FIRST. Every read here goes through cachedListTitle, and an unprimed cache resolves to the
    // legacy `DMS …` titles, which 404 on a renamed site and present as a MISSING list rather than a
    // list called something else. A priming FAILURE must not stop the read: the legacy title is correct
    // on a site that was never renamed.
    await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);

    // Matched on INTERNAL name, never display name. A matching display name over a different internal
    // name fails exactly like an absent column and looks correct.
    const fieldRes: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${configList()}')/fields` +
        `?$select=Id,InternalName,Title,Choices,FillInChoice` +
        `&$filter=InternalName eq '${COLUMN_INTERNAL_NAME}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!fieldRes.ok) {
      throw new Error(
        `${cachedListTitle(LIST_SUFFIX.config)} could not be read (HTTP ${fieldRes.status}).`,
      );
    }
    const fields = ((await fieldRes.json()).value ?? []) as Array<{
      Id?: string;
      Choices?: string[];
    }>;
    const field = fields[0];
    const present = field !== undefined;
    setColumnPresent(present);
    setFieldId(field?.Id ?? "");
    setChoices(present ? normalizeFileTypes(field?.Choices ?? []) : []);

    // Only $select the column when it EXISTS — selecting an absent column returns HTTP 400 for the whole
    // request, which would read as "the list is broken" rather than "the column is missing".
    const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${configList()}')/items` +
        `?$select=Id${present ? `,${COLUMN_INTERNAL_NAME}` : ""}` +
        `&$filter=Title eq '${CONFIG_ROW_TITLE}'&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!itemRes.ok) {
      throw new Error(`The ${CONFIG_ROW_TITLE} row could not be read (HTTP ${itemRes.status}).`);
    }
    const rows = ((await itemRes.json()).value ?? []) as Array<{ Id?: number }>;
    const row = rows[0];
    setItemId(row?.Id);
    // readAllowedFileTypesField distinguishes "emptied" (null -> []) from "absent" (undefined) — the
    // difference between a deliberate block and an unprovisioned column.
    const raw = row ? readAllowedFileTypesField(row as { AllowedFileTypes?: unknown }) : undefined;
    setTicked(normalizeFileTypes(raw ?? []));
    if (present && !row) {
      setNote({
        tone: "err",
        text:
          `${cachedListTitle(LIST_SUFFIX.config)} has no "${CONFIG_ROW_TITLE}" row, so there is nowhere ` +
          `to store the selection. Add an item with that exact Title before using this page.`,
      });
    }
    setLoading(false);
  };

  useEffect(() => {
    load().catch((err: Error) => {
      setLoadError(err.message);
      setLoading(false);
    });
  }, []);

  /* ── Writes ────────────────────────────────────────────────────────────────── */

  const digest = async (): Promise<string> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    return (await res.json()).FormDigestValue as string;
  };

  /** MERGE the ITEM's value. Needs item edit only. */
  const writeTicked = async (next: string[]): Promise<void> => {
    if (itemId === undefined) throw new Error(`there is no "${CONFIG_ROW_TITLE}" row to update`);
    const token = await digest();
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${configList()}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": token,
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        body: JSON.stringify({ AllowedFileTypes: next }),
      },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status}. ${body.slice(0, 140)}`);
    }
  };

  /**
   * MERGE the FIELD's Choices. Needs MANAGE LISTS, not item edit — a different permission from every
   * other write here, which is why its failure says something different.
   *
   * Field updates require the type in `__metadata`, which exists only in the verbose payload, so this
   * one call goes odata=verbose while everything else stays nometadata.
   */
  const writeChoices = async (next: string[]): Promise<void> => {
    const token = await digest();
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${configList()}')/fields(guid'${fieldId}')`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": token,
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
          // Load-bearing (CLAUDE.md gotcha, same as spAuditLog.ts's WRITE_HEADERS): SPFx's
          // SPHttpClient injects `odata-version: 4.0` on every request regardless of what the Accept/
          // Content-Type headers say, and OData 4 does not agree with the OData 3 verbose dialect this
          // write needs (a Field's `Choices` MERGE requires the __metadata envelope, which only verbose
          // supports). Without this override SharePoint answers 400 "Parsing JSON Light feeds or entries
          // in requests without entity set is not supported" — found live 2026-08-26 adding a file type.
          "odata-version": "",
        },
        body: JSON.stringify({
          __metadata: { type: "SP.FieldMultiChoice" },
          Choices: { results: next },
        }),
      },
    );
    if (!res.ok) {
      if (res.status === 403) {
        throw new Error(
          "adding a file type changes the column itself, which needs site owner rights (Manage Lists). " +
            "Switching existing types on and off does not",
        );
      }
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status}. ${body.slice(0, 140)}`);
    }
  };

  /* ── Toggle ────────────────────────────────────────────────────────────────── */

  const applyToggle = async (ext: string, on: boolean): Promise<void> => {
    const before = ticked;
    const next = on ? ticked.concat([ext]) : ticked.filter((t) => t !== ext);
    setBusyExt(ext);
    setNote(undefined);
    // Optimistic, and REVERTED on failure: a switch that stays on after a failed save is a UI lying about
    // a permission setting, and the lie survives until someone reloads.
    setTicked(next);
    try {
      await writeTicked(next);
      setNote({
        tone: "ok",
        text: on
          ? `${ext} is now allowed. Anyone with the upload form open will need to refresh.`
          : `${ext} is now blocked for new uploads. Files already uploaded are untouched.`,
      });
      // The resulting SET is recorded, not just the delta: a reader months later needs to know what
      // the policy became, and reconstructing it from a chain of deltas is exactly the work an audit
      // log should have already done.
      await logPolicy(
        on ? `File type ${ext} allowed` : `File type ${ext} blocked`,
        [
          on ? `Allowed ${ext}.` : `Blocked ${ext} for new uploads.`,
          `Allowed after this change: ${next.slice().sort().join(", ") || "(none — all uploads blocked)"}`,
        ],
      );
    } catch (err) {
      setTicked(before);
      setNote({ tone: "err", text: `Could not save ${ext} — ${(err as Error).message}` });
    } finally {
      setBusyExt(undefined);
    }
  };

  const onToggle = (ext: string, currentlyOn: boolean): void => {
    if (isBlockedType(ext)) return;
    // Turning off the LAST enabled type blocks every upload in every unit. Allowed (client, 2026-08-13)
    // but never silently: the confirm names all three consequences.
    if (currentlyOn && ticked.length === 1) {
      setConfirmLast(ext);
      return;
    }
    applyToggle(ext, !currentlyOn).catch(() => undefined);
  };

  /* ── Add ───────────────────────────────────────────────────────────────────── */

  const commitAdd = async (ext: string): Promise<void> => {
    setAdding(true);
    setNote(undefined);
    setSuggest(undefined);
    try {
      const nextChoices = choices.concat([ext]);
      await writeChoices(nextChoices);
      setChoices(nextChoices);
      // Added in order to be allowed, so it arrives enabled. If this second write fails the choice still
      // exists — reported, not hidden, because the row would otherwise render unticked and look deliberate.
      const nextTicked = ticked.concat([ext]);
      try {
        await writeTicked(nextTicked);
        setTicked(nextTicked);
        setNote({
          tone: "ok",
          text: `${ext} added and allowed. Anyone with the upload form open will need to refresh.`,
        });
        await logPolicy(`File type ${ext} added and allowed`, [
          `Added ${ext} to the list of choices.`,
          `Allowed after this change: ${nextTicked.slice().sort().join(", ")}`,
        ]);
      } catch (err) {
        setNote({
          tone: "warn",
          text:
            `${ext} was added to the list but could not be enabled — ${(err as Error).message}. ` +
            `Switch it on below.`,
        });
      }
      setDraft("");
      setHighlight(ext);
    } catch (err) {
      setNote({ tone: "err", text: `Could not add ${ext} — ${(err as Error).message}` });
    } finally {
      setAdding(false);
    }
  };

  /* ── Remove ────────────────────────────────────────────────────────────────── */

  /**
   * Remove a file type from the column's Choices entirely.
   *
   * ⚠ THE ORDER IS THE WHOLE SAFETY: UNTICK FIRST, THEN REMOVE THE CHOICE. The list on screen is the
   * column's `Choices`; what actually permits an upload is the ITEM's stored value. Drop a choice that
   * is still ticked and SharePoint can leave the stale value behind in the item — so the type would go
   * on being ALLOWED with no row on screen showing it, on the one page whose entire job is stating
   * upload policy. Doing it this way, the worst case is the reverse and harmless: the untick lands, the
   * Choices write fails, and the type is left blocked-but-listed — visible, and removable again.
   *
   * That is also why this is one action rather than making the admin untick first: a two-step the admin
   * can get half-way through is how the dangerous ordering happens by hand.
   */
  const commitRemove = async (ext: string): Promise<void> => {
    setRemoveBusy(true);
    setNote(undefined);
    const wasOn = ticked.indexOf(ext) !== -1;
    try {
      // Step 1 — untick, only if it is on. An already-off type needs no item write at all.
      if (wasOn) {
        const nextTicked = ticked.filter((t) => t !== ext);
        await writeTicked(nextTicked);
        setTicked(nextTicked);
      }
      // Step 2 — drop it from the column.
      const nextChoices = choices.filter((c) => c !== ext);
      await writeChoices(nextChoices);
      setChoices(nextChoices);
      setNote({
        tone: "ok",
        text:
          `${ext} removed. It is no longer offered on this page` +
          (wasOn ? " and is blocked for new uploads." : ".") +
          " Files already uploaded are untouched, and you can add it again at any time.",
      });
      await logPolicy(`File type ${ext} removed`, [
        `Removed ${ext} from the list of choices.`,
        wasOn
          ? `It was ALLOWED at the time, so this also blocked it for new uploads.`
          : `It was already blocked, so what may be uploaded did not change.`,
        `Allowed after this change: ${ticked.filter((t) => t !== ext).slice().sort().join(", ") || "(none — all uploads blocked)"}`,
      ]);
      setRemovingExt(undefined);
    } catch (err) {
      // Named per half, because the two failures need opposite follow-ups: a failed untick leaves the
      // type fully in place, while a failed Choices write leaves it blocked but still listed.
      setNote({
        tone: "err",
        text:
          `Could not remove ${ext} — ${(err as Error).message}. ` +
          (wasOn
            ? "Check whether it is still switched on below before trying again."
            : "It is still switched off, so nothing can be uploaded with it either way."),
      });
      setRemovingExt(undefined);
    } finally {
      setRemoveBusy(false);
    }
  };

  const onAdd = (): void => {
    const verdict: AdditionVerdict = classifyAddition(draft, choices, ticked);
    setSuggest(undefined);
    if (verdict.kind === "invalid") {
      setNote({ tone: "err", text: verdict.reason });
      return;
    }
    if (verdict.kind === "blocked") {
      setNote({
        tone: "err",
        text:
          `${verdict.ext} cannot be allowed — it is a program or script, not a document. Allowing it ` +
          `would turn the repository into a place to share executables.`,
      });
      // Recorded even though nothing changed. An attempt to allow executables is exactly what an
      // audit log should surface, and it is one of the two events Purview cannot see either —
      // nothing ever reached SharePoint.
      logPolicy(
        `Refused to allow ${verdict.ext}`,
        [
          `Someone tried to allow ${verdict.ext}, a program or script type.`,
          "The policy was not changed.",
        ],
        "Refused",
      ).catch(() => undefined);
      return;
    }
    if (verdict.kind === "exists-enabled") {
      setNote({ tone: "info", text: `${verdict.ext} is already in the list and already allowed.` });
      setHighlight(verdict.ext);
      setDraft("");
      return;
    }
    if (verdict.kind === "exists-disabled") {
      // Reported, NOT auto-enabled (client, 2026-08-13). The row is highlighted so it need not be hunted
      // for, and the decision stays with the admin.
      setNote({
        tone: "warn",
        text: `${verdict.ext} is already in the list but switched off. Switch it on below to allow it.`,
      });
      setHighlight(verdict.ext);
      setDraft("");
      return;
    }
    if (verdict.suggestion) {
      setSuggest({ typed: verdict.ext, better: verdict.suggestion });
      return;
    }
    commitAdd(verdict.ext).catch(() => undefined);
  };

  /** Used by the did-you-mean dialog, where the better spelling may already be in the list. */
  const addOrPointAt = (ext: string): void => {
    const verdict = classifyAddition(ext, choices, ticked);
    if (verdict.kind === "new") {
      commitAdd(verdict.ext).catch(() => undefined);
      return;
    }
    if (verdict.kind === "exists-disabled") {
      setNote({
        tone: "warn",
        text: `${ext} is already in the list but switched off. Switch it on below to allow it.`,
      });
    } else {
      setNote({ tone: "info", text: `${ext} is already allowed.` });
    }
    setHighlight(ext);
    setDraft("");
  };

  /* ── Render ────────────────────────────────────────────────────────────────── */

  const state = panelState(columnPresent, ticked);
  const stray = blockedButTicked(ticked);
  const rows = choices.slice().sort((a, b) => a.localeCompare(b));
  const listName = cachedListTitle(LIST_SUFFIX.config);

  const toneStyle = (tone: Note["tone"]): React.CSSProperties =>
    tone === "err" ? s.err : tone === "warn" ? s.warn : tone === "ok" ? s.ok : s.info;

  if (loading) {
    return (
      <section style={s.wrap}>
        <h2 style={s.h2}>File Type Management</h2>
        <p style={{ fontSize: 13, color: "#605e5c" }}>Loading file type settings&hellip;</p>
      </section>
    );
  }

  return (
    <section style={s.wrap}>
      <h2 style={s.h2}>File Type Management</h2>
      <p style={s.subtitle}>Control which file types can be uploaded to the repository.</p>

      {loadError && <div style={{ ...s.msg, ...s.err }}>{loadError}</div>}
      {note && <div style={{ ...s.msg, ...toneStyle(note.tone) }}>{note.text}</div>}

      {/* An audit gap an admin KNOWS about is worth far more than one nobody does — so this is said
          plainly, while never having blocked the change itself. */}
      {auditWarn && (
        <div style={{ ...s.msg, ...toneStyle("warn") }}>
          Your change was applied, but could not be recorded in the audit log. The audit log may not be
          set up yet, or may not be writable by you — worth checking, because the change itself went
          through.
        </div>
      )}

      {/* A blocked type that is somehow ALLOWED cannot be produced on this page, so it arrived from a
          hand edit or a migration. Said loudly, with the fix one click away. */}
      {stray.length > 0 && (
        <div style={{ ...s.msg, ...s.err }}>
          <strong>{stray.join(", ")}</strong> {stray.length === 1 ? "is" : "are"} currently allowed and
          should not be — {stray.length === 1 ? "it is a program or script" : "they are programs or scripts"}.
          Switch {stray.length === 1 ? "it" : "them"} off below.
        </div>
      )}

      {state === "no-column" ? (
        <>
          <div style={{ ...s.msg, ...s.warn }}>
            This site has no <strong>{COLUMN_INTERNAL_NAME}</strong> column, so uploads are running on the
            built-in list below and the client&apos;s own policy is <strong>not in force</strong>. To
            manage file types here, add a column to <strong>{listName}</strong>: type{" "}
            <strong>Choice</strong>, name <strong>{COLUMN_INTERNAL_NAME}</strong> (no space), allow{" "}
            <strong>multiple selections</strong>, and leave{" "}
            <strong>&quot;Allow fill-in choices&quot; off</strong>.
          </div>
          <div style={s.card}>
            <p style={s.cardTitle}>Built-in file types — currently in force</p>
            <div style={s.hint}>
              Read-only. These come from the code and apply only while the column is missing.
            </div>
            {/* THREE columns here, not the shared four: this panel is read-only (the column is
                missing, so these come from the code) and can never carry a Remove. Inheriting the
                four-column grid would reserve an empty 84px gutter on every row. */}
            <div style={{ ...s.head, gridTemplateColumns: "84px minmax(0,1fr) 92px" }}>
              <span>Extension</span>
              <span>Description</span>
              <span />
            </div>
            {normalizeFileTypes(FALLBACK_FILE_TYPES).map((ext) => (
              <div key={ext} style={{ ...s.row, gridTemplateColumns: "84px minmax(0,1fr) 92px" }}>
                <span>
                  <span style={s.badge}>{badgeFor(ext)}</span>
                </span>
                <span>
                  <span style={s.ext}>{ext}</span>
                  <span style={{ ...s.desc, marginLeft: 10 }}>{describeExtension(ext)}</span>
                </span>
                <span style={{ fontSize: 11, color: "#8a8886" }}>in force</span>
              </div>
            ))}
          </div>
        </>
      ) : (
        <>
          {state === "all-blocked" && (
            <div style={{ ...s.msg, ...s.err }}>
              <strong>No file types are allowed — uploads are blocked for everyone.</strong> Files already
              uploaded are untouched. Switch on any type below to allow uploads again.
            </div>
          )}

          <div style={s.card}>
            {/* The "File type settings" heading came off 2026-08-30 at the client's request: the page
                is now titled File Type Management, so a card heading repeating it was one label too
                many on a page that does exactly one thing. */}
            <div style={s.hint}>
              Enable or disable file extensions. Applies to <strong>new uploads</strong> — anyone with the
              upload form open will need to refresh. Disabling a type never affects files already
              uploaded.
            </div>

            <div style={s.head}>
              <span>Extension</span>
              <span>Description</span>
              <span>Allowed</span>
              <span />
            </div>

            {rows.length === 0 && (
              <div style={{ fontSize: 13, color: "#8a8886", padding: 10 }}>
                The {COLUMN_INTERNAL_NAME} column has no choices yet — add the first file type below.
              </div>
            )}

            {rows.map((ext) => {
              const on = ticked.indexOf(ext) !== -1;
              const blocked = isBlockedType(ext);
              return (
                <div
                  key={ext}
                  style={{ ...s.row, background: highlight === ext ? "#fff8e1" : undefined }}
                >
                  <span>
                    <span style={s.badge}>{badgeFor(ext)}</span>
                  </span>
                  <span>
                    <span style={s.ext}>{ext}</span>
                    <span style={{ ...s.desc, marginLeft: 10 }}>{describeExtension(ext)}</span>
                    {blocked && <span style={{ ...s.locked, marginLeft: 10 }}>Blocked by policy</span>}
                  </span>
                  <span>
                    <Toggle
                      on={on && !blocked}
                      disabled={blocked}
                      busy={busyExt === ext}
                      label={`Allow ${ext}`}
                      onClick={() => onToggle(ext, on)}
                    />
                  </span>
                  <span>
                    {/* NOT offered for a blocked type. Those rows exist as the guard itself — the list
                        is how an admin sees that `.exe` is refused by policy, and removing the row
                        would remove the evidence while `BLOCKED_TYPES` goes on refusing it anyway. */}
                    {blocked ? (
                      <span style={{ fontSize: 11, color: "#c8c6c4" }}>—</span>
                    ) : (
                      <button
                        style={busyExt === ext || removeBusy ? s.removeOff : s.removeLink}
                        disabled={busyExt === ext || removeBusy}
                        title={`Remove ${ext} from the list`}
                        onClick={() => setRemovingExt(ext)}
                      >
                        Remove
                      </button>
                    )}
                  </span>
                </div>
              );
            })}

            <div style={{ display: "flex", alignItems: "center", gap: 10, paddingTop: 14, flexWrap: "wrap" }}>
              <input
                style={s.input}
                value={draft}
                placeholder=".txt"
                aria-label="New file extension"
                onChange={(e) => setDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") onAdd();
                }}
              />
              <button
                style={adding || draft.trim() === "" ? s.off : s.btn}
                disabled={adding || draft.trim() === ""}
                onClick={onAdd}
              >
                {adding ? "Adding…" : "+ Add file type"}
              </button>
              <span style={{ fontSize: 11, color: "#8a8886" }}>
                {/* "Adding needs site owner rights" removed 2026-08-30 — this page is admin-only, so
                    it warned the only people who already have the right. The refusal, if it ever
                    comes, is reported by the add itself. */}
                The dot is added for you.
              </span>
            </div>
          </div>
        </>
      )}

      {/* Remove confirmation. A plain confirm, not a typed one: removing a type is reversible in one
          click (add it again), and a typed gate on a reversible action is what teaches people to type
          through gates that are not. */}
      {removingExt !== undefined && (
        <div style={s.overlay}>
          <div style={s.modal}>
            <p style={{ ...s.h2, fontSize: 17 }}>Remove {removingExt}?</p>
            <p style={{ fontSize: 13, lineHeight: 1.6, color: "#323130" }}>
              It stops being offered on this page, and{" "}
              {ticked.indexOf(removingExt) !== -1 ? (
                <strong>can no longer be uploaded.</strong>
              ) : (
                <>it already could not be uploaded, so what may be uploaded does not change.</>
              )}
            </p>
            {/* The one genuinely consequential case: this was the last thing anyone could upload. Same
                fact the toggle's own last-one-off confirm exists for, said here too because removing is
                a different button and an admin may only ever meet this one. */}
            {ticked.length === 1 && ticked.indexOf(removingExt) !== -1 && (
              <div style={{ ...s.msg, ...s.warn }}>
                This is the only allowed file type left. Removing it <strong>blocks every upload
                across the whole site</strong> until another type is added.
              </div>
            )}
            <p style={{ fontSize: 12, lineHeight: 1.6, color: "#605e5c" }}>
              Files already uploaded are untouched — this only affects new uploads. You can add{" "}
              {removingExt} again at any time.
            </p>
            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button
                style={removeBusy ? s.off : s.danger}
                disabled={removeBusy}
                onClick={() => {
                  const ext = removingExt;
                  if (ext !== undefined) commitRemove(ext).catch(() => undefined);
                }}
              >
                {removeBusy ? "Removing…" : `Remove ${removingExt}`}
              </button>
              <button
                style={s.ghost}
                disabled={removeBusy}
                onClick={() => setRemovingExt(undefined)}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Did-you-mean. Usability, not security: a near miss fails CLOSED, blocking a legitimate type. */}
      {suggest && (
        <div style={s.overlay}>
          <div style={s.modal}>
            <p style={{ ...s.h2, fontSize: 17 }}>Check the spelling</p>
            <p style={{ fontSize: 13, lineHeight: 1.6, color: "#323130" }}>
              <strong>{suggest.typed}</strong> isn&apos;t a known file type — did you mean{" "}
              <strong>{suggest.better}</strong>?
            </p>
            <p style={{ fontSize: 12, lineHeight: 1.6, color: "#605e5c" }}>
              A mistyped extension blocks the real one: nobody could upload{" "}
              <strong>{suggest.better}</strong> files, and they would be told the file type is not
              permitted.
            </p>
            <div style={{ display: "flex", gap: 10, marginTop: 16, flexWrap: "wrap" }}>
              <button
                style={s.btn}
                onClick={() => {
                  const better = suggest.better;
                  setSuggest(undefined);
                  addOrPointAt(better);
                }}
              >
                Use {suggest.better}
              </button>
              <button
                style={s.ghost}
                onClick={() => {
                  const typed = suggest.typed;
                  setSuggest(undefined);
                  commitAdd(typed).catch(() => undefined);
                }}
              >
                Add {suggest.typed} anyway
              </button>
              <button style={s.ghost} onClick={() => setSuggest(undefined)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Turning off the last enabled type. Allowed, never silent. */}
      {confirmLast && (
        <div style={s.overlay}>
          <div style={s.modal}>
            <p style={{ ...s.h2, fontSize: 17 }}>Block all uploads?</p>
            <p style={{ fontSize: 13, lineHeight: 1.6, color: "#323130" }}>
              <strong>{confirmLast}</strong> is the only file type still allowed. Switching it off means:
            </p>
            <ul
              style={{
                fontSize: 13, lineHeight: 1.7, color: "#323130", paddingLeft: 20, margin: "8px 0 0",
              }}
            >
              <li>
                <strong>Nobody can upload anything</strong>, in any unit, until a type is switched back on
              </li>
              <li>Documents already uploaded are untouched and stay visible</li>
              <li>Switching any type back on restores uploads immediately</li>
            </ul>
            <div style={{ display: "flex", gap: 10, marginTop: 18, flexWrap: "wrap" }}>
              <button
                style={{ ...s.btn, background: "#a4262c" }}
                onClick={() => {
                  const ext = confirmLast;
                  setConfirmLast(undefined);
                  applyToggle(ext, false).catch(() => undefined);
                }}
              >
                Switch it off
              </button>
              <button style={s.ghost} onClick={() => setConfirmLast(undefined)}>
                Keep {confirmLast} allowed
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

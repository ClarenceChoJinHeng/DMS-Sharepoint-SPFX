// The side panel behind the library command bar's "Approve or reject".
//
// Built for FIFTY files at a time (client, 2026-08-25: "most probably 50 files"), which is what
// shapes almost every decision here: the file list scrolls, progress is per file and named, the work
// is sequential, and a failure never stops the run.
//
// ⚠ EVERY CHECK COMES FROM `shared/approvalGuards.ts`. This panel decides WHEN to check and what to
// say about the answer; it must never decide WHAT to check. A bulk route with weaker checks than the
// approval page is a faster way to overwrite documents and to publish into unlocked folders.
import * as React from "react";
import { useState } from "react";
import * as ReactDOM from "react-dom";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  checkApproveRight,
  checkDestinationClash,
  checkUnitFolderReady,
  type GuardResult,
} from "../../../shared/approvalGuards";
import { libraryHasColumns, APPROVED_BY_COLUMN } from "../../../shared/optionalColumns";
import {
  cachedHcLibraries,
  cachedListTitle,
  DOCUMENTS_URL_SEGMENT,
  documentsLibraryTitle,
  libraryUrlSegment,
  LIST_SUFFIX,
} from "../../../shared/naming";
import { parseLevels } from "../../../shared/formModel";

export type SelectedDoc = { id: number; name: string; fileSru: string };

type Decision = "approve" | "reject";
type Phase = "choose" | "running" | "done";
type Outcome = { doc: SelectedDoc; ok: boolean; reason?: string };

type PanelProps = {
  sp: SPHttpClient;
  webUrl: string;
  webSru: string;
  listTitle: string;
  docs: SelectedDoc[];
  /** The signed-in approver's email, for the `ApprovedBy` stamp (2026-09-01) — the panel has no
      other way to know who is running it: `SPHttpClient` carries no identity of its own. */
  approverEmail: string;
  onClose: () => void;
  onDone: () => void;
};

const s: Record<string, React.CSSProperties> = {
  scrim: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.2)", zIndex: 1000000 },
  panel: {
    position: "fixed", top: 0, right: 0, bottom: 0, width: 420, maxWidth: "100vw",
    background: "#fff", boxShadow: "-2px 0 12px rgba(0,0,0,0.18)", zIndex: 1000001,
    display: "flex", flexDirection: "column",
    font: '400 14px "Segoe UI", system-ui, sans-serif', color: "#323130",
  },
  head: { padding: "16px 20px", borderBottom: "1px solid #edebe9", display: "flex", alignItems: "center", justifyContent: "space-between" },
  title: { fontSize: 18, fontWeight: 600, margin: 0 },
  body: { padding: "16px 20px", overflowY: "auto", flex: 1 },
  foot: { padding: "12px 20px", borderTop: "1px solid #edebe9", display: "flex", gap: 8 },
  hint: { fontSize: 12, color: "#605e5c", lineHeight: 1.5 },
  label: { fontSize: 13, fontWeight: 600, margin: "16px 0 8px" },
  list: { maxHeight: 220, overflowY: "auto", border: "1px solid #edebe9", borderRadius: 4, padding: 8 },
  row: { fontSize: 12, padding: "3px 0", wordBreak: "break-all" },
  bad: { fontSize: 12, padding: "6px 0", color: "#a4262c", wordBreak: "break-word" },
  // `resize: none` — the panel is a fixed-width column, and a drag-resized textarea pushes the
  // document list and the footer buttons out of view with no way back short of reopening.
  textarea: { width: "100%", height: 90, boxSizing: "border-box", font: "inherit", fontSize: 13, padding: 8, border: "1px solid #8a8886", borderRadius: 2, resize: "none" },
  btn: { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 2, padding: "8px 16px", fontSize: 14, cursor: "pointer" },
  btnOff: { background: "#c8c6c4", color: "#fff", border: "none", borderRadius: 2, padding: "8px 16px", fontSize: 14, cursor: "not-allowed" },
  ghost: { background: "#fff", color: "#323130", border: "1px solid #8a8886", borderRadius: 2, padding: "8px 16px", fontSize: 14, cursor: "pointer" },
  bar: { height: 4, background: "#edebe9", borderRadius: 2, overflow: "hidden", margin: "12px 0" },
  barFill: { height: "100%", background: "#0f6c3f" },
};

function BulkApprovePanel(p: PanelProps): React.ReactElement {
  const [decision, setDecision] = useState<Decision>("approve");
  const [comment, setComment] = useState("");
  const [phase, setPhase] = useState<Phase>("choose");
  const [step, setStep] = useState({ done: 0, total: p.docs.length, what: "" });
  const [results, setResults] = useState<Outcome[]>([]);

  /** Which library an approved document lands in — the other half of whichever pair this one is. */
  const hc = cachedHcLibraries();
  const isHc = hc !== undefined && p.listTitle.toLowerCase() === hc.approval.title.toLowerCase();
  const sourceSegment = isHc && hc ? hc.approval.urlSegment : libraryUrlSegment();
  // Blank rather than a fallback for an unresolved HC pair: falling back to the open library would
  // point an HC document's checks — and the message an approver acts on — at the wrong place.
  const destSegment = isHc ? (hc ? hc.documents.urlSegment : "") : DOCUMENTS_URL_SEGMENT;
  // Display only — `checkUnitFolderReady` puts it in the message an approver acts on. Resolved
  // since 2026-08-28: naming a library that no longer exists sends them looking for the wrong thing.
  const destTitle = isHc ? (hc ? hc.documents.title : "HC Documents") : documentsLibraryTitle();

  /**
   * Permissioned tier count per segment folder, from the mode rows' `Levels`.
   *
   * Read ONCE for the whole run, not per file. It cannot come from the documents themselves: every
   * tier column has a `<Base>Tid` twin and a below-Unit tier such as SubUnit carries one exactly like
   * a permissioned tier, so counting those lands on a folder that inherits by design.
   */
  const loadTierCounts = async (): Promise<Record<string, number>> => {
    const out: Record<string, number> = {};
    try {
      const res: SPHttpClientResponse = await p.sp.get(
        `${p.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
          `/items?$select=Title,StagingFolder,Levels&$filter=ConfigType eq 'mode'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return out;
      const data = await res.json();
      for (const r of (data.value ?? []) as Array<Record<string, string>>) {
        const folder = (r.StagingFolder ?? "").trim().toLowerCase();
        if (!folder) continue;
        const tiers = parseLevels(r.Levels ?? "").filter((l) => l.permissioned !== false).length;
        // A segment whose Levels will not parse is LEFT OUT rather than stored as 0: absent reads as
        // unknown and the guard refuses, where a stored 0 would claim we know the segment is flat.
        if (tiers > 0) out[folder] = tiers;
      }
    } catch {
      /* Left empty — every guard then refuses by name, which is the safe direction. */
    }
    return out;
  };

  const segmentFolderOf = (fileSru: string): string => {
    const prefix = `${p.webSru}/${sourceSegment}/`.toLowerCase();
    if (!sourceSegment || fileSru.toLowerCase().indexOf(prefix) !== 0) return "";
    return fileSru.slice(prefix.length).split("/")[0] ?? "";
  };

  /* ⚠ `OData__ModerationStatus` = 0 (a MERGE) DOES NOT RESTAMP `Editor` EITHER — proven live
     2026-09-01, and it was the first theory tested before `File.approve()` was ruled out too.
     Neither approval route restamps it, because setting moderation status is not an edit. So this
     panel needs the SAME `ApprovedBy` stamp as the approval page, or a document approved through
     here still names its uploader as the approver and still suppresses the notification email. */
  const setStatus = async (doc: SelectedDoc, stampApprover: boolean): Promise<GuardResult> => {
    const itemUrl = `${p.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(p.listTitle)}')/items(${doc.id})`;
    const mergeHeaders = {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };
    // ⚠ SHAREPOINT REJECTS A MERGE THAT SETS OData__ModerationStatus ALONGSIDE ANY OTHER FIELD —
    // proven live 2026-09-01 (500, "You cannot change moderation status and set other item
    // properties at that same time") on the approval page's identical combined-MERGE attempt.
    // ModerationStatus + ModerationComments together IS allowed (both are moderation fields); it is
    // specifically an UNRELATED field like ApprovedBy in the same call that 500s. So the ApprovedBy
    // stamp has to be its own request, done FIRST and best-effort (a failure here must never be
    // reported as a failed approval — `setStatus`'s caller only sees the result below). Writing it
    // first, then the real moderation write, means the moderation write's own status/comment values
    // are what stand at the end even though the ApprovedBy edit alone would otherwise revert the item
    // to Pending — the second write re-asserts the real outcome regardless of order.
    if (stampApprover) {
      try {
        await p.sp.post(itemUrl, SPHttpClient.configurations.v1, {
          headers: mergeHeaders,
          body: JSON.stringify({ ApprovedBy: p.approverEmail }),
        });
      } catch {
        /* Best-effort — see comment above. */
      }
    }
    const res: SPHttpClientResponse = await p.sp.post(itemUrl, SPHttpClient.configurations.v1, {
      headers: mergeHeaders,
      body: JSON.stringify({
        OData__ModerationStatus: decision === "approve" ? 0 : 1,
        OData__ModerationComments: comment.slice(0, 500),
      }),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      return { ok: false, reason: `the update was refused (HTTP ${res.status}) ${body.slice(0, 120)}` };
    }
    return { ok: true };
  };

  const run = async (): Promise<void> => {
    setPhase("running");
    const tierCounts = decision === "approve" ? await loadTierCounts() : {};
    // Checked ONCE for the whole run — `p.listTitle` is fixed for the panel's lifetime, so the
    // answer cannot change between files. A rejection never approves anything, so it never needs it.
    const stampApprover =
      decision === "approve" &&
      (await libraryHasColumns(p.sp, p.webUrl, p.listTitle, [APPROVED_BY_COLUMN]).catch(() => false));
    /* The unit-folder check is per UNIT, not per file: fifty documents in one folder ask once.
       Keyed on the folder path the guard would test, so two units in one selection still get two
       checks and a third file in either reuses the answer. */
    const folderVerdict: Record<string, GuardResult> = {};
    const out: Outcome[] = [];

    for (let i = 0; i < p.docs.length; i++) {
      const doc = p.docs[i];
      setStep({ done: i, total: p.docs.length, what: doc.name });

      // Always: may this person decide this document at all. Cheap, and the only check a REJECTION
      // needs — rejecting copies nothing anywhere, so there is no destination to verify.
      const right = await checkApproveRight({
        sp: p.sp, webUrl: p.webUrl, listTitle: p.listTitle, itemId: doc.id,
      });
      if (right === "denied") {
        out.push({ doc, ok: false, reason: "you are not the approver for that unit" });
        setResults(out.slice());
        continue;
      }

      if (decision === "approve") {
        const unitKey = doc.fileSru.slice(0, doc.fileSru.lastIndexOf("/")).toLowerCase();
        if (!folderVerdict[unitKey]) {
          folderVerdict[unitKey] = await checkUnitFolderReady({
            sp: p.sp, webUrl: p.webUrl, webSru: p.webSru,
            fileSru: doc.fileSru,
            sourceSegment, destSegment, destLibTitle: destTitle,
            permissionedTiers: tierCounts[segmentFolderOf(doc.fileSru).toLowerCase()],
          });
        }
        const folderOk = folderVerdict[unitKey];
        if (!folderOk.ok) {
          out.push({ doc, ok: false, reason: `the destination folder is not ready — ${folderOk.reason}` });
          setResults(out.slice());
          continue;
        }
        // Per FILE, never cached: it is about this document's name at this moment.
        const clash = await checkDestinationClash({
          sp: p.sp, webUrl: p.webUrl, fileSru: doc.fileSru, fileName: doc.name,
          sourceSegment, destSegment,
        });
        if (!clash.ok) {
          out.push({ doc, ok: false, reason: clash.reason });
          setResults(out.slice());
          continue;
        }
      }

      const wrote = await setStatus(doc, stampApprover).catch((e) => ({ ok: false, reason: String(e) }));
      out.push({ doc, ok: wrote.ok, reason: wrote.reason });
      setResults(out.slice());
    }

    setStep({ done: p.docs.length, total: p.docs.length, what: "" });
    setPhase("done");
  };

  const okCount = results.filter((r) => r.ok).length;
  const badCount = results.length - okCount;
  const pct = step.total === 0 ? 0 : Math.round((step.done / step.total) * 100);

  return (
    <>
      <div style={s.scrim} onClick={() => { if (phase !== "running") p.onClose(); }} />
      <div style={s.panel} role="dialog" aria-label="Approve or reject documents">
        <div style={s.head}>
          <h2 style={s.title}>Approval</h2>
          {phase !== "running" && (
            <button type="button" style={s.ghost} onClick={p.onClose} aria-label="Close">✕</button>
          )}
        </div>

        <div style={s.body}>
          {phase === "choose" && (
            <>
              <p style={s.hint}>
                Please review the documents and their details.
              </p>

              <div style={s.label}>Approval status</div>
              {(["approve", "reject"] as Decision[]).map((d) => (
                <label key={d} style={{ display: "block", marginBottom: 8, cursor: "pointer" }}>
                  <input
                    type="radio" name="crs-bulk-decision" checked={decision === d}
                    onChange={() => setDecision(d)} style={{ marginRight: 8 }}
                  />
                  {d === "approve" ? "Approve" : "Reject"}
                </label>
              ))}

              <div style={s.label}>Comment</div>
              <textarea
                style={s.textarea}
                value={comment}
                maxLength={500}
                placeholder="Use this field to enter any comments about why the items were approved or rejected."
                onChange={(e) => setComment(e.target.value.slice(0, 500))}
              />
              {/* Said plainly, because a rejection reason is normally per document and this one is
                  not. An approver who assumes otherwise writes "wrong vendor" onto fifty files. */}
              <p style={s.hint}>
                {comment.length}/500 — this same comment is written to all {p.docs.length} documents.
              </p>

              <div style={s.label}>Documents</div>
              <div style={s.list}>
                {p.docs.map((d) => <div key={d.id} style={s.row}>{d.name}</div>)}
              </div>
              <p style={s.hint}>
                {decision === "approve"
                  /* ⚠ SHORTENED 2026-08-30 at the client's request. The three checks it used to
                     recite — destination exists, folder locked down, no name clash — ALL STILL RUN,
                     per file, and a failure still names the file and the reason. Only the preamble
                     is gone. Do not read the shorter sentence as the guards having been relaxed. */
                  ? `Approved documents are automatically moved to ${destTitle} after approval.`
                  : "Rejected documents stay in this library, so the uploader can see the comment and fix them."}
              </p>
            </>
          )}

          {phase === "running" && (
            <>
              <p style={s.hint}>
                {decision === "approve" ? "Approving" : "Rejecting"} {step.done} of {step.total}…
              </p>
              <div style={s.bar}><div style={{ ...s.barFill, width: `${pct}%` }} /></div>
              <p style={s.hint}>{step.what}</p>
              <p style={s.hint}>Leave this panel open — there is no resume.</p>
            </>
          )}

          {phase === "done" && (
            <>
              <p style={{ fontSize: 14, fontWeight: 600 }}>
                {okCount} {decision === "approve" ? "approved" : "rejected"}
                {badCount > 0 ? `, ${badCount} skipped` : ""}
              </p>
              {badCount > 0 && (
                <>
                  {/* NAMED, never counted. "3 skipped" is unactionable; the whole point of doing
                      fifty at once is not having to work out afterwards which three. */}
                  <div style={s.label}>Not done, and why</div>
                  {results.filter((r) => !r.ok).map((r) => (
                    <div key={r.doc.id} style={s.bad}>
                      <strong>{r.doc.name}</strong> — {r.reason}
                    </div>
                  ))}
                  <p style={s.hint}>
                    Everything above is unchanged and still pending. Fix the reason and select them
                    again.
                  </p>
                </>
              )}
            </>
          )}
        </div>

        <div style={s.foot}>
          {phase === "choose" && (
            <>
              <button type="button" style={p.docs.length === 0 ? s.btnOff : s.btn}
                      disabled={p.docs.length === 0} onClick={() => { run().catch(() => setPhase("done")); }}>
                Proceed
              </button>
              <button type="button" style={s.ghost} onClick={p.onClose}>Cancel</button>
            </>
          )}
          {phase === "running" && <button type="button" style={s.btnOff} disabled>Working…</button>}
          {phase === "done" && (
            <button type="button" style={s.btn} onClick={p.onDone}>Close and refresh</button>
          )}
        </div>
      </div>
    </>
  );
}

/**
 * Mount the panel over the library view.
 *
 * Its own container appended to `body`, removed on close: a list view is not our DOM, and leaving a
 * node behind would accumulate one per invocation.
 */
export function openBulkApprovePanel(o: {
  sp: SPHttpClient;
  webUrl: string;
  webSru: string;
  listTitle: string;
  docs: SelectedDoc[];
  approverEmail: string;
  onDone: () => void;
}): void {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const unmount = (): void => {
    ReactDOM.unmountComponentAtNode(host);
    if (host.parentNode) host.parentNode.removeChild(host);
  };
  ReactDOM.render(
    React.createElement(BulkApprovePanel, {
      sp: o.sp,
      webUrl: o.webUrl,
      webSru: o.webSru,
      listTitle: o.listTitle,
      docs: o.docs,
      approverEmail: o.approverEmail,
      onClose: unmount,
      onDone: () => { unmount(); o.onDone(); },
    }),
    host,
  );
}

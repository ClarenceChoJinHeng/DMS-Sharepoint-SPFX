// "How roles and permission levels work" — the reference card.
//
// ⚠ THIS TEXT HAD BEEN RENDERING NOWHERE SINCE 2026-08-23. It lived inside `GroupMapBuilder`, gated
// on `show !== "members"`, and that mount was retired when hand-mapping left Group Management — so
// the whole block became unreachable along with the form it sat above. CLAUDE.md flagged the loss at
// the time: it is **the only place naming how to create `CRS Upload` / `CRS Approve` / `CRS Delete`**
// (Site settings → Permission levels) and what a missing level costs. Both live sites already have
// the levels, so nothing was broken — but a third site provisioned without this guidance would meet
// reconciliation's `no "CRS Approve" role definition on site` line with no idea what to do about it.
//
// Brought back on the Group Management step of the guided flows (client, 2026-08-30), which is where
// an admin creating groups needs it now that Folder Access is gone.
//
// ⚠ ONE DEFINITION. `GroupMapBuilder` renders this component rather than keeping its own copy. Two
// copies of a reference nobody reads often are exactly the pair that drifts, and the drifting one is
// always the rarely-seen one.
import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";

import { permissionLevelNames } from "./spNaming";

const s: Record<string, React.CSSProperties> = {
  /* Every key used below must EXIST here: `s` is a `Record<string, CSSProperties>`, so a missing key
     yields `undefined` and the element renders unstyled with a completely green build. */
  card: {
    border: "1px solid #e6e6e6", borderRadius: 12, background: "#fff",
    padding: "18px 20px", marginBottom: 16,
  },
  head: {
    display: "flex", alignItems: "center", gap: 10, width: "100%", background: "none",
    border: "none", padding: 0, cursor: "pointer", font: "inherit", textAlign: "left",
  },
  headText: {
    fontSize: 13, fontWeight: 700, color: "#0f6c3f", textTransform: "uppercase",
    letterSpacing: ".04em", flex: "1 1 auto",
  },
  row: { display: "flex", gap: 12, alignItems: "flex-start", marginTop: 16 },
  rowIcon: { flexShrink: 0, color: "#0f6c3f", marginTop: 2 },
  body: { fontSize: 12.5, color: "#3b3a39", lineHeight: 1.6, margin: 0 },
  chips: { display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 },
  chip: {
    display: "inline-flex", alignItems: "center", gap: 6, borderRadius: 6,
    padding: "5px 10px", fontSize: 11.5, fontWeight: 600, whiteSpace: "nowrap",
  },
  list: { margin: "8px 0 0", paddingLeft: 18, fontSize: 12.5, color: "#3b3a39", lineHeight: 1.7 },
};

/* Coloured by what the role DOES, not decoratively. The two delete roles deliberately share a
   family: they are one letter apart and are the pair people confuse. */
const CHIP_TONE: Record<string, React.CSSProperties> = {
  read:    { background: "#eef7f1", color: "#0f6c3f" },
  upload:  { background: "#eef3fb", color: "#1d4f91" },
  approve: { background: "#eef7f1", color: "#0f6c3f" },
  del:     { background: "#fdf1e7", color: "#8a4b00" },
  dels:    { background: "#f3eefb", color: "#5b2d90" },
};

function Chip({ tone, children }: { tone: string; children: React.ReactNode }): React.ReactElement {
  return <span style={{ ...s.chip, ...(CHIP_TONE[tone] ?? CHIP_TONE.read) }}>{children}</span>;
}

export interface IRolesReferenceProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Start expanded. The flow step opens it; the (unmounted) mapping form kept it collapsed. */
  defaultOpen?: boolean;
}

export function RolesReference(
  { spHttpClient, siteUrl, defaultOpen }: IRolesReferenceProps,
): React.ReactElement {
  const [open, setOpen] = React.useState(defaultOpen !== false);
  /**
   * The three custom levels, named for THIS site.
   *
   * ⚠ RESOLVED, NEVER HARDCODED. Telling an admin to create "DMS Upload" on a site where
   * reconciliation grants "CRS Upload" sends them to build a level nothing will ever use — and the
   * resulting log line looks identical either way. The values below are only what shows for the
   * moment before the read lands.
   */
  const [levels, setLevels] = React.useState({ upload: "Upload", approve: "Approve", del: "Delete" });

  React.useEffect(() => {
    let live = true;
    permissionLevelNames(spHttpClient, siteUrl)
      .then((n) => {
        if (live) setLevels(n);
      })
      // A failed read leaves the generic names standing. This card is guidance, and guidance that
      // cannot be shown at all is worse than guidance with a slightly wrong noun in it.
      .catch(() => undefined);
    return () => {
      live = false;
    };
  }, [spHttpClient, siteUrl]);

  return (
    <div style={s.card}>
      <button type="button" style={s.head} onClick={() => setOpen((v) => !v)} aria-expanded={open}>
        <span style={{ color: "#0f6c3f", flexShrink: 0 }} aria-hidden="true">
          <svg width="18" height="18" viewBox="0 0 20 20" fill="none">
            <path
              d="M3 4.5h5.5c.8 0 1.5.7 1.5 1.5v10c0-.8-.7-1.5-1.5-1.5H3v-10zM17 4.5h-5.5c-.8 0-1.5.7-1.5 1.5v10c0-.8.7-1.5 1.5-1.5H17v-10z"
              stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round"
            />
          </svg>
        </span>
        <span style={s.headText}>How roles and permission levels work</span>
        <span style={{ color: "#605e5c", flexShrink: 0 }} aria-hidden="true">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="none"
            style={{ transform: open ? "rotate(180deg)" : undefined }}>
            <path d="M3.5 6L8 10.5 12.5 6" stroke="currentColor" strokeWidth="1.6"
              strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </span>
      </button>

      {open ? (
        <>
          <div style={s.row}>
            <span style={s.rowIcon} aria-hidden="true">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M9 1.5 3.5 9H7l-.5 5.5L12.5 7H9l0-5.5z" stroke="currentColor"
                  strokeWidth="1.4" strokeLinejoin="round" />
              </svg>
            </span>
            <div>
              <p style={s.body}>
                Member changes take effect <strong>immediately</strong>; new or deleted rows need a{" "}
                <strong>Folder Reconciliation</strong> run to apply folder permissions.
              </p>
              <div style={s.chips}>
                <Chip tone="read">MEMBER &rarr; Read</Chip>
                <Chip tone="upload">UPL &rarr; {levels.upload}</Chip>
                <Chip tone="approve">APR &rarr; {levels.approve}</Chip>
                <Chip tone="del">DEL &rarr; {levels.del} on Documents</Chip>
                <Chip tone="dels">DELS &rarr; {levels.del} on approval library</Chip>
              </div>
            </div>
          </div>

          <div style={s.row}>
            <span style={s.rowIcon} aria-hidden="true">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <circle cx="6" cy="5" r="2.6" stroke="currentColor" strokeWidth="1.4" />
                <path d="M1.8 13.5c0-2.3 1.9-3.8 4.2-3.8 1 0 1.9.3 2.6.8" stroke="currentColor"
                  strokeWidth="1.4" strokeLinecap="round" />
                <circle cx="12" cy="11.5" r="2.4" stroke="currentColor" strokeWidth="1.4" />
              </svg>
            </span>
            <div>
              {/* ⚠ UPLOAD IS CALLED OUT SEPARATELY FROM THE OTHER TWO because its failure mode is
                  worse. APR and DEL previously had no level at all, so a missing one only meant
                  "not yet in effect". UPL used to point at Contribute, which always exists — so once
                  it points at the custom level, a site without that level gives every NEWLY
                  provisioned unit no uploader grant whatsoever. Existing uploaders keep the
                  Contribute grant already on their folder (nothing revokes it), which is exactly why
                  nobody notices until a new unit is onboarded. */}
              <p style={s.body}>
                <strong>{levels.upload}</strong>, <strong>{levels.approve}</strong> and{" "}
                <strong>{levels.del}</strong> are custom permission levels an administrator creates
                once per site <em>(Site settings &rarr; Site permissions &rarr; Permission levels)</em>:
              </p>
              <ul style={s.list}>
                <li>Copy <em>Contribute</em> and untick Delete Items</li>
                <li>
                  Copy <em>Contribute</em>, tick Approve Items and untick Add Items, Delete Items and
                  Delete Versions
                </li>
                <li>Copy <em>Read</em> and tick Delete Items</li>
              </ul>
              <p style={{ ...s.body, marginTop: 10 }}>
                Until they exist, reconciliation reports no{" "}
                <em>&ldquo;{levels.approve}&rdquo;</em> role definition on site and skips those grants
                &mdash; nobody loses access, but approve-only and delete do not take effect, and a
                newly provisioned unit gets no uploader grant at all.
              </p>
            </div>
          </div>

          <div style={s.row}>
            <span style={s.rowIcon} aria-hidden="true">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M3 4.5h10M6.5 4.5V3h3v1.5M4.5 4.5l.7 9h5.6l.7-9" stroke="currentColor"
                  strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </span>
            {/* ⚠ DEL vs DELS IS ONE LETTER FOR TWO DIFFERENT LIBRARIES, so it is spelled out here as
                well as in the role tooltips. An admin who picks DEL intending "can clear out junk in
                the approval library" grants delete over APPROVED documents instead — and the run log
                looks identical either way. */}
            <p style={s.body}>
              <strong>DEL</strong> deletes <em>approved</em> documents in the Documents library.
              <br />
              <strong>DELS</strong> deletes <em>pending</em> files in the approval library. They share
              one permission level and differ only in which library they are allowed to reach, so
              picking the wrong one grants delete over the wrong set of documents.
              <br />
              Uploaders (<strong>UPL</strong>) cannot delete at all &mdash; that is deliberate, so a
              PIC must ask a head of unit.
            </p>
          </div>
        </>
      ) : undefined}
    </div>
  );
}

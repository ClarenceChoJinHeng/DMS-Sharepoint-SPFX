import * as React from "react";

/**
 * Common chrome for the four access pages.
 *
 * Split out of Folder & Group Manager on 2026-08-07 (spec
 * `2026-08-07-access-webpart-split-design.md`), then split again into ONE PAGE PER SURFACE at
 * the client's direction. Each page is its own web part with its own Page Access row, so the
 * person who manages folder mappings need not be the person who manages site entry.
 *
 * This exists so the four pages cannot drift apart visually — they are peers and should read
 * as peers.
 */
const s: Record<string, React.CSSProperties> = {
  wrap:     { maxWidth: 880, margin: "32px auto", padding: "0 24px 48px", fontFamily: "'Segoe UI', sans-serif" },
  h2:       { fontSize: 22, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle: { fontSize: 13, color: "#666", margin: "0 0 24px" },
  note:     { fontSize: 12, color: "#605e5c", lineHeight: 1.5, margin: "0 0 20px", padding: "10px 12px", background: "#f6f8f6", border: "1px solid #e1e8e3", borderRadius: 6 },
};

export interface AccessShellProps {
  title: string;
  subtitle: string;
  /** Only Folder Access needs one — see FolderAccessPage. */
  note?: React.ReactNode;
  children: React.ReactNode;
}

export default function AccessShell({ title, subtitle, note, children }: AccessShellProps): React.ReactElement {
  return (
    <section style={s.wrap}>
      <h2 style={s.h2}>{title}</h2>
      <p style={s.subtitle}>{subtitle}</p>
      {note !== undefined && <p style={s.note}>{note}</p>}
      {children}
    </section>
  );
}

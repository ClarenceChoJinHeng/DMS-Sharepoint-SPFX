import * as React from "react";
import { useEffect, useState } from "react";
import { primeNames } from "../../../shared/spNaming";
import { libraryTitle } from "../../../shared/naming";
import { IUserAccessProps } from "./IUserAccessProps";
import GroupMapBuilder from "./GroupMapBuilder";
import SiteAccess from "./SiteAccess";
import StagingAccess from "./StagingAccess";
import PageAccess from "./PageAccess";

/**
 * User Access — who can reach what.
 *
 * Split out of Folder & Group Manager on 2026-08-07 (spec
 * `2026-08-07-access-webpart-split-design.md`). These four surfaces used to sit as SUB-tabs
 * under a "User Access" tab, one level below folder browsing and reconciliation — so the task
 * done most often (a person joins, moves, or leaves) was buried under the ones done least.
 * The client reported that as "too complicated"; the nesting was the complexity.
 *
 * Each child is self-contained and primes its own names, so this shell only picks between them.
 */
type AccessTab = "Folder" | "Site" | "Library" | "Page";

const s: Record<string, React.CSSProperties> = {
  wrap:      { maxWidth: 880, margin: "32px auto", padding: "0 24px 48px", fontFamily: "'Segoe UI', sans-serif" },
  h2:        { fontSize: 22, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle:  { fontSize: 13, color: "#666", margin: "0 0 24px" },
  toggleWrap:{ display: "flex", justifyContent: "center", marginBottom: 24 },
  // Matches Folder Manager's top-level bar: these are peers of that page, not a nested level.
  // flexWrap because "Approval Document Access" is long once the title resolves.
  seg:       { display: "flex", flexWrap: "wrap", justifyContent: "center", maxWidth: "100%", border: "1px solid #0f6c3f", borderRadius: 8, overflow: "hidden" },
  segBtn:    { padding: "8px 16px", fontSize: 13, fontFamily: "'Segoe UI', sans-serif", fontWeight: 600, cursor: "pointer", background: "#fff", color: "#0f6c3f", border: "none", borderRight: "1px solid #0f6c3f" },
  segActive: { background: "#0f6c3f", color: "#fff" },
  note:      { fontSize: 12, color: "#605e5c", lineHeight: 1.5, margin: "0 0 20px", padding: "10px 12px", background: "#f6f8f6", border: "1px solid #e1e8e3", borderRadius: 6 },
};

export default function UserAccess({ context }: IUserAccessProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const [tab, setTab] = useState<AccessTab>("Folder");
  // The library tab's LABEL is the live title ("Approval Document"), never the logical key.
  // Primed here purely so the tab bar reads correctly on first paint; the children prime
  // independently for their own reads.
  const [libLabel, setLibLabel] = useState<string>(libraryTitle());

  useEffect(() => {
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      .then(() => setLibLabel(libraryTitle()))
      .catch(() => undefined);
  }, []);

  const tabs: Array<{ key: AccessTab; label: string }> = [
    { key: "Folder",  label: "Folder Access" },
    { key: "Site",    label: "Site Access" },
    // Resolved, not hardcoded: the library is titled "Approval Document" on this site and
    // "Staging" only survives as a LOGICAL key. A hardcoded label here would be the one place
    // a client still sees the retired name.
    { key: "Library", label: `${libLabel} Access` },
    { key: "Page",    label: "Page Access" },
  ];

  return (
    <section style={s.wrap}>
      <h2 style={s.h2}>User Access</h2>
      <p style={s.subtitle}>Decide who can reach each folder, the site, the approval library and each page.</p>

      <div style={s.toggleWrap}>
        <div style={s.seg}>
          {tabs.map((t, i, arr) => (
            <button
              key={t.key}
              onClick={() => setTab(t.key)}
              style={{ ...s.segBtn, ...(i === arr.length - 1 ? { borderRight: "none" } : {}), ...(tab === t.key ? s.segActive : {}) }}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>

      {/* The seam created by splitting the pages: a mapping written here changes no folder ACL
          until reconciliation runs on Folder Manager. That was always true, but the two used to
          be one click apart — so it now has to be said out loud. Member changes are the
          exception and take effect immediately, because they change the GROUP, not the grant. */}
      <p style={s.note}>
        Adding or removing <strong>members</strong> of a group takes effect immediately. Creating
        or deleting a <strong>mapping</strong> only changes folder permissions after a{" "}
        <strong>Folder Reconciliation</strong> run on the Folder Manager page.
      </p>

      {tab === "Folder" ? (
        <GroupMapBuilder context={context} siteUrl={siteUrl} />
      ) : tab === "Site" ? (
        <SiteAccess context={context} siteUrl={siteUrl} />
      ) : tab === "Library" ? (
        // "Staging" is the LOGICAL LibTarget key, mapped to the live title at the API
        // boundary — not the library's name. See naming.ts / libApiTitle.
        <StagingAccess context={context} siteUrl={siteUrl} library="Staging" />
      ) : (
        <PageAccess context={context} siteUrl={siteUrl} />
      )}
    </section>
  );
}

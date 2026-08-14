// CRS Settings — the landing page for every admin screen.
//
// Spec: docs/superpowers/specs/2026-08-14-crs-settings-landing-page-design.md
//
// Built to the client's own mockup: a two-column grid of cards, each with a pastel icon, a blurb, and
// either rows of links or a single arrow on the card itself. What it lists and where each row points
// live in shared/adminPages.ts — pure and tested, because a landing page that sends an admin to the
// wrong screen is worse than no landing page at all.
//
// It is a DIRECTORY, deliberately, at the client's decision (2026-08-14). It does not teach the order
// of work — the more common failure, since missing the abbreviation step makes reconciliation create
// nothing silently — and the client has that covered separately.
import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  AdminCard,
  AdminLink,
  LinkTarget,
  SitePage,
  columnCards,
  resolveAll,
  suggestedPageName,
} from "../../../shared/adminPages";
import { CardIcon } from "./icons";

export interface CrsSettingsProps {
  context: WebPartContext;
  siteUrl: string;
  /** Per-link URL overrides from the property pane, keyed by AdminLink.key. */
  overrides: Record<string, string>;
  /** Optional heading override, so the client can retitle without a redeploy. */
  heading?: string;
  subheading?: string;
}

const s: Record<string, React.CSSProperties> = {
  wrap:      { fontFamily: '"Segoe UI", system-ui, sans-serif', color: "#242424", padding: "4px 0 24px" },
  head:      { padding: "0 0 18px", borderBottom: "1px solid #eceaea", marginBottom: 22 },
  h1:        { margin: 0, fontSize: 26, fontWeight: 600, letterSpacing: "-0.01em" },
  sub:       { margin: "6px 0 0", fontSize: 13.5, color: "#5f5f5f" },
  // Two columns that collapse to one. `minmax(min(100%, 380px), 1fr)` rather than `1fr`: a long blurb
  // in a grid child otherwise refuses to shrink and pushes the second column off screen.
  grid:      { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 380px), 1fr))", gap: 20, alignItems: "start" },
  col:       { display: "flex", flexDirection: "column", gap: 20 },
  card:      { border: "1px solid #e8e6e6", borderRadius: 12, background: "#fff", padding: 20, boxShadow: "0 1px 2px rgba(0,0,0,.04)" },
  cardHead:  { display: "flex", alignItems: "flex-start", gap: 14 },
  cardText:  { flex: 1, minWidth: 0 },
  cardTitle: { margin: 0, fontSize: 17, fontWeight: 600, lineHeight: 1.3 },
  cardBlurb: { margin: "5px 0 0", fontSize: 12.5, color: "#616161", lineHeight: 1.5 },
  rows:      { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 190px), 1fr))", gap: 12, marginTop: 18 },
  row:       { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "12px 14px", border: "1px solid #e8e6e6", borderRadius: 8, background: "#fff", textDecoration: "none", color: "#242424", fontSize: 12.5, fontWeight: 500, lineHeight: 1.35, cursor: "pointer" },
  rowDead:   { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "12px 14px", border: "1px dashed #dcdcdc", borderRadius: 8, background: "#fafafa", color: "#9a9a9a", fontSize: 12.5, fontWeight: 500, lineHeight: 1.35, cursor: "default" },
  arrow:     { color: "#0f6c3f", fontSize: 15, flexShrink: 0 },
  arrowBox:  { display: "flex", alignItems: "center", justifyContent: "center", width: 38, height: 38, border: "1px solid #e8e6e6", borderRadius: 8, color: "#0f6c3f", fontSize: 15, textDecoration: "none", flexShrink: 0 },
  note:      { fontSize: 11, color: "#8a8886", marginTop: 4, lineHeight: 1.4 },
  warnNote:  { fontSize: 11, color: "#8a4b00", marginTop: 4, lineHeight: 1.4 },
  danger:    { marginBottom: 20, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 8, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  loading:   { fontSize: 13, color: "#8a8886", padding: "8px 0" },
  mono:      { fontFamily: "Consolas, monospace", fontSize: 11 },
};

/** One clickable row, or a disabled one naming the page to create. */
function Row({ link, target }: { link: AdminLink; target: LinkTarget }): React.ReactElement {
  if (target === undefined || target.state === "missing") {
    // NOT a dead arrow. A navigation page whose links go nowhere teaches the client the tool is
    // broken, so the row says what is missing and what to call it.
    return (
      <div style={s.rowDead} title={`No page found for ${link.label}`}>
        <span>
          {link.label}
          <div style={s.note}>
            No page yet — create <span style={s.mono}>{suggestedPageName(link)}</span>
          </div>
        </span>
      </div>
    );
  }
  return (
    <a style={s.row} href={target.url}>
      <span>
        {link.label}
        {target.state === "ambiguous" && (
          // Surfaced rather than silently picked: a duplicated or half-renamed page is exactly the
          // state where the arrow works and lands somewhere nobody meant.
          <div style={s.warnNote}>
            More than one page matches — also {target.others.join(", ")}
          </div>
        )}
      </span>
      <span style={s.arrow}>&#8594;</span>
    </a>
  );
}

function Card({
  card,
  targets,
}: {
  card: AdminCard;
  targets: Record<string, LinkTarget>;
}): React.ReactElement {
  const self = card.self ? targets[card.self.key] : undefined;
  const selfMissing = self === undefined || self.state === "missing";
  return (
    <div style={s.card}>
      <div style={s.cardHead}>
        <CardIcon name={card.icon} />
        <div style={s.cardText}>
          <h2 style={s.cardTitle}>{card.title}</h2>
          <p style={s.cardBlurb}>{card.blurb}</p>
          {/* A single-link card carries its arrow on the header, as in the mockup. When the page is
              missing the arrow is replaced by the guidance a row would give — the card must not look
              identical to a working one. */}
          {card.self && selfMissing && (
            <div style={s.warnNote}>
              No page yet — create <span style={s.mono}>{suggestedPageName(card.self)}</span>
            </div>
          )}
          {card.self && self && self.state === "ambiguous" && (
            <div style={s.warnNote}>
              More than one page matches — also {self.others.join(", ")}
            </div>
          )}
        </div>
        {card.self && self && self.state !== "missing" && (
          <a style={s.arrowBox} href={self.url} title={card.title}>
            &#8594;
          </a>
        )}
      </div>
      {card.links.length > 0 && (
        <div style={s.rows}>
          {card.links.map((l) => (
            <Row key={l.key} link={l} target={targets[l.key]} />
          ))}
        </div>
      )}
    </div>
  );
}

export default function CrsSettings(props: CrsSettingsProps): React.ReactElement {
  const [pages, setPages] = useState<SitePage[] | undefined>(undefined);
  const [error, setError] = useState<string | undefined>(undefined);

  /**
   * The site's pages, read once.
   *
   * `FileRef` rather than a path assembled from the library name: a list's URL is independent of its
   * title (gotcha #12), and `SitePages` is not guaranteed to be the segment on every site.
   */
  useEffect(() => {
    const load = async (): Promise<void> => {
      const res: SPHttpClientResponse = await props.context.spHttpClient.get(
        `${props.siteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items` +
          `?$select=Id,FileLeafRef,Title,FileRef&$top=500`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) throw new Error(`Site Pages: HTTP ${res.status}`);
      const data = await res.json();
      setPages(
        ((data.value ?? []) as Array<{ FileLeafRef?: string; Title?: string; FileRef?: string }>)
          .filter((p) => (p.FileLeafRef ?? "").toLowerCase().indexOf(".aspx") !== -1)
          .map((p) => ({
            fileName: p.FileLeafRef ?? "",
            title: p.Title ?? "",
            serverRelativeUrl: p.FileRef ?? "",
          })),
      );
      setError(undefined);
    };
    load().catch((e) => {
      // Empty ≠ unknown, as everywhere else here. An unreadable list must not render as "none of
      // these pages exist", which would tell the admin to create ten pages that already do.
      setPages(undefined);
      setError((e as Error).message);
    });
  }, [props.siteUrl]);

  // Overrides still apply when the read failed, so a fully-overridden page keeps working.
  const targets = resolveAll(pages ?? [], props.overrides);

  return (
    <section style={s.wrap}>
      <div style={s.head}>
        <h1 style={s.h1}>{props.heading || "CRS Settings"}</h1>
        <p style={s.sub}>
          {props.subheading || "Manage repository settings, folders, access, and mappings."}
        </p>
      </div>

      {error !== undefined && (
        <div style={s.danger}>
          Could not read this site&rsquo;s pages ({error}), so the links below could not be worked
          out. This does <strong>not</strong> mean those pages are missing. Reload the page, or set the
          addresses by hand in this web part&rsquo;s settings.
        </div>
      )}

      {pages === undefined && error === undefined ? (
        <div style={s.loading}>Loading&hellip;</div>
      ) : (
        <div style={s.grid}>
          {([1, 2] as Array<1 | 2>).map((col) => (
            <div key={col} style={s.col}>
              {columnCards(col).map((c) => (
                <Card key={c.key} card={c} targets={targets} />
              ))}
            </div>
          ))}
        </div>
      )}
    </section>
  );
}

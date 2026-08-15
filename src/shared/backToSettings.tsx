/**
 * The band that takes an admin back to CRS Settings, and the shared look of every back band.
 *
 * Added 2026-08-15: CRS Settings is a directory, so every page it links to is a dead end without a
 * way back — the client hit exactly that on Folder Management ("I can't go back to CRS Settings").
 *
 * The destination is RESOLVED FROM SITE PAGES, never hardcoded, for the same reason the landing page
 * resolves its own links: the client renames every DMS-named artefact at import (memory
 * `dms-to-crs-rename-pending`), and a hardcoded `CRS-Settings.aspx` fails as a dead link — no error,
 * no clue. `resolveLink` is reused rather than re-written so the matching rule keeps one home.
 *
 * When the page cannot be found the band still renders, but as a plain "Back" driven by browser
 * history: the label must not promise a destination this code could not locate. A band that vanished
 * whenever the Site Pages read failed would leave the admin in the dead end they reported, which is
 * the one outcome worth avoiding here.
 */

import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { AdminLink, SitePage, resolveLink } from "./adminPages";

/**
 * How CRS Settings is recognised.
 *
 * `crs.?settings` covers both `CRS-Settings.aspx` and the title `CRS Settings`; the anchored
 * `^settings` catches a site that dropped the prefix. NOT a bare `/settings/i` — that would claim a
 * `Site Settings` or `List Settings` page and send the admin somewhere nobody meant, the same
 * collision the landing page's patterns are pinned against.
 */
export const SETTINGS_LINK: AdminLink = {
  key: "settings",
  label: "CRS Settings",
  match: /crs.?settings|^settings\b/i,
  pageName: "CRS Settings",
};

/**
 * The band's look, exported so every back affordance in the product is literally the same object.
 *
 * Originally ApprovalDocument.tsx:88-89; it lives here now that more than one screen wants it. Only
 * the TEXT is clickable, not the band — a full-width click target for "go back" loses your place on
 * a stray click near the margin.
 */
export const backBandStyle: React.CSSProperties = {
  background: "rgba(0, 104, 74, 0.08)",
  borderRadius: 4,
  padding: "10px 16px",
  marginBottom: 20,
};

export const backLinkStyle: React.CSSProperties = {
  border: "none",
  background: "transparent",
  padding: 0,
  font: "inherit",
  color: "rgba(0, 104, 74, 1)",
  fontSize: 14,
  fontWeight: 600,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  textDecoration: "none",
};

export const backChevronStyle: React.CSSProperties = {
  fontSize: 16,
  lineHeight: 1,
  fontWeight: 700,
};

/**
 * A back band with an arbitrary label and handler — the shape every exit uses.
 *
 * `<` is written as an expression because JSX reads a literal left angle bracket in children as the
 * start of a tag.
 */
export function BackBand({
  label,
  onClick,
  href,
}: {
  label: string;
  onClick?: () => void;
  href?: string;
}): React.ReactElement {
  return (
    <div style={backBandStyle}>
      {href === undefined ? (
        <button type="button" style={backLinkStyle} onClick={onClick}>
          <span style={backChevronStyle}>{"<"}</span>
          {label}
        </button>
      ) : (
        <a href={href} style={backLinkStyle}>
          <span style={backChevronStyle}>{"<"}</span>
          {label}
        </a>
      )}
    </div>
  );
}

/**
 * Read the site's pages once, the same way CrsSettings does.
 *
 * `FileRef`, not a path assembled from the library title — a list's URL is independent of its title
 * (gotcha #12) and `SitePages` is not guaranteed to be the segment on every site.
 */
async function readSitePages(context: WebPartContext, siteUrl: string): Promise<SitePage[]> {
  const res: SPHttpClientResponse = await context.spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items` +
      `?$select=Id,FileLeafRef,Title,FileRef&$top=500`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) throw new Error(`Site Pages: HTTP ${res.status}`);
  const data = await res.json();
  return ((data.value ?? []) as Array<{ FileLeafRef?: string; Title?: string; FileRef?: string }>)
    .filter((p) => (p.FileLeafRef ?? "").toLowerCase().indexOf(".aspx") !== -1)
    .map((p) => ({
      fileName: p.FileLeafRef ?? "",
      title: p.Title ?? "",
      serverRelativeUrl: p.FileRef ?? "",
    }));
}

/**
 * "‹ Back to CRS Settings" — or a plain "‹ Back" while the page is unknown.
 *
 * `override` is the same escape hatch the landing page offers: on a site whose page names this code
 * cannot guess, an address typed into the property pane wins outright and is never pattern-checked.
 */
export function BackToSettings({
  context,
  siteUrl,
  override,
}: {
  context: WebPartContext;
  siteUrl: string;
  override?: string;
}): React.ReactElement {
  const [url, setUrl] = useState<string | undefined>(undefined);

  useEffect(() => {
    const manual = (override ?? "").trim();
    if (manual.length > 0) {
      setUrl(manual);
      return;
    }
    readSitePages(context, siteUrl)
      .then((pages) => {
        const target = resolveLink(SETTINGS_LINK, pages);
        // `missing` leaves `url` undefined, which renders the history fallback rather than a dead
        // arrow. `ambiguous` still navigates — a best guess beats stranding the admin.
        setUrl(target.state === "missing" ? undefined : target.url);
      })
      .catch(() => setUrl(undefined));
  }, [siteUrl, override]);

  if (url === undefined) {
    return (
      <BackBand
        label="Back"
        onClick={() => {
          if (typeof window !== "undefined") window.history.back();
        }}
      />
    );
  }
  return <BackBand label="Back to CRS Settings" href={url} />;
}

/**
 * Put the band above a web part's own element.
 *
 * Called from `render()` — the WEB PART boundary — and never from inside a component. That seam is
 * load-bearing: several of these components are also mounted as STEPS of a Folder Management guided
 * flow, and a band offering the way out of the flow, styled identically to the band that goes back one
 * step within it, would read as part of the flow. Wrapping at render() means an embedded mount cannot
 * inherit it, with no `embedded` prop for a caller to forget.
 *
 * Folder Administration is the deliberate exception and does NOT use this: its picker renders
 * `BackToSettings` itself, so the deeper views keep their own "Back to Folder Management" instead of
 * showing two bands.
 */
export function withBackToSettings(
  context: WebPartContext,
  element: React.ReactElement,
): React.ReactElement {
  return React.createElement(
    "div",
    undefined,
    React.createElement(BackToSettings, {
      key: "back",
      context,
      siteUrl: context.pageContext.web.absoluteUrl,
    }),
    element,
  );
}

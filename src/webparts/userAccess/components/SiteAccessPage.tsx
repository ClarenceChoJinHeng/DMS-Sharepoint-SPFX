import * as React from "react";
import AccessShell from "./AccessShell";
import { IAccessProps } from "./IAccessProps";
import { resolveLink, CARDS } from "../../../shared/adminPages";
import { readSitePages } from "../../../shared/backToSettings";

/**
 * RETIRED (2026-09-02), same treatment as `FolderAccessPage.tsx` and
 * `ApprovalLibraryAccessPage.tsx` — read `FolderAccessPage.tsx`'s comment for the full reasoning.
 *
 * Client: *"Approval Library Access, Site Access is not needed, its confusing them as they already
 * know that adding a user in the group from Group Management it will automatically allow them have
 * access to site and library. Just keep Page access."*
 *
 * ⚠ IT IS A SIGNPOST, NOT A DELETION. The web part stays registered and the `Site-Access.aspx` page
 * stays on the site — removing a component from `componentIds` is the change that has already
 * bitten this project twice (CRS Requests undeployable for weeks, the `+ New Folder` customizer
 * inert for months, both silently), and an admin may have this URL bookmarked or in site
 * navigation. `pageAccessPolicy` still locks it, so nothing is newly exposed. `SiteAccess.tsx`
 * itself is untouched — this page simply stops mounting it.
 */
export default function SiteAccessPage({ context }: IAccessProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const [href, setHref] = React.useState<string | undefined>(undefined);

  /**
   * The Group Management address is RESOLVED from Site Pages, never hardcoded — the client renames
   * every page at import, so a literal `Group-Management.aspx` would fail as a DEAD LINK on the one
   * page whose only job is saying where to go. Same rule, and the same `resolveLink`, as the other
   * two retired pages and the CRS Settings landing page.
   *
   * A failed read leaves the link out rather than guessing: the text below still names the page.
   */
  React.useEffect(() => {
    const link = CARDS
      .filter((c) => c.key === "access")[0]
      ?.links.filter((l) => l.key === "groups")[0];
    if (!link) return;
    readSitePages(context, siteUrl)
      .then((pages) => {
        const t = resolveLink(link, pages);
        if (t.state !== "missing") setHref(t.url);
      })
      .catch(() => undefined);
  }, []);

  return (
    <AccessShell
      title="Site Access has moved"
      subtitle="Who may open this site at all is now decided entirely by Group Management."
    >
      <div style={{ fontSize: 13, lineHeight: 1.6, maxWidth: 760 }}>
        <p>
          Adding a person to any group in <strong>CRS_SITE_MEMBERS</strong> or a unit&rsquo;s
          uploader/approver group already gives them everything they need to reach the site. There
          is nothing to configure separately.
        </p>
        <p>
          Nothing has changed about who can reach what. Access is still applied by a{" "}
          <strong>Folder Reconciliation</strong> run, from the Folder Administration page.
        </p>
        {href !== undefined
          ? (
            <p>
              <a href={href} style={{ fontWeight: 600 }}>Go to Group Management</a>
            </p>
          )
          : (
            <p style={{ color: "#666" }}>
              Open the <strong>Group Management</strong> page from CRS Settings.
            </p>
          )}
      </div>
    </AccessShell>
  );
}

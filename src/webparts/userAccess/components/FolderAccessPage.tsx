import * as React from "react";
import AccessShell from "./AccessShell";
import { IAccessProps } from "./IAccessProps";
import { resolveLink, CARDS } from "../../../shared/adminPages";
import { readSitePages } from "../../../shared/backToSettings";

/**
 * RETIRED (2026-08-23). This page did two jobs and both moved.
 *
 * Client: *"I think we can remove folder access and also in the step, reason being I separated Folder
 * Access to allow client to add the user separately but since the Group Management will be doing most
 * of the job of bulk group and allow client to add user on their own, that makes Folder Access
 * redundant."*
 *
 * - **Mapping a group** stopped being work an admin does: creating a group writes its Group Map rows,
 *   and Bulk provisioning writes every row a segment needs. The hand-mapping form left this page on
 *   2026-08-18 and left Group Management on 2026-08-23.
 * - **Membership** came here on 2026-08-18 and went back to Group Management today, where the group
 *   list itself now expands to add and remove people.
 *
 * \u26a0 IT IS A SIGNPOST, NOT A DELETION, and that is deliberate. The web part stays registered and the
 * `Folder-Access.aspx` page stays on the site, because:
 *   - Removing a component from `componentIds` is the change that has already bitten this project
 *     twice \u2014 CRS Requests was undeployable for weeks and the `+ New Folder` customizer was inert for
 *     months, both silently. Deploying a package that deletes a web part currently on a live page
 *     renders that page broken, with nothing saying why.
 *   - An admin may have the URL bookmarked, or the page in site navigation. A page that explains
 *     where the work went is worth more than one that errors.
 *   - `pageAccessPolicy` still locks this page, so nothing is newly exposed.
 * Delete the page and the registration together, after the migration, as their own reviewable change.
 */
export default function FolderAccessPage({ context }: IAccessProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const [href, setHref] = React.useState<string | undefined>(undefined);

  /**
   * The Group Management address is RESOLVED from Site Pages, never hardcoded \u2014 the client renames
   * every page at import, so a literal `Group-Management.aspx` would fail as a DEAD LINK on the one
   * page whose only job is saying where to go. Same rule, and the same `resolveLink`, as the CRS
   * Settings landing page.
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
        // `missing` leaves the link out; `ambiguous` still navigates, as it does everywhere else — a
        // best guess beats stranding the admin on a page that only exists to send them elsewhere.
        if (t.state !== "missing") setHref(t.url);
      })
      .catch(() => undefined);
  }, []);

  return (
    <AccessShell
      title="Folder Access has moved"
      subtitle="Everything this page did is now on Group Management."
    >
      <div style={{ fontSize: 13, lineHeight: 1.6, maxWidth: 760 }}>
        <p>
          <strong>Mapping a group to a folder</strong> is no longer a separate step. Creating a group
          writes its folder mappings, and <em>Create all groups for a segment</em> writes every
          mapping a segment needs.
        </p>
        <p>
          <strong>Adding and removing people</strong> is in the group list on Group Management \u2014
          expand any group to see who is in it.
        </p>
        <p>
          Nothing has changed about who can reach what. Folder permissions are still applied by a{" "}
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

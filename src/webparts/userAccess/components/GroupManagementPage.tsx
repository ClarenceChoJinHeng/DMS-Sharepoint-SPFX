import * as React from "react";
import { useState } from "react";
import AccessShell from "./AccessShell";
import GroupManager from "./GroupManager";
import BulkGroupProvisioner from "./BulkGroupProvisioner";
import { IAccessProps } from "./IAccessProps";

/**
 * ⚠ THE "Create one group" MODE SWITCH IS GONE (2026-09-02, client: "remove the Create one group,
 * since we are automatically creating the group for them"). Bulk provisioning is now the ONLY path
 * on this page — it writes every group a segment needs AND their folder mappings in one run, which
 * is what made the manual form redundant for the normal case.
 *
 * `GroupManager`'s create-one-group FORM is not deleted, only permanently hidden here
 * (`hideCreateForm={true}`, unconditionally) — the same "kept, not deleted" pattern this project
 * already uses for `GroupMapBuilder`'s form (`show="form"`, unreachable since 2026-08-23) and the
 * old folder-tree code. The form existed for the site-entry group and genuine one-offs bulk
 * provisioning cannot express; if that need resurfaces, re-enabling is passing `false` here again,
 * not rebuilding a form. Do not remove `GroupManager`'s underlying create logic on the strength of
 * this page alone.
 */
export default function GroupManagementPage({ context }: IAccessProps): React.ReactElement {
  // The group list reads its mapping badges once, so a bulk run below leaves every group reading
  // `not mapped` until this bumps. The standalone page needs it as much as the guided flow did.
  const [refreshKey, setRefreshKey] = useState(0);

  return (
    <AccessShell
      title="Group Management"
      subtitle="Create the SharePoint groups a segment needs, see what any person can reach, and find the groups that grant nothing. Creating a group also writes its folder mappings."
    >
      {/* ⚠ ORDER SWAPPED (client's mockup, 2026-09-03: "Move this to be below the section, after
          Group on this site"). Quick Search and the group list now come FIRST, Create group last —
          the create form was moved lower, not the lookup/list moved higher. */}
      <GroupManager
        context={context}
        siteUrl={context.pageContext.web.absoluteUrl}
        hideCreateForm={true}
        refreshKey={refreshKey}
      />

      <BulkGroupProvisioner
        context={context}
        siteUrl={context.pageContext.web.absoluteUrl}
        onRunComplete={() => setRefreshKey((n) => n + 1)}
      />
    </AccessShell>
  );
}

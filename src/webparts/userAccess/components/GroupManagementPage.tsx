import * as React from "react";
import { useState } from "react";
import AccessShell from "./AccessShell";
import GroupManager from "./GroupManager";
import BulkGroupProvisioner from "./BulkGroupProvisioner";
import { IAccessProps } from "./IAccessProps";

export default function GroupManagementPage({ context }: IAccessProps): React.ReactElement {
  // The group list reads its mapping badges once, so a bulk run below leaves every group reading
  // `not mapped` until this bumps. The standalone page needs it as much as the guided flow does.
  const [refreshKey, setRefreshKey] = useState(0);
  return (
    <AccessShell
      title="Group Management"
      subtitle="Create SharePoint groups and manage who is in them. Which folders a group can reach is set separately, on Folder Access."
    >
      <GroupManager
        context={context}
        siteUrl={context.pageContext.web.absoluteUrl}
        refreshKey={refreshKey}
      />
      {/* BELOW the single-group form, not above it and not on its own page. The two share every rule — the
          persona, the derived name, the Group Map rows — so splitting them would be two mount points for one
          model. Order matters the other way round too: an admin who has just created one group by hand is
          exactly the person who should see that there is a button for all of them. */}
      <BulkGroupProvisioner
        context={context}
        siteUrl={context.pageContext.web.absoluteUrl}
        onRunComplete={() => setRefreshKey((n) => n + 1)}
      />
    </AccessShell>
  );
}

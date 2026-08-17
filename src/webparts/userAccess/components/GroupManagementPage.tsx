import * as React from "react";
import AccessShell from "./AccessShell";
import GroupManager from "./GroupManager";
import BulkGroupProvisioner from "./BulkGroupProvisioner";
import { IAccessProps } from "./IAccessProps";

export default function GroupManagementPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Group Management"
      subtitle="Create SharePoint groups and manage who is in them. Which folders a group can reach is set separately, on Folder Access."
    >
      <GroupManager context={context} siteUrl={context.pageContext.web.absoluteUrl} />
      {/* BELOW the single-group form, not above it and not on its own page. The two share every rule — the
          persona, the derived name, the Group Map rows — so splitting them would be two mount points for one
          model. Order matters the other way round too: an admin who has just created one group by hand is
          exactly the person who should see that there is a button for all of them. */}
      <BulkGroupProvisioner context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}

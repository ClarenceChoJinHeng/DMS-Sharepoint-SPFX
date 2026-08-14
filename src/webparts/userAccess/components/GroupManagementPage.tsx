import * as React from "react";
import AccessShell from "./AccessShell";
import GroupManager from "./GroupManager";
import { IAccessProps } from "./IAccessProps";

export default function GroupManagementPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Group Management"
      subtitle="Create SharePoint groups and manage who is in them. Which folders a group can reach is set separately, on Folder Access."
    >
      <GroupManager context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}

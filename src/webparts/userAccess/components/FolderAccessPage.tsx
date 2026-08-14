import * as React from "react";
import AccessShell from "./AccessShell";
import GroupMapBuilder from "./GroupMapBuilder";
import { IAccessProps } from "./IAccessProps";

export default function FolderAccessPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Folder Access"
      subtitle="Map an existing SharePoint group to a segment, tier and role — this is what controls who can reach each folder."
      /* This note belongs on THIS page and no other. Folder Access writes Group Map ROWS;
         the folder ACLs are applied later by reconciliation. Site, Library and Page Access
         all call addroleassignment directly and take effect immediately — verified
         2026-08-07 — so repeating this warning there would be false.

         The membership half of the old note left with membership itself (2026-08-14): telling
         an admin that "adding members takes effect immediately" on a page that can no longer
         add members sends them looking for a control that is not here. */
      note={
        <>
          Groups are created, and their members managed, on the{" "}
          <strong>Group Management</strong> page — those changes take effect immediately. Adding
          or deleting a <strong>mapping</strong> here only changes folder permissions after a{" "}
          <strong>Folder Reconciliation</strong> run — the fourth tab of the{" "}
          <strong>Folder Administration</strong> web part.
        </>
      }
    >
      <GroupMapBuilder context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}

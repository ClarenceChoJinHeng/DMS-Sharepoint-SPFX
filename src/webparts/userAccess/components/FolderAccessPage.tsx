import * as React from "react";
import AccessShell from "./AccessShell";
import GroupMapBuilder from "./GroupMapBuilder";
import { IAccessProps } from "./IAccessProps";

export default function FolderAccessPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Folder Access"
      subtitle="Map a SharePoint group to a segment, tier and role — this is what controls who can reach each folder."
      /* This note belongs on THIS page and no other. Folder Access writes Group Map ROWS;
         the folder ACLs are applied later by reconciliation. Site, Library and Page Access
         all call addroleassignment directly and take effect immediately — verified
         2026-08-07 — so repeating this warning there would be false. */
      note={
        <>
          Adding or removing <strong>members</strong> of a group takes effect immediately.
          Creating or deleting a <strong>mapping</strong> only changes folder permissions after
          a <strong>Folder Reconciliation</strong> run — the fourth tab of the{" "}
          <strong>Folder Administration</strong> web part.
        </>
      }
    >
      <GroupMapBuilder context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}

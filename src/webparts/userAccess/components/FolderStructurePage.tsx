import * as React from "react";
import AccessShell from "./AccessShell";
import StructureManager from "./StructureManager";
import { IAccessProps } from "./IAccessProps";

export default function FolderStructurePage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Folder Structure"
      subtitle="Add or reorder the folder levels beneath Unit for a business segment."
      /* The note names the one consequence an admin cannot undo from this page.
         Everything else here is reversible by editing the structure again; moving
         documents that are already filed is not, and needs a separate migration. */
      note={
        <>
          Changes apply to <strong>new uploads only</strong>. Documents already filed stay where
          they are, so adding a level to a segment that is already in use leaves two folder
          shapes side by side until someone moves them.
        </>
      }
    >
      <StructureManager context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}

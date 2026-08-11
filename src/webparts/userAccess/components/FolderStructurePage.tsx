import * as React from "react";
import { useState } from "react";
import AccessShell from "./AccessShell";
import StructureManager from "./StructureManager";
import SubtreeMigrator from "./SubtreeMigrator";
import { IAccessProps } from "./IAccessProps";

/**
 * Two tabs, deliberately on ONE page: changing the structure and migrating what is already
 * filed are two halves of one job. Separate web parts would let an admin finish the first, see
 * a success message, and never learn the second exists — which is the state that leaves a unit
 * with two folder shapes indefinitely.
 */
const tab = (active: boolean): React.CSSProperties => ({
  background: "none",
  border: "none",
  borderBottom: active ? "2px solid #0f6c3f" : "2px solid transparent",
  color: active ? "#0f6c3f" : "#605e5c",
  fontSize: 13,
  fontWeight: active ? 700 : 400,
  padding: "8px 2px",
  marginRight: 20,
  cursor: "pointer",
});

export default function FolderStructurePage({ context }: IAccessProps): React.ReactElement {
  const [view, setView] = useState<"structure" | "migrate">("structure");

  return (
    <AccessShell
      title="Folder Structure"
      subtitle="Add or reorder the folder levels beneath Unit, and move documents already filed into the new shape."
      /* The note names the one consequence an admin cannot undo from the first tab, and points
         at the second tab, which is where it gets resolved. */
      note={
        <>
          Changing the structure applies to <strong>new uploads only</strong>. Documents already
          filed stay where they are, so adding a level to a segment that is already in use leaves
          two folder shapes side by side — use <strong>Move existing folders</strong> to bring the
          old ones into line.
        </>
      }
    >
      <div style={{ borderBottom: "1px solid #edebe9", marginBottom: 20 }}>
        <button style={tab(view === "structure")} onClick={() => setView("structure")}>
          Folder levels
        </button>
        <button style={tab(view === "migrate")} onClick={() => setView("migrate")}>
          Move existing folders
        </button>
      </div>

      {view === "structure" ? (
        <StructureManager context={context} siteUrl={context.pageContext.web.absoluteUrl} />
      ) : (
        <SubtreeMigrator context={context} siteUrl={context.pageContext.web.absoluteUrl} />
      )}
    </AccessShell>
  );
}

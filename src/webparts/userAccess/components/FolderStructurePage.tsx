import * as React from "react";
import { useState } from "react";
import AccessShell from "./AccessShell";
import StructureManager from "./StructureManager";
import SubtreeMigrator from "./SubtreeMigrator";
import SegmentCreator from "./SegmentCreator";
import { IAccessProps } from "./IAccessProps";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

/**
 * Three tabs, deliberately on ONE page: changing the structure and migrating what is already
 * filed are two halves of one job. Separate web parts would let an admin finish the first, see
 * a success message, and never learn the second exists — which is the state that leaves a unit
 * with two folder shapes indefinitely.
 *
 * "Segments" joins them rather than taking its own page for the same reason: what it creates is
 * the same `Levels` chain the first tab edits, and the levels it names are the ones the second tab
 * would have to move. One page keeps the whole of a segment's shape in one place.
 *
 * That tab was "New segment" until 2026-08-14, when it gained the ability to DELETE one (spec
 * `2026-08-14-delete-segment-design.md`). Deletion belongs beside creation rather than on "Folder
 * levels", where a Delete button under a list of levels reads as "delete this level" — the most
 * expensive misreading available on that screen, since removing a level triggers a migration.
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
  type View = "structure" | "migrate" | "new";
  const [view, setView] = useState<View>("structure");
  /** True while the structure editor holds unsaved changes — see the tab guard below. */
  const [dirty, setDirty] = useState(false);
  const [blocked, setBlocked] = useState(false);

  /**
   * Switching tabs unmounts the editor, which silently discards whatever was being edited. So
   * while it is dirty the switch is REFUSED rather than confirmed: the Save button is a few
   * pixels away, and offering "discard" here would put losing the work one click behind
   * something that looks like ordinary navigation.
   */
  const go = (next: View): void => {
    if (next !== view && dirty) {
      setBlocked(true);
      return;
    }
    setBlocked(false);
    setView(next);
  };

  return (
    <AccessShell
      title="Folder Structure"
      subtitle="Add a business segment, change the folder levels beneath Unit, and move documents already filed into the new shape."
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
        <button style={tab(view === "structure")} onClick={() => go("structure")}>
          Folder levels
        </button>
        <button style={tab(view === "migrate")} onClick={() => go("migrate")}>
          Move existing folders
        </button>
        <button style={tab(view === "new")} onClick={() => go("new")}>
          Segments
        </button>
      </div>

      {blocked && (
        <div
          style={{
            fontSize: 13,
            padding: "10px 12px",
            borderRadius: 6,
            marginBottom: 16,
            lineHeight: 1.5,
            ...NOTICE_ATTENTION,
          }}
        >
          Finish or clear what you are editing first — leaving this tab would lose it.
        </div>
      )}

      {view === "structure" && (
        <StructureManager
          context={context}
          siteUrl={context.pageContext.web.absoluteUrl}
          onDirtyChange={(d) => {
            setDirty(d);
            // Clear the refusal as soon as the reason for it is gone, so a saved edit does not
            // leave a warning sitting above the page telling them to do what they just did.
            if (!d) setBlocked(false);
          }}
        />
      )}
      {view === "migrate" && (
        <SubtreeMigrator context={context} siteUrl={context.pageContext.web.absoluteUrl} />
      )}
      {view === "new" && (
        // Shares the same dirty guard: a half-typed segment costs more to retype than a level
        // edit, and losing it to a tab click would be the same silent discard.
        <SegmentCreator
          context={context}
          siteUrl={context.pageContext.web.absoluteUrl}
          onDirtyChange={(d) => {
            setDirty(d);
            if (!d) setBlocked(false);
          }}
        />
      )}
    </AccessShell>
  );
}

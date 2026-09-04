// The read-only member-count half of the Approval Library Access and Page Access screens.
//
// Spec: docs/superpowers/specs/2026-08-14-per-person-access-removal-design.md (superseded —
// see docs/superpowers/specs/2026-09-02-access-pages-read-only-design.md). The per-person
// REMOVAL feature this file used to hold (`MemberRemovalDialog`, `MemberRows`, `removalFor`,
// `doRemoveMember`) was retired 2026-09-02: both screens that mounted it stopped, because
// removing a person is now done on Group Management, and reconciliation ends access to
// everything a group grants the moment membership changes there — the "did this actually end
// their access" check that dialog existed for has no meaning once removal itself moved.
//
// What is left is display-only: a member COUNT per group, shown collapsed. Kept because both
// screens still show groups and still want to say how many people are in one, without a second
// definition of what "loading" / "no members" / "could not be read" mean on a count.
import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { SpGroup } from "../../../shared/spGroupsFilter";
import { getGroupMembers } from "../../../shared/spGroups";
import { MemberMap, loadOf, memberCountLabel, memberLabel, sortMembers } from "../../../shared/accessMembers";

const m: Record<string, React.CSSProperties> = {
  toggle: { border: "none", background: "transparent", padding: 0, color: "#0f6c3f", fontSize: 12, cursor: "pointer", textAlign: "left" },
  chev:   { display: "inline-block", width: 12, color: "#666" },
  amber:  { color: "#b45309" },
  modalBg:  { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" },
  modal:    { background: "#fff", borderRadius: 6, padding: 20, width: "min(420px, 92vw)", maxHeight: "76vh", overflowY: "auto", boxShadow: "0 8px 32px rgba(0,0,0,.25)" },
  head:     { fontWeight: 600, fontSize: 15, margin: "0 0 12px" },
  no:       { color: "#8a8886", fontSize: 12 },
  dangerBox:{ padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c" },
  memberList: { margin: 0, paddingLeft: 18, fontSize: 13, lineHeight: 1.6 },
  closeBtn: { padding: "6px 14px", border: "1px solid #c8c6c4", borderRadius: 2, background: "#fff", fontSize: 13, cursor: "pointer" },
};

/** Bounded below the point this codebase's own earlier comment on this hook named as unsafe: "ten
 * simultaneous calls is how a tenant starts returning 429s". 6 stays under that with room to spare. */
const CONCURRENCY = 6;

/**
 * Fetch and hold members for a set of groups.
 *
 * ⚠ WAS (twice): first, one `getGroupMembers` request PER group, fully SEQUENTIAL — an 8-minute load
 * on ~700 groups (Site Access, 2026-09-02). Then ONE `fetchAllGroupMembers` request
 * (`sitegroups?$expand=Users`) covering every group at once — which traded the request-COUNT problem
 * for a request-SIZE problem: SharePoint still took 1-2 minutes to compute and return that one huge
 * response, confirmed live on the deployed fix. Neither extreme works.
 *
 * FIXED: bounded-concurrency parallel reads, `CONCURRENCY` groups' `getGroupMembers` in flight at
 * once, chunk by chunk — fast (most of 700 requests' wall-clock time is now hidden behind the ones
 * running alongside it), safe (well under the 429 threshold), and rows fill in progressively per
 * chunk rather than the whole list sitting on "loading…" until everything finishes at once.
 *
 * A FAILED read is recorded as an error rather than as []. "No members" on a group that merely
 * failed to load is what stops an admin trusting the count at all.
 *
 * `refresh` re-reads ONE group — kept as a single-group call, since refreshing on demand for one
 * row is a rare interactive action, not a mount-time load of everything.
 */
export function useGroupMembers(
  sp: SPHttpClient,
  siteUrl: string,
  groupIds: readonly number[],
): { members: MemberMap; refresh: (groupId: number) => Promise<void> } {
  const [members, setMembers] = useState<MemberMap>({});
  // Joined, so the effect re-runs when the SET changes rather than on every render. Sorted first:
  // the same group ids in a different order are the same set of work.
  const key = groupIds.slice().sort((a, b) => a - b).join(",");

  const fetchOne = async (gid: number): Promise<{ gid: number; load: MemberMap[number] }> => {
    try {
      const list = await getGroupMembers(sp, siteUrl, gid);
      return { gid, load: { state: "loaded", members: list } };
    } catch (e) {
      return { gid, load: { state: "error", message: (e as Error).message } };
    }
  };

  useEffect(() => {
    let cancelled = false;
    const ids = key.length > 0 ? key.split(",").map(Number) : [];
    if (ids.length === 0) return undefined;
    setMembers((prev) => {
      const next = { ...prev };
      for (const gid of ids) if (!next[gid]) next[gid] = { state: "loading" };
      return next;
    });
    const run = async (): Promise<void> => {
      for (let i = 0; i < ids.length; i += CONCURRENCY) {
        if (cancelled) return;
        const chunk = ids.slice(i, i + CONCURRENCY);
        const results = await Promise.all(chunk.map(fetchOne));
        if (cancelled) return;
        setMembers((prev) => {
          const next = { ...prev };
          for (const r of results) next[r.gid] = r.load;
          return next;
        });
      }
    };
    run().catch(() => undefined);
    return () => { cancelled = true; };
  }, [key, siteUrl]);

  const refresh = async (gid: number): Promise<void> => {
    setMembers((prev) => ({ ...prev, [gid]: { state: "loading" } }));
    try {
      const list = await getGroupMembers(sp, siteUrl, gid);
      setMembers((prev) => ({ ...prev, [gid]: { state: "loaded", members: list } }));
    } catch (e) {
      setMembers((prev) => ({ ...prev, [gid]: { state: "error", message: (e as Error).message } }));
    }
  };

  return { members, refresh };
}

/** The collapsed summary shown in a group row's Members cell. */
export function MemberSummary(props: {
  group: SpGroup;
  members: MemberMap;
  expanded: boolean;
  onToggle: () => void;
  /** False for a group that grants nothing here — the count still helps, the chevron does not. */
  expandable: boolean;
}): React.ReactElement {
  const load = loadOf(props.members, props.group.id);
  const label = memberCountLabel(load);
  // Amber for a group that is granted but empty: it reads as configured and grants nobody
  // anything, which is the one state worth interrupting a scan for.
  const style =
    load.state === "error" ? { color: "#a4262c" } :
    load.state === "loaded" && load.members.length === 0 ? m.amber : undefined;

  if (!props.expandable) return <span style={style}>{label}</span>;
  return (
    <button
      style={m.toggle}
      onClick={props.onToggle}
      title={props.expanded ? "Hide members" : "Show members"}
    >
      <span style={m.chev}>{props.expanded ? "▾" : "▸"}</span>
      <span style={style}>{label}</span>
    </button>
  );
}

/**
 * `MemberSummary` (expandable) plus the popup it opens into — the whole "click a count, see who"
 * pattern in ONE component, managing its own open/closed state, so a caller needs nothing beyond
 * the group and its already-fetched `members` map.
 *
 * First built for Page Access (2026-09-02) and extracted here so Approval Library Access and Site
 * Access reuse the identical implementation rather than a second and third copy of a modal — this
 * codebase has repeatedly hit "two lists of the same people drift" when a member-list pattern was
 * duplicated instead of shared.
 */
export function MemberCountWithPopup(props: {
  group: SpGroup;
  members: MemberMap;
}): React.ReactElement {
  const [open, setOpen] = useState(false);
  const load = loadOf(props.members, props.group.id);
  return (
    <>
      <MemberSummary
        group={props.group}
        members={props.members}
        expanded={open}
        expandable={true}
        onToggle={() => setOpen(true)}
      />
      {open && (
        <div style={m.modalBg} onClick={() => setOpen(false)}>
          <div style={m.modal} onClick={(e) => e.stopPropagation()}>
            <p style={m.head}>{props.group.title}</p>
            {load.state === "loading" && <div style={m.no}>Loading members&hellip;</div>}
            {load.state === "error" && (
              <div style={m.dangerBox}>Could not read this group&rsquo;s members: {load.message}</div>
            )}
            {load.state === "loaded" && (
              sortMembers(load.members).length === 0 ? (
                <div style={m.no}>No members.</div>
              ) : (
                <ul style={m.memberList}>
                  {sortMembers(load.members).map((mem) => (
                    <li key={mem.id}>{memberLabel(mem)}</li>
                  ))}
                </ul>
              )
            )}
            <div style={{ marginTop: 16, textAlign: "right" }}>
              <button style={m.closeBtn} onClick={() => setOpen(false)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

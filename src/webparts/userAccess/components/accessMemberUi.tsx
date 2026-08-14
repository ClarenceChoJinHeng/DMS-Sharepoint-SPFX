// The per-person half of the Approval Library Access and Page Access screens.
//
// Spec: docs/superpowers/specs/2026-08-14-per-person-access-removal-design.md
//
// ONE copy, mounted by both screens. The rules it enforces are the ones an admin cannot check for
// themselves — whether a removal actually ends someone's access, and whether the list it was
// decided from was complete — so two copies would eventually disagree about a permission change,
// and each screen would still be individually convincing.
//
// It renders INTO an existing table: the member rows are <tr>s with a colSpan cell, so they sit
// under their group and inherit its column layout. A nested table would misalign the moment either
// screen adds a column, and the two screens do not have the same columns.
import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { SpGroup, SpGroupMember } from "../../../shared/spGroupsFilter";
import { getGroupMembers, removeGroupMember } from "../../../shared/spGroups";
import {
  MemberMap,
  MemberRemoval,
  buildRemoval,
  loadOf,
  memberCountLabel,
  memberLabel,
  membershipIndex,
  removalVerdict,
  sortMembers,
} from "../../../shared/accessMembers";

const m: Record<string, React.CSSProperties> = {
  subCell:  { padding: 0, borderBottom: "1px solid #f0f0f0", background: "#fbfbfb" },
  subWrap:  { padding: "8px 10px 10px 36px" },
  memberRow:{ display: "flex", alignItems: "center", gap: 10, padding: "4px 0", flexWrap: "wrap" },
  name:     { fontWeight: 600, fontSize: 12 },
  email:    { color: "#666", fontSize: 11, userSelect: "text" },
  alsoIn:   { fontSize: 11, color: "#8a4b00" },
  spacer:   { flex: 1, minWidth: 12 },
  del:      { padding: "2px 10px", fontSize: 11, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  delOff:   { padding: "2px 10px", fontSize: 11, background: "#e6e6e6", color: "#999", border: "1px solid #e0e0e0", borderRadius: 4, cursor: "not-allowed" },
  toggle:   { border: "none", background: "transparent", padding: 0, color: "#0f6c3f", fontSize: 12, cursor: "pointer", textAlign: "left" },
  chev:     { display: "inline-block", width: 12, color: "#666" },
  note:     { fontSize: 11, color: "#666", marginTop: 6, lineHeight: 1.45 },
  errNote:  { fontSize: 11, color: "#a4262c", marginTop: 4, lineHeight: 1.45 },
  amber:    { color: "#b45309" },
  overlay:  { position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 120, padding: 16 },
  box:      { background: "#fff", borderRadius: 8, width: "100%", maxWidth: 520, boxShadow: "0 8px 30px rgba(0,0,0,.25)" },
  head:     { padding: "12px 16px", borderBottom: "1px solid #eee", fontWeight: 600, fontSize: 14 },
  body:     { padding: "12px 16px", fontSize: 13, lineHeight: 1.55 },
  foot:     { padding: "10px 16px", borderTop: "1px solid #eee", display: "flex", justifyContent: "flex-end", gap: 8 },
  ghost:    { padding: "5px 14px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  danger:   { padding: "5px 14px", fontSize: 12, color: "#fff", background: "#a4262c", border: "none", borderRadius: 4, cursor: "pointer" },
  dangerOff:{ padding: "5px 14px", fontSize: 12, color: "#fff", background: "#c7c7c7", border: "none", borderRadius: 4, cursor: "not-allowed" },
  warnBlock:{ marginTop: 10, padding: "8px 10px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5 },
  infoBlock:{ marginTop: 10, padding: "8px 10px", border: "1px solid #cfd8e3", background: "#f4f7fb", borderRadius: 4, fontSize: 12, color: "#2b3f56", lineHeight: 1.5 },
};

/**
 * Fetch and hold members for a set of groups.
 *
 * SEQUENTIAL, not Promise.all — the reason PageAccess's original version gave still applies: ten
 * simultaneous calls is how a tenant starts returning 429s, and rows filling in visibly also tells
 * the admin it is still working.
 *
 * A FAILED read is now recorded as an error rather than as []. That was acceptable while this was a
 * decoration column; it is not once a removal is decided from it, because "no members" on a group
 * that merely failed to load is what stops an admin opening the row holding the person they came
 * for.
 *
 * `refresh` re-reads ONE group, which is what a removal needs. Reloading everything would re-fetch
 * every other group to show one row disappearing.
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

  const fetchOne = async (gid: number): Promise<void> => {
    try {
      const list = await getGroupMembers(sp, siteUrl, gid);
      setMembers((prev) => ({ ...prev, [gid]: { state: "loaded", members: list } }));
    } catch (e) {
      setMembers((prev) => ({ ...prev, [gid]: { state: "error", message: (e as Error).message } }));
    }
  };

  useEffect(() => {
    let cancelled = false;
    // Read from a local copy rather than from `members`: the loop outlives the render it started
    // in, so the closed-over state is stale by the second iteration and every group would look
    // unfetched. The map only ever gains keys, so tracking them here cannot drop one.
    const answered: Record<number, true> = {};
    setMembers((prev) => {
      for (const k of Object.keys(prev)) if (prev[Number(k)]) answered[Number(k)] = true;
      return prev;
    });
    const run = async (): Promise<void> => {
      for (const gid of key.length > 0 ? key.split(",").map(Number) : []) {
        if (cancelled) return;
        // Skip anything already answered, including a previous ERROR: retrying it on every set
        // change would hammer a group failing for a durable reason (deleted, no rights). The
        // explicit "Try again" on the expanded row is how it is asked for again.
        if (answered[gid]) continue;
        answered[gid] = true;
        setMembers((prev) => (prev[gid] ? prev : { ...prev, [gid]: { state: "loading" } }));
        await fetchOne(gid);
      }
    };
    run().catch(() => undefined);
    return () => { cancelled = true; };
  }, [key, siteUrl]);

  const refresh = async (gid: number): Promise<void> => {
    setMembers((prev) => ({ ...prev, [gid]: { state: "loading" } }));
    await fetchOne(gid);
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
 * The expanded member rows for one group, as a single <tr>.
 *
 * Every name carries the OTHER allowed groups it is also in. Put on the row rather than saved for
 * the dialog because it changes which row an admin clicks: removing someone from `..._UPL` when
 * they are also in `..._APR` is not the removal that was meant, and finding that out inside the
 * confirm dialog is finding it out after deciding.
 */
export function MemberRows(props: {
  group: SpGroup;
  members: MemberMap;
  allowedGroupIds: readonly number[];
  titleOf: (groupId: number) => string;
  colSpan: number;
  busy: boolean;
  /** User ids with a removal in flight. */
  pending: readonly number[];
  onRemove: (user: SpGroupMember) => void;
  onRetry: () => void;
  /** Shown when access does not currently depend on membership — an unrestricted page. */
  inertNote?: string;
}): React.ReactElement {
  const load = loadOf(props.members, props.group.id);
  const index = membershipIndex(props.allowedGroupIds, props.members);

  return (
    <tr>
      <td style={m.subCell} colSpan={props.colSpan}>
        <div style={m.subWrap}>
          {load.state === "loading" && (
            <span style={{ color: "#8a8886", fontSize: 12 }}>Loading members&hellip;</span>
          )}

          {load.state === "error" && (
            <div style={m.errNote}>
              Could not read the members of <strong>{props.group.title}</strong> ({load.message}).
              This is <strong>not</strong> the same as the group being empty — there may well be
              people in it.{" "}
              <button style={m.toggle} onClick={props.onRetry}>Try again</button>
            </div>
          )}

          {load.state === "loaded" && load.members.length === 0 && (
            <div style={{ ...m.note, ...m.amber }}>
              This group has no members, so it grants nobody access. Removing the group itself is
              what stops it appearing here.
            </div>
          )}

          {load.state === "loaded" && sortMembers(load.members).map((u) => {
            const also = (index.get(u.id) ?? [])
              .filter((gid) => gid !== props.group.id)
              .map((gid) => props.titleOf(gid))
              .filter((t) => t.trim().length > 0);
            const inFlight = props.pending.indexOf(u.id) !== -1;
            return (
              <div key={u.id} style={m.memberRow}>
                <span style={m.name}>{memberLabel(u)}</span>
                {u.email && <span style={m.email}>{u.email}</span>}
                {/* Named, not counted. "also in 1 group" cannot be checked against what the admin
                    intended; the group's name can. */}
                {also.length > 0 && (
                  <span style={m.alsoIn} title="Removing them here leaves that access in place">
                    also in {also.join(", ")}
                  </span>
                )}
                <span style={m.spacer} />
                <button
                  style={props.busy || inFlight ? m.delOff : m.del}
                  disabled={props.busy || inFlight}
                  onClick={() => props.onRemove(u)}
                >
                  {inFlight ? "Removing…" : "Remove person"}
                </button>
              </div>
            );
          })}

          {props.inertNote && <div style={m.note}>{props.inertNote}</div>}
        </div>
      </td>
    </tr>
  );
}

/**
 * Confirm one person's removal.
 *
 * Leads with whether this actually ends their access, in three states rather than two — "ends",
 * "does not end", "cannot tell". Flattening the third is how a screen makes a promise it has not
 * checked, and here the unchecked half is invisible: a group whose members failed to load renders
 * as a collapsed row like any other.
 *
 * `scopeWarning` belongs to the caller: the wider effect differs per screen and neither version is
 * derivable here.
 */
export function MemberRemovalDialog(props: {
  removal: MemberRemoval;
  /** The wider consequence, in the caller's own words. */
  scopeWarning: React.ReactNode;
  /** Shown when membership is not currently what grants access — an unrestricted page. */
  inertWarning?: React.ReactNode;
  busy: boolean;
  onCancel: () => void;
  onConfirm: () => void;
}): React.ReactElement {
  const r = props.removal;
  const verdict = removalVerdict(r);
  const who = memberLabel(r.user);

  return (
    <div style={m.overlay} onClick={props.onCancel}>
      <div style={m.box} onClick={(e) => e.stopPropagation()}>
        <div style={m.head}>Remove {who} from {r.groupName}?</div>
        <div style={m.body}>
          {verdict === "ends" && (
            <>
              <strong>{who}</strong> will be taken out of <strong>{r.groupName}</strong>, and no
              other group listed here lets them in.
            </>
          )}
          {verdict === "survives" && (
            <>
              <strong>{who}</strong> will be taken out of <strong>{r.groupName}</strong> — but they{" "}
              <strong>keep access</strong>, because they are also in{" "}
              <strong>{r.survivingGroups.join(", ")}</strong>. Remove them there too if their access
              should stop.
            </>
          )}
          {verdict === "unknown" && (
            <>
              <strong>{who}</strong> will be taken out of <strong>{r.groupName}</strong>. Whether
              that ends their access <strong>cannot be confirmed</strong>: the members of{" "}
              <strong>{r.unreadable.join(", ")}</strong> could not be read, so they may still be in
              one of those.
            </>
          )}

          {/* The consequence nobody expects, and the reason this dialog exists at all. */}
          <div style={m.warnBlock}>{props.scopeWarning}</div>

          {props.inertWarning && <div style={m.infoBlock}>{props.inertWarning}</div>}

          {r.lastMember && (
            <div style={m.infoBlock}>
              This is the last member of <strong>{r.groupName}</strong>. The group keeps its access —
              it will simply have nobody in it. Use <strong>Remove</strong> on the group row itself
              if the group should no longer have access at all.
            </div>
          )}

          <div style={m.note}>
            The person themselves is not deleted, and nothing they have uploaded is affected.
          </div>
        </div>
        <div style={m.foot}>
          <button style={m.ghost} onClick={props.onCancel}>Cancel</button>
          <button
            style={props.busy ? m.dangerOff : m.danger}
            disabled={props.busy}
            onClick={props.onConfirm}
          >
            Remove {who}
          </button>
        </div>
      </div>
    </div>
  );
}

/** Build the removal a dialog needs. Thin wrapper so callers do not re-derive the argument shape. */
export function removalFor(args: {
  user: SpGroupMember;
  group: SpGroup;
  allowedGroupIds: readonly number[];
  members: MemberMap;
  titleOf: (groupId: number) => string;
}): MemberRemoval {
  return buildRemoval({
    user: args.user,
    groupId: args.group.id,
    groupName: args.group.title,
    allowedGroupIds: args.allowedGroupIds,
    map: args.members,
    titleOf: args.titleOf,
  });
}

/** Remove a person from a group. Returns an error message, or undefined on success. */
export async function doRemoveMember(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
  userId: number,
): Promise<string | undefined> {
  try {
    await removeGroupMember(sp, siteUrl, groupId, userId);
    return undefined;
  } catch (e) {
    return (e as Error).message;
  }
}

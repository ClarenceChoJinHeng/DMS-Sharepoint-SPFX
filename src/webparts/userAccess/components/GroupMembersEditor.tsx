// Who is in one SharePoint group — read, add, remove.
//
// Spec: docs/superpowers/specs/2026-08-18-folder-access-membership-design.md §3
//
// MOVED here from GroupManager, not copied. Membership now lives on Folder Access, because putting a
// person into their unit's group is the one genuinely recurring job in this system, and it sat on a
// page an admin otherwise visits twice in a segment's life. Two lists of the same people drift — and
// re-merging the two pages is exactly what the client called confusing on 2026-08-14 — so there is
// one implementation and one mount point.
//
// ⚠ THIS IS A GROUP CHANGE, and can only be. A group's grants are held at FOLDER scope, so adding
// someone here gives them that unit's folder in every library the group is mapped to; removing them
// takes all of it away. It is never scoped to the screen it is rendered on, however the surrounding
// page reads, which is why the caveat is repeated in the hint and in every toast.
//
// NOT accessMemberUi.tsx, which is removal-only and built around the library/page grant rows on
// StagingAccess and PageAccess — including a removalVerdict that has no meaning here. Those screens
// gain no Add, deliberately: adding a person there would look scoped to that library or page, and
// never is.
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  getGroupMembers,
  removeGroupMember,
  searchTenantPeople,
  SpGroupMember,
  PersonPick,
} from "../../../shared/spGroups";
import { addMemberWithSiteEntry } from "../../../shared/siteEntryGroup";
import { siteEntryGroupTitle, isSiteEntryGroupTitle } from "../../../shared/groupMapModel";
import { EVENT } from "../../../shared/auditLog";
import { writeAudit } from "../../../shared/spAuditLog";

type Props = {
  context: WebPartContext;
  siteUrl: string;
  /** The group being edited. Identified by ID — a stored name can be stale or duplicated. */
  group: { id: number; title: string };
  /** Surfaced by the host, so one page has one toast. */
  showToast: (message: string, error: boolean) => void;
};

const s: Record<string, React.CSSProperties> = {
  wrap:    { padding: "8px 10px 14px 24px", background: "#fff", borderTop: "1px solid #f5f4f4" },
  row:     { display: "flex", alignItems: "center", gap: 10, padding: "4px 0" },
  name:    { flex: 1, fontSize: 13 },
  sub:     { color: "#666", fontSize: 12 },
  ghost:   { padding: "5px 12px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  danger:  { padding: "5px 12px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  input:   { width: "100%", maxWidth: 460, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  ddwrap:  { position: "relative", maxWidth: 460, marginTop: 10 },
  dd:      { position: "absolute", top: "100%", left: 0, right: 0, zIndex: 20, background: "#fff", border: "1px solid #c7c7c7", borderRadius: 4, maxHeight: 240, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,.12)" },
  ddItem:  { padding: "7px 10px", cursor: "pointer", borderBottom: "1px solid #f0f0f0", fontSize: 13 },
  hint:    { fontSize: 11, color: "#666", marginTop: 6, lineHeight: 1.45 },
  err:     { fontSize: 11.5, color: "#a4262c", marginTop: 6, lineHeight: 1.45 },
};

export default function GroupMembersEditor({
  context,
  siteUrl,
  group,
  showToast,
}: Props): React.ReactElement {
  /**
   * THREE states, never two. `undefined` is loading and `failed` is a read that did not work —
   * neither is "nobody is in this group". An admin told a group is empty stops looking for the person
   * they came for, and on this page they would then add a second copy of someone already in it.
   */
  const [members, setMembers] = useState<SpGroupMember[] | undefined>(undefined);
  const [failed, setFailed] = useState(false);
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<PersonPick[]>([]);
  const [confirmRemove, setConfirmRemove] = useState<number | undefined>(undefined);
  const [busy, setBusy] = useState(false);

  const reload = async (): Promise<void> => {
    try {
      setMembers(await getGroupMembers(context.spHttpClient, siteUrl, group.id));
      setFailed(false);
    } catch {
      setMembers(undefined);
      setFailed(true);
    }
  };

  // Keyed on the group id, so expanding a different group re-reads rather than showing the last
  // one's people under this one's name.
  useEffect(() => {
    setMembers(undefined);
    setFailed(false);
    setQuery("");
    setResults([]);
    setConfirmRemove(undefined);
    reload().catch(() => undefined);
  }, [group.id]);

  useEffect(() => {
    const q = query.trim();
    if (q.length < 2) { setResults([]); return undefined; }
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q).then(setResults).catch(() => setResults([]));
    }, 350);
    return () => clearTimeout(t);
  }, [query]);

  /** Fire-and-forget: the act has happened, and no admin action should fail because the log was down. */
  const log = (
    summary: string,
    details: string[],
    outcome: "Success" | "Failed" = "Success",
  ): void => {
    writeAudit(context.spHttpClient, siteUrl, {
      event: EVENT.membersChanged,
      outcome,
      source: "FolderAccess",
      at: new Date(),
      actorName: context.pageContext.user.displayName,
      actorEmail: context.pageContext.user.email,
      summary,
      details,
    }).catch(() => undefined);
  };

  const onAdd = async (p: PersonPick): Promise<void> => {
    setBusy(true);
    try {
      const res = await addMemberWithSiteEntry(
        context.spHttpClient, siteUrl, { id: group.id, title: group.title }, p.loginName,
      );
      await reload();
      setQuery("");
      setResults([]);
      // res.note is never dropped: it means they ARE in the group but may not be able to open the
      // site, which neither they nor the admin would discover until they tried.
      showToast(
        `${p.displayName} added to ${group.title} — this grants that group's folders immediately.` +
          (res.note ? ` ${res.note}` : ""),
        !!res.note,
      );
      log(
        `Member added — ${p.displayName} into ${group.title}`,
        [
          `Group: ${group.title}`,
          `Added: ${p.displayName}${p.email ? ` <${p.email}>` : ""}`,
          isSiteEntryGroupTitle(group.title)
            ? "This IS the site-entry group."
            : res.note || `Also added to ${siteEntryGroupTitle()}.`,
        ],
        res.note ? "Failed" : "Success",
      );
    } catch (e) {
      showToast(`Could not add ${p.displayName}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  const onRemove = async (m: SpGroupMember): Promise<void> => {
    setBusy(true);
    try {
      await removeGroupMember(context.spHttpClient, siteUrl, group.id, m.id);
      await reload();
      setConfirmRemove(undefined);
      showToast(
        `${m.title} removed from ${group.title} — that group's folder access is revoked immediately.`,
        false,
      );
      log(
        `Member removed — ${m.title} from ${group.title}`,
        [
          `Group: ${group.title}`,
          `Removed: ${m.title}${m.email ? ` <${m.email}>` : ""}`,
          // Removing from one group is not removing from the site. Saying so stops a reader
          // concluding the person was de-provisioned.
          `This removes ONE group. Their membership of ${siteEntryGroupTitle()} and any other group is unchanged.`,
        ],
      );
    } catch (e) {
      showToast(`Could not remove ${m.title}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={s.wrap}>
      {members === undefined && !failed && <p style={s.hint}>Loading members&hellip;</p>}
      {failed && (
        <p style={s.err}>
          Could not read this group&apos;s members — so this is <strong>not</strong> a statement that
          nobody is in it. Adding someone still works.
        </p>
      )}
      {members !== undefined && members.length === 0 && (
        <p style={s.hint}>
          Nobody is in this group yet. Its folder permissions exist and apply the moment someone is added.
        </p>
      )}
      {(members ?? []).map((m) => (
        <div key={m.id} style={s.row}>
          <span style={s.name}>
            {m.title} <span style={s.sub}>{m.email}</span>
          </span>
          {confirmRemove === m.id ? (
            <>
              <button type="button" style={s.danger} disabled={busy} onClick={() => { onRemove(m).catch(() => undefined); }}>
                Confirm remove
              </button>
              <button type="button" style={s.ghost} onClick={() => setConfirmRemove(undefined)}>Cancel</button>
            </>
          ) : (
            <button type="button" style={s.ghost} disabled={busy} onClick={() => setConfirmRemove(m.id)}>Remove</button>
          )}
        </div>
      ))}

      <div style={s.ddwrap}>
        <input
          style={s.input}
          value={query}
          placeholder="Add a person&hellip;"
          disabled={busy}
          onChange={(e) => setQuery(e.target.value)}
        />
        {results.length > 0 && (
          <div style={s.dd}>
            {results.map((p) => (
              <div key={p.loginName} style={s.ddItem} onClick={() => { onAdd(p).catch(() => undefined); }}>
                {p.displayName} <span style={s.sub}>{p.email}</span>
              </div>
            ))}
          </div>
        )}
      </div>
      <p style={s.hint}>
        Everyone added here also joins <strong>{siteEntryGroupTitle()}</strong>, without which they
        cannot open the site at all. Membership is a <strong>group</strong> change: it applies to every
        folder this group is mapped to, in every library — not only to what is listed above.
      </p>
    </div>
  );
}

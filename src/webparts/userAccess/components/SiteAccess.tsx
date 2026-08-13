// Site Access tab — who can open the site at all, and the membership that makes it work.
//
// This is the layer everything else sits on. A folder grant gives a user Limited Access up the
// parent chain, which resolves a DIRECT link but confers no View Pages — so without site-level
// Read a correctly provisioned uploader still cannot open the site's home page. That is the
// failure this tab exists to make visible and fixable without going to groups.aspx.
//
// Three things, in the order they fail:
//   1. Does the site-entry group EXIST and hold a role on the web? If not, nobody gets in.
//   2. Who is in it? Adding a person here is the whole job of onboarding a user.
//   3. Who is in a mapped group but NOT in it? That set is the silent failure — those users
//      have folder permissions and cannot reach the site. Reconciliation heals it on a run;
//      this tab names them and fixes them on demand.
//
// See docs/superpowers/specs/2026-08-03-access-scope-mapping-design.md and memory
// dms-two-layer-access-site-plus-folder.
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { siteEntryGroupTitle } from "../../../shared/groupMapModel";
import {
  fetchAllSiteGroups,
  createSiteGroup,
  getGroupMembers,
  addGroupMember,
  removeGroupMember,
  searchTenantPeople,
  SpGroup,
  SpGroupMember,
  PersonPick,
} from "../../../shared/spGroups";
import { AuditOutcome, EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";

type Props = { context: WebPartContext; siteUrl: string };

// Resolved, not hardcoded — the client renames this to "CRS Group Map" at import (confirmed on
// their site 2026-08-05). Read from the cache primed in reload(), which runs on mount before the
// Group Map is read.
const groupMapList = (): string => cachedListTitle(LIST_SUFFIX.groupMap);

type WebGrant = { principalId: number; title: string; levels: string[] };
/** A user who holds a mapped group but is missing site entry. */
type Missing = { loginName: string; title: string; via: string };

const s: Record<string, React.CSSProperties> = {
  wrap:     { fontSize: 13, color: "#242424", lineHeight: 1.5 },
  intro:    { fontSize: 13, color: "#444", margin: "0 0 16px" },
  card:     { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  head:     { fontWeight: 600, fontSize: 13, margin: "0 0 8px" },
  table:    { width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 8 },
  th:       { textAlign: "left", padding: "6px 8px", borderBottom: "2px solid #e1e1e1", fontWeight: 600, color: "#555" },
  td:       { padding: "6px 8px", borderBottom: "1px solid #f0f0f0", verticalAlign: "middle" },
  input:    { width: "100%", maxWidth: 420, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  ddwrap:   { position: "relative", maxWidth: 420 },
  dd:       { position: "absolute", top: "100%", left: 0, right: 0, zIndex: 20, background: "#fff", border: "1px solid #c7c7c7", borderRadius: 4, maxHeight: 240, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,.12)" },
  ddItem:   { padding: "7px 10px", cursor: "pointer", borderBottom: "1px solid #f0f0f0" },
  addBtn:   { padding: "5px 14px", fontSize: 12, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  delBtn:   { padding: "4px 12px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  off:      { padding: "5px 14px", fontSize: 12, background: "#c7c7c7", color: "#fff", border: "none", borderRadius: 4, cursor: "not-allowed" },
  ghost:    { padding: "5px 14px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  okBox:    { marginBottom: 16, padding: "10px 12px", border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 4, fontSize: 12, color: "#0f6c3f", lineHeight: 1.5 },
  warnBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5 },
  dangerBox:{ marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  hint:     { fontSize: 11, color: "#666", marginTop: 8, lineHeight: 1.45 },
  no:       { color: "#8a8886" },
  mono:     { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
  toast:    { position: "fixed", bottom: 20, right: 20, padding: "10px 16px", borderRadius: 6, color: "#fff", fontSize: 13, zIndex: 50, boxShadow: "0 4px 14px rgba(0,0,0,.18)" },
  spinner:  { display: "inline-block", width: 11, height: 11, marginRight: 6, border: "2px solid rgba(255,255,255,.45)", borderTopColor: "#fff", borderRadius: "50%", verticalAlign: "-1px", animation: "sa2-spin .6s linear infinite" },
};

const SPIN_KEYFRAMES = "@keyframes sa2-spin{to{transform:rotate(360deg)}}";

export default function SiteAccess({ context, siteUrl }: Props): React.ReactElement {
  const [entry, setEntry]     = useState<SpGroup | undefined>(undefined);
  const [entryHasRole, setEntryHasRole] = useState<boolean | undefined>(undefined);
  const [members, setMembers] = useState<SpGroupMember[] | undefined>(undefined);
  const [webGrants, setWebGrants] = useState<WebGrant[] | undefined>(undefined);
  const [missing, setMissing] = useState<Missing[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy]       = useState(false);
  const [progress, setProgress] = useState<{ done: number; total: number } | undefined>(undefined);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [toast, setToast]     = useState<{ message: string; error: boolean } | undefined>(undefined);

  // People picker
  const [query, setQuery]     = useState("");
  const [results, setResults] = useState<PersonPick[]>([]);
  const [searching, setSearching] = useState(false);

  const GET = { Accept: "application/json;odata=nometadata" };

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(undefined), 7000);
  };

  const loadWebGrants = async (): Promise<WebGrant[] | undefined> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings` +
        `&$select=PrincipalId,Member/Title,RoleDefinitionBindings/Name`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) return undefined;
    const j = await res.json();
    return ((j.value ?? []) as Array<{ PrincipalId?: number; Member?: { Title?: string }; RoleDefinitionBindings?: Array<{ Name?: string }> }>)
      .map((ra) => ({
        principalId: ra.PrincipalId ?? 0,
        title: ra.Member?.Title ?? "",
        levels: (ra.RoleDefinitionBindings ?? []).map((b) => b.Name ?? "").filter(Boolean),
      }));
  };

  /**
   * Every group id the Group Map references.
   *
   * GroupId ONLY: that column predates Scope/Target, so this cannot hit the whole-request 400
   * that naming a nonexistent $select column causes (CLAUDE.md #11), and it covers every scope
   * in one read.
   */
  const loadMappedGroupIds = async (): Promise<number[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(groupMapList())}')/items?$select=GroupId&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) return [];
    const j = await res.json();
    const ids = new Set<number>();
    for (const r of (j.value ?? []) as Array<{ GroupId?: string }>) {
      const n = Number((r.GroupId ?? "").toString().trim());
      if (n > 0 && n % 1 === 0) ids.add(n);
    }
    return Array.from(ids);
  };

  const reload = async (firstLoad = false): Promise<void> => {
    if (firstLoad) setLoading(true);
    try {
      // Before the Group Map read: resolves "CRS Group Map" vs "DMS Group Map" once per session.
      await primeNames(context.spHttpClient, siteUrl);
      const all = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
      const eg = all.find((g) => g.title.trim().toLowerCase() === siteEntryGroupTitle().toLowerCase());
      setEntry(eg);
      const grants = await loadWebGrants();
      setWebGrants(grants);
      setEntryHasRole(
        eg === undefined ? false : grants === undefined ? undefined : grants.some((g) => g.principalId === eg.id),
      );

      if (!eg) { setMembers(undefined); setMissing([]); setLoadError(undefined); return; }
      const mem = await getGroupMembers(context.spHttpClient, siteUrl, eg.id);
      setMembers(mem);

      // Who holds a mapped group but is missing from site entry? Computed here rather than
      // waiting for a reconciliation run, because this is the state that presents as "only some
      // users cannot open the site" and it is invisible everywhere else.
      const have = new Set(mem.map((m) => m.loginName.toLowerCase()));
      const mappedIds = await loadMappedGroupIds();
      const gap = new Map<string, Missing>();
      for (const gid of mappedIds) {
        if (gid === eg.id) continue;
        const g = all.find((x) => x.id === gid);
        const gm = await getGroupMembers(context.spHttpClient, siteUrl, gid).catch(() => []);
        for (const m of gm) {
          const key = m.loginName.toLowerCase();
          if (have.has(key) || gap.has(key)) continue;
          gap.set(key, { loginName: m.loginName, title: m.title, via: g?.title ?? String(gid) });
        }
      }
      setMissing(Array.from(gap.values()));
      setLoadError(undefined);
    } catch (e) {
      setLoadError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { reload(true).catch(() => undefined); }, []);

  // Debounced people search — 3 characters minimum, same threshold as the Group Map builder.
  useEffect(() => {
    const q = query.trim();
    if (q.length < 3) { setResults([]); return undefined; }
    let cancelled = false;
    setSearching(true);
    const t = setTimeout(() => {
      // .catch LAST, so the chain is not a floating promise. Clearing `searching` has to happen
      // on both paths or a failed search leaves "Searching…" on screen forever.
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then((r) => { if (!cancelled) { setResults(r); setSearching(false); } })
        .catch(() => { if (!cancelled) { setResults([]); setSearching(false); } });
    }, 350);
    return () => { cancelled = true; clearTimeout(t); };
  }, [query]);

  /**
   * Record a site-entry change. Fire-and-forget: the access change has already happened, and no
   * caller should be able to fail because the log was unreachable.
   */
  const logSite = (event: string, summary: string, details: string[], outcome: AuditOutcome): void => {
    writeAudit(context.spHttpClient, siteUrl, {
      event,
      outcome,
      source: "SiteAccess",
      at: new Date(),
      actorName: context.pageContext.user.displayName,
      actorEmail: context.pageContext.user.email,
      summary,
      details,
    }).catch(() => undefined);
  };

  /**
   * Create the site-entry group if absent and give it Read on the web.
   *
   * Both halves, because either alone is useless: a group with no role grants nothing, and there
   * is no role to grant without a group. Same work as the reconciliation site-entry pass, offered
   * here so a new site can be made usable before the first run.
   */
  const onSetUp = async (): Promise<void> => {
    setBusy(true);
    try {
      let eg = entry;
      if (!eg) eg = await createSiteGroup(context.spHttpClient, siteUrl, siteEntryGroupTitle());
      const defs = await context.spHttpClient.get(
        `${siteUrl}/_api/web/roledefinitions?$select=Id,Name`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!defs.ok) throw new Error(`roledefinitions HTTP ${defs.status}`);
      const dj = await defs.json();
      const readId = ((dj.value ?? []) as Array<{ Id: number; Name: string }>).find((d) => d.Name === "Read")?.Id;
      if (readId === undefined) throw new Error('no "Read" permission level on this site');
      const grant = await context.spHttpClient.post(
        `${siteUrl}/_api/web/roleassignments/addroleassignment(principalid=${eg.id},roledefid=${readId})`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!grant.ok) throw new Error(`addroleassignment HTTP ${grant.status}`);
      await reload();
      showToast(`${siteEntryGroupTitle()} is set up and can open the site.`, false);
      logSite(
        EVENT.accessGranted,
        `Site entry set up — ${siteEntryGroupTitle()} can open the site`,
        [
          `Group: ${siteEntryGroupTitle()}${entry ? " (already existed)" : " (created)"}`,
          "Granted Read on the site itself. This grants nothing inside either library.",
        ],
        "Success",
      );
    } catch (e) {
      showToast(`Set-up failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  const onAddPerson = async (p: PersonPick): Promise<void> => {
    if (!entry) return;
    setBusy(true);
    setQuery("");
    setResults([]);
    try {
      await addGroupMember(context.spHttpClient, siteUrl, entry.id, p.loginName);
      await reload();
      showToast(`${p.displayName} can now open the site.`, false);
      logSite(
        EVENT.accessGranted,
        `Site access granted — ${p.displayName} can open the site`,
        [
          `Added ${p.displayName} to ${siteEntryGroupTitle()}.`,
          "This grants site entry only — it confers nothing inside either library.",
        ],
        "Success",
      );
    } catch (e) {
      showToast(`Could not add ${p.displayName}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  const onRemovePerson = async (m: SpGroupMember): Promise<void> => {
    if (!entry) return;
    setBusy(true);
    try {
      await removeGroupMember(context.spHttpClient, siteUrl, entry.id, m.id);
      await reload();
      // Stated rather than implied: their folder permissions are untouched, so this is a site
      // lock-out, not a de-provisioning. An admin who expects the latter would stop here.
      showToast(`${m.title} can no longer open the site. Their folder permissions are unchanged.`, false);
      // The folder note is carried into the RECORD as well as the toast. A row reading only
      // "site access revoked" would later be mistaken for a de-provisioning, and someone would
      // stop looking while every folder ACL was still in place.
      logSite(
        EVENT.accessRevoked,
        `Site access revoked — ${m.title} can no longer open the site`,
        [
          `Removed ${m.title} from ${siteEntryGroupTitle()}.`,
          "Their folder permissions are UNCHANGED — this is a site lock-out, not a de-provisioning.",
        ],
        "Success",
      );
    } catch (e) {
      showToast(`Could not remove ${m.title}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  /** Add everyone who holds a mapped group but is missing site entry. Sequential, for throttling. */
  const onFixMissing = async (): Promise<void> => {
    if (!entry || missing.length === 0) return;
    setBusy(true);
    setProgress({ done: 0, total: missing.length });
    let ok = 0;
    const failed: string[] = [];
    for (let i = 0; i < missing.length; i++) {
      try {
        await addGroupMember(context.spHttpClient, siteUrl, entry.id, missing[i].loginName);
        ok++;
      } catch {
        failed.push(missing[i].title);
      }
      setProgress({ done: i + 1, total: missing.length });
    }
    setProgress(undefined);
    await reload();
    setBusy(false);
    showToast(
      failed.length === 0
        ? `${ok} user(s) can now open the site.`
        : `${ok} added, ${failed.length} failed: ${failed.join(", ")}`,
      failed.length > 0,
    );
    // One row for the run, not one per user — a fix-missing over a whole segment would otherwise
    // fill the log's first page by itself.
    logSite(
      EVENT.accessGranted,
      `Site access granted — ${ok} user(s) can now open the site` +
        (failed.length > 0 ? `, ${failed.length} failed` : ""),
      [
        `Added ${ok} user(s) to ${siteEntryGroupTitle()} who held a mapped group but had no site entry.`,
        failed.length > 0 ? `Failed: ${failed.join(", ")}` : "No failures.",
      ],
      failed.length > 0 ? "Failed" : "Success",
    );
  };

  const setUpDone = entry !== undefined && entryHasRole === true;

  return (
    <div style={s.wrap}>
      <style>{SPIN_KEYFRAMES}</style>
      <p style={s.intro}>
        Who can <strong>open this site</strong>. Folder and library permissions are not enough on
        their own — without site access a user can only reach a folder by direct link, and the home
        page refuses them. Everyone who uses the system needs to be here.
      </p>

      {loadError && <div style={s.dangerBox}>Could not load: {loadError}</div>}

      {loading ? (
        <div style={s.card}><span style={s.no}>Loading&hellip;</span></div>
      ) : (
        <>
          {!setUpDone && (
            <div style={s.dangerBox}>
              {!entry ? (
                <>
                  <strong>{siteEntryGroupTitle()} does not exist.</strong> Until it does, nobody can
                  open this site except owners and administrators.
                </>
              ) : entryHasRole === undefined ? (
                <>
                  <strong>{siteEntryGroupTitle()} exists</strong>, but its site permissions could
                  not be read, so whether it works is unknown.
                </>
              ) : (
                <>
                  <strong>{siteEntryGroupTitle()} exists but holds no permission on the site</strong>,
                  so its members still cannot open it.
                </>
              )}
              <div style={{ marginTop: 8 }}>
                <button style={busy ? s.off : s.addBtn} disabled={busy} onClick={() => { onSetUp().catch(() => undefined); }}>
                  {busy ? <><span style={s.spinner} />Working&hellip;</> : "Set up site access"}
                </button>
              </div>
            </div>
          )}

          {setUpDone && (
            <div style={s.okBox}>
              <strong>{siteEntryGroupTitle()}</strong> holds <strong>Read</strong> on the site.
              Anyone in it can open the site; what they see inside is decided by their folder and
              library permissions.
            </div>
          )}

          {missing.length > 0 && entry && (
            <div style={s.warnBox}>
              <strong>{missing.length} user(s) have folder permissions but cannot open the
              site.</strong> They were added to a group directly in SharePoint, which does not grant
              site access on its own — so for them the system looks broken while everyone else is
              fine.
              <div style={s.mono}>
                {missing.slice(0, 12).map((m) => `${m.title} (via ${m.via})`).join(", ")}
                {missing.length > 12 ? `, +${missing.length - 12} more` : ""}
              </div>
              <div style={{ marginTop: 8 }}>
                <button style={busy ? s.off : s.addBtn} disabled={busy} onClick={() => { onFixMissing().catch(() => undefined); }}>
                  {progress
                    ? <><span style={s.spinner} />Adding {progress.done} of {progress.total}&hellip;</>
                    : `Give all ${missing.length} site access`}
                </button>
              </div>
            </div>
          )}

          {entry && (
            <div style={s.card}>
              <div style={s.head}>People who can open the site</div>
              <div style={s.ddwrap}>
                <input
                  style={s.input}
                  value={query}
                  disabled={busy}
                  placeholder="Search for a person to add…"
                  onChange={(e) => setQuery(e.target.value)}
                />
                {results.length > 0 && (
                  <div style={s.dd}>
                    {results.map((p) => (
                      <div
                        key={p.loginName}
                        style={s.ddItem}
                        onClick={() => { onAddPerson(p).catch(() => undefined); }}
                      >
                        {p.displayName}
                        {p.email ? <span style={s.mono}> {p.email}</span> : undefined}
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <div style={s.hint}>
                {searching
                  ? "Searching…"
                  : "Type at least 3 characters. Adding someone here lets them open the site — it grants no documents."}
              </div>

              {members === undefined ? (
                <div style={s.no}>Members could not be read.</div>
              ) : members.length === 0 ? (
                <div style={s.no}>Nobody has site access yet.</div>
              ) : (
                <table style={s.table}>
                  <thead>
                    <tr>
                      <th style={s.th}>Name</th>
                      <th style={s.th}>Email</th>
                      <th style={s.th} />
                    </tr>
                  </thead>
                  <tbody>
                    {members.map((m) => (
                      <tr key={m.id}>
                        <td style={s.td}>{m.title}</td>
                        <td style={s.td}><span style={s.mono}>{m.email}</span></td>
                        <td style={s.td}>
                          <button
                            style={busy ? s.off : s.delBtn}
                            disabled={busy}
                            onClick={() => { onRemovePerson(m).catch(() => undefined); }}
                          >
                            Remove
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}

          <div style={s.card}>
            <div style={s.head}>Everything with permission on the site itself</div>
            <div style={s.hint}>
              Read-only. Site-level permissions reach every library and folder that still inherits,
              so anything unexpected here is worth understanding before it is removed — removing the
              wrong entry can lock administrators out of the site.
            </div>
            {webGrants === undefined ? (
              <div style={s.no}>Could not read the site&rsquo;s permissions.</div>
            ) : (
              <table style={s.table}>
                <thead>
                  <tr>
                    <th style={s.th}>Name</th>
                    <th style={s.th}>Permission</th>
                  </tr>
                </thead>
                <tbody>
                  {webGrants.map((g) => (
                    <tr key={g.principalId}>
                      <td style={s.td}>{g.title || g.principalId}</td>
                      <td style={s.td}>{g.levels.join(", ") || "none"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#a4262c" : "#0f6c3f" }}>
          {toast.message}
        </div>
      )}
    </div>
  );
}

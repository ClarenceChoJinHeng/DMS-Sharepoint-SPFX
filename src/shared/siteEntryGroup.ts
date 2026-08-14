/**
 * The site-entry group — finding it, creating it, and keeping members inside it.
 *
 * Spec: docs/superpowers/specs/2026-08-14-group-management-separation-design.md §2 D5
 * Background: docs/superpowers/specs/2026-07-27-site-entry-access-layer-design.md
 *
 * WHY THIS FILE EXISTS. A folder grant alone confers Limited Access: the holder can open that one
 * folder by direct link, but the site root denies them, so they can reach nothing by navigating.
 * Every DMS group member therefore also has to be in the site-entry group. That rule was written
 * out three separate times — in GroupMapBuilder, in SiteAccess and inside Folder Reconciliation —
 * and the Group Management page would have made a fourth. It is load-bearing and invisible when
 * wrong: the symptom is a correctly provisioned user who "cannot see anything", which reads as a
 * permissions bug and is not one.
 *
 * ABSENT AND UNREADABLE ARE DIFFERENT STATES, and this module never collapses them. An absent
 * entry group is a setup step the admin can take, and saying so is useful. An unreadable group
 * list proves nothing, and reporting it as "not found" sends the admin off to create a group that
 * already exists — where they meet a duplicate-name error for a group they cannot see. The same
 * rule as `unknown` ≠ empty everywhere else in this codebase.
 */
import { SPHttpClient } from "@microsoft/sp-http";
import { findSiteEntryGroup, isSiteEntryGroupTitle, siteEntryGroupTitle } from "./groupMapModel";
import { addGroupMember, createSiteGroup, fetchAllSiteGroups, SpGroup } from "./spGroups";

export type SiteEntryState = "found" | "absent" | "unknown";

export interface SiteEntryLookup {
  state: SiteEntryState;
  /** Set only when `state === "found"`. */
  group?: SpGroup;
}

/**
 * Session cache of the lookup, per site.
 *
 * The PROMISE is cached, assigned synchronously, so there is no read-modify-write of a shared
 * value across an await — the pattern `require-atomic-updates` exists to catch. Without a cache,
 * every keystroke of a people search would re-read the site's whole group list.
 *
 * `unknown` is deliberately NOT kept: it means a request failed, and a transient failure must not
 * pin the wrong answer for the rest of the page session.
 */
const lookupCache = new Map<string, Promise<SiteEntryLookup>>();

/** Drop the cached lookup. Call after creating the group, or the next read still says absent. */
export function clearSiteEntryCache(siteUrl?: string): void {
  if (siteUrl === undefined) lookupCache.clear();
  else lookupCache.delete(siteUrl);
}

/** Find the site-entry group. Never throws: a failed read is reported as `unknown`. */
export async function resolveSiteEntryGroup(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<SiteEntryLookup> {
  const cached = lookupCache.get(siteUrl);
  if (cached) return cached;

  const pending: Promise<SiteEntryLookup> = fetchAllSiteGroups(sp, siteUrl).then(
    (all): SiteEntryLookup => {
      const hit = findSiteEntryGroup(all);
      return hit ? { state: "found", group: hit } : { state: "absent" };
    },
    (): SiteEntryLookup => ({ state: "unknown" }),
  );
  lookupCache.set(siteUrl, pending);

  const result = await pending;
  if (result.state === "unknown") lookupCache.delete(siteUrl);
  return result;
}

export interface EnsureResult {
  group: SpGroup;
  /** True when this call created it. Callers log the two cases differently — only one is news. */
  created: boolean;
}

/**
 * Find the site-entry group, creating it when it is absent.
 *
 * THROWS when the group list could not be read. Creating on `unknown` would attempt a group that
 * may already exist, and the duplicate-name failure that follows is a confusing way to report a
 * network error. Callers that can carry on without it should catch — Folder Reconciliation does,
 * because provisioning folders is still worth doing.
 *
 * This creates the group ONLY. Granting it Read on the web is a separate act with its own failure
 * mode, and stays with the callers that do it (SiteAccess, Folder Reconciliation): a group holding
 * no role grants nothing, so creation alone must never be reported as success.
 */
export async function ensureSiteEntryGroup(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<EnsureResult> {
  const found = await resolveSiteEntryGroup(sp, siteUrl);
  if (found.state === "found" && found.group) return { group: found.group, created: false };
  if (found.state === "unknown") {
    throw new Error(`could not read this site's groups, so ${siteEntryGroupTitle()} was not checked`);
  }
  const made = await createSiteGroup(sp, siteUrl, siteEntryGroupTitle());
  clearSiteEntryCache(siteUrl);
  return { group: made, created: true };
}

export interface AddMemberResult {
  /** True when the person reached the target group. False is unreachable — the call throws first. */
  added: boolean;
  /**
   * A caveat to show alongside the success message, or "" when there is none.
   *
   * Non-empty means the person IS in the group they were added to but may not be able to OPEN THE
   * SITE — which neither they nor the admin would discover until they tried.
   */
  note: string;
}

/**
 * Add a person to a group, and to the site-entry group with them.
 *
 * The target add is the half that may throw: if it fails, nothing happened and the caller should
 * say so plainly. The entry-group half NEVER throws — the person is already in the group they were
 * added to, and undoing that to report a secondary problem would be worse than describing it. It
 * comes back as `note` instead, and every caller MUST surface it. An unspoken note here is exactly
 * the silent half-success this module exists to prevent.
 *
 * Skipped when the target IS the entry group. Re-adding would be harmless, but the note below
 * would then describe a group as missing from itself.
 */
export async function addMemberWithSiteEntry(
  sp: SPHttpClient,
  siteUrl: string,
  target: { id: number; title: string },
  loginName: string,
): Promise<AddMemberResult> {
  await addGroupMember(sp, siteUrl, target.id, loginName);

  if (isSiteEntryGroupTitle(target.title)) return { added: true, note: "" };

  const entry = await resolveSiteEntryGroup(sp, siteUrl);
  if (entry.state === "absent") {
    return {
      added: true,
      note: `${siteEntryGroupTitle()} does not exist yet — set it up on the Site Access page, or they will not be able to open the site.`,
    };
  }
  if (entry.state !== "found" || !entry.group) {
    return {
      added: true,
      note: `Could not check ${siteEntryGroupTitle()}, so site entry is unconfirmed. Check it on the Site Access page.`,
    };
  }
  try {
    await addGroupMember(sp, siteUrl, entry.group.id, loginName);
    return { added: true, note: "" };
  } catch (e) {
    return {
      added: true,
      note: `Could not add them to ${siteEntryGroupTitle()} (${(e as Error).message}) — they may not be able to open the site.`,
    };
  }
}

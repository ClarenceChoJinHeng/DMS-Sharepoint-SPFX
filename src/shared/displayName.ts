/**
 * A readable name derived from an email address, for any screen that shows "who did this" and only
 * has an email address to show it with.
 *
 * Extracted from `AuditLog.tsx`'s `nameFromEmail` (added there 2026-09-13, client QA item #48.4:
 * "standardise all email address to user name") — that screen already treats a raw address as
 * something to be repaired for display, and the same complaint now applies to every OTHER screen
 * that names a requester/approver/reviser by their bare email: `Requests.tsx` ("Requested by …",
 * "Decided by …", "Revoked by …", "asked by …", "approved by …") and `MySubmissions.tsx`
 * ("Approved by …", the Requests tab's "by …" line, the "Approved By" table column). One person
 * asked twice, in two files, is how the two definitions drift — this is the one copy.
 *
 * ⚠ DISPLAY ONLY, always. The underlying stored value (an `ActorEmail`, `DecidedBy`, `ApprovedBy`,
 * `RequestedBy`, `RevokedBy` field) is NEVER touched by this — every writer, every filter/search
 * comparison, every CSV export and every audit row keeps the real address. Only the on-screen LABEL
 * changes. Getting this backwards (deriving a name and then writing THAT back anywhere) would lose
 * the one thing that never lies about who actually did something.
 *
 * Splits on the local part only (before `@`), replaces dots and underscores with spaces, and
 * title-cases each word — `chiew.wei.chien@sdguthrie.com` reads as "Chiew Wei Chien" rather than the
 * bare address. This is a GUESS at a real name from an address convention, not a directory lookup —
 * it cannot be right for every possible mailbox naming scheme, and does not try to be. It is strictly
 * better than a raw address for the common `first.last@domain` shape this tenant actually uses.
 */
export function nameFromEmail(email: string): string {
  const local = email.split("@")[0] ?? "";
  if (!local) return email;
  return local
    .split(/[._]+/)
    .filter((w) => w.length > 0)
    .map((w) => w[0].toUpperCase() + w.slice(1))
    .join(" ");
}

/**
 * `nameFromEmail`, but safe over an OPTIONAL/blank value — the shape every one of these call sites
 * actually has (`r.decidedBy`, `r.approvedBy`, `r.requestedBy` are all `string | undefined`, and can
 * be an empty string too). Returns `undefined` for anything not worth naming, so a caller's existing
 * `r.decidedBy && …` guard keeps working unchanged — this never turns "nothing to show" into the
 * word "undefined" or a lone "—".
 */
export function displayNameFor(email: string | undefined): string | undefined {
  const trimmed = (email ?? "").trim();
  if (trimmed.length === 0) return undefined;
  return canonicalServiceAccountName(trimmed) ?? nameFromEmail(trimmed);
}

/**
 * The two proxy/service accounts this project's Power Automate flows run as, keyed by the LOCAL
 * PART only (lower-cased) so a bare mailbox alias and a full address both resolve — see
 * `canonicalServiceAccountName`'s own comment for why both shapes turn up in stored data.
 *
 * `gdc` is the CURRENT proxy (`gdc@sdguthrie.com`, "Guthrie Document Centre" — the migration from
 * `crs@sdguthrie.com` is written up in
 * `docs/superpowers/specs/2026-09-19-service-account-migration-crs-to-gdc-runbook.md`, and is not
 * finished across every flow at the time this was written). `crs` is the RETIRED one
 * (`crs@sdguthrie.com`, "Guthrie Central Repository System") — kept here, not deleted, because
 * historical audit rows written before the migration still carry it and must go on reading
 * correctly for as long as this log exists (this list is append-only and never deleted).
 */
const KNOWN_SERVICE_ACCOUNTS: Readonly<Record<string, string>> = {
  gdc: "Guthrie Document Centre",
  crs: "Guthrie Central Repository System",
};

/**
 * Recognises one of THIS project's own proxy accounts and returns its real, canonical display
 * name — checked FIRST, ahead of any guessed or stored name, because a service account is the one
 * case where the true name is known exactly rather than guessed from an address.
 *
 * Built 2026-09-20 after live evidence (client screenshot) showed the SAME account rendering three
 * different ways across three adjacent Audit Log rows: "Guthrie Document Centre" (correct, an
 * `ActorName` a flow happened to write out in full), `gdc@sdguthrie.com` (the raw address, stored
 * literally AS `ActorName` by a different flow action) and `gdc` (a bare alias, stored the same
 * way by a third). None of those three are "wrong" data exactly — they are inconsistent COPIES of
 * one fact, written by different Power Automate actions that were never told to agree on a shape.
 * Fixing that on the flow side means finding and correcting every such action across (at least) the
 * request-decision flows; fixing it HERE, once, closes it for every screen that shows an actor
 * immediately, and cannot be defeated by a fourth flow doing it a fourth way tomorrow.
 *
 * Matched on the LOCAL PART, case-insensitively, so `gdc`, `GDC`, `gdc@sdguthrie.com` and
 * `gdc@SDGuthrie.com` — every shape actually observed live — all resolve identically. Returns
 * `undefined` for anything else, so a real person's address falls straight through to the normal
 * `nameFromEmail` guess untouched.
 */
export function canonicalServiceAccountName(
  value: string,
): string | undefined {
  const trimmed = value.trim().toLowerCase();
  if (trimmed.length === 0) return undefined;
  const local = trimmed.split("@")[0] ?? "";
  return KNOWN_SERVICE_ACCOUNTS[local];
}

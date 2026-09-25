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
 * The proxy/service account IDENTITY this project's Power Automate flows currently run as, keyed by
 * the LOCAL PART only (lower-cased) so a bare mailbox alias and a full address both resolve — see
 * `canonicalServiceAccountName`'s own comment for why both shapes turn up in stored data.
 *
 * ⚠ SUPERSEDED 2026-09-25 — `crs` IS NO LONGER IN THIS MAP, AND IT WAS FOR ONE DAY.
 * `docs/superpowers/specs/2026-09-19-service-account-migration-crs-to-gdc-runbook.md` moved the
 * FLOWS' own connection from `crs@sdguthrie.com` to `gdc@sdguthrie.com`; separately, the client
 * confirmed `crs@sdguthrie.com` "wont be the system admin anymore... crs is now a normal account" —
 * it is a live human test/approver account, not a retired proxy identity. On 2026-09-23 this map
 * briefly collapsed `crs` into `gdc`'s display name, on the reasoning that the client saw the two as
 * one identity across time. That broke the very next day: the client approved a document AS `crs`
 * and the Audit Log's "Approved" row (correctly carrying `crs`'s own address as `ApprovedBy`) showed
 * "Guthrie Document Centre" instead — reading as though the automated proxy had approved it, when a
 * person actually had. **A genuine human action performed by `crs` must show `crs`, not the system
 * that only ever performs UNATTENDED steps (routing, archiving, deletion execution).** `crs` is
 * removed from this map entirely — not merely un-aliased to a different label — so it falls straight
 * through `resolveActorDisplay` to the ordinary name-guessing path, exactly as any other real
 * account's address would. If `crs@sdguthrie.com` is ever retired again into a purely automated
 * role, re-add it here under `CURRENT_PROXY_ACCOUNT_NAME` at that point, not before.
 */
const KNOWN_SERVICE_ACCOUNTS: Readonly<Record<string, string>> = {
  gdc: "Guthrie Document Centre",
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

/**
 * The label an Audit Log "Who" cell should show, given the row's stored `ActorEmail`/`ActorName`.
 *
 * Built 2026-09-20 after checking the ACTUAL flow definitions (`PowerAutomateFlowsSDG/*.zip`), not
 * assuming from a runbook. Two different flows write two different "already partly processed, but
 * still not a real name" shapes into `ActorName`, and BOTH defeated the naive `r.ActorName ||
 * nameFromEmail(r.ActorEmail)` this file used to build the Who cell with:
 *
 *   - `CRS — Audit request activity` writes the FULL raw address into `ActorName`
 *     (`item/ActorName: "@outputs('ActorEmail')"`) — e.g. `clarence@trinergydigital.com`.
 *   - `CRS — Execute approved deletion` writes just the LOCAL PART, dots and all
 *     (`first(split(..., '@'))`) — e.g. `goh.kheng.wei`.
 *
 * Neither is blank and neither is a real display name, so the old logic treated both as "already
 * good, show verbatim" and never reached the name-guessing fallback at all. This function instead
 * treats `ActorName` as INPUT worth normalising rather than a value to trust outright: a KNOWN
 * service account (either field) wins first; otherwise `ActorName` (falling back to `ActorEmail` if
 * blank) is run through `nameFromEmail` unconditionally. This is a safe no-op for a genuinely good
 * display name — Auto-route's own `ActorName` (`triggerOutputs()?['body/Author/DisplayName']`, a
 * real SharePoint People field's DisplayName like "Clarence Cho") has no `@`, `.` or `_` to split
 * on, so it passes through essentially unchanged — and it correctly cleans up either broken shape
 * above, because `nameFromEmail` itself already splits on `@` first.
 */
export function resolveActorDisplay(
  actorEmail: string | undefined,
  actorName: string | undefined,
): string {
  const known =
    canonicalServiceAccountName(actorEmail ?? "") ??
    canonicalServiceAccountName(actorName ?? "");
  if (known) return known;

  const source =
    actorName && actorName.trim().length > 0 ? actorName : actorEmail;
  return source ? nameFromEmail(source) : "—";
}

/**
 * The CURRENT proxy account's full address — the one Power Automate flows run as, and the one
 * `shared/dmsFolderMap.ts`'s `stampEditorAsProxy` resolves and writes into `Modified By` on files
 * this app's own CODE moves (as opposed to a flow moving them, where the flow's OWN connection
 * identity already IS this account and needs no separate stamp).
 *
 * ⚠ ONE LITERAL, HERE ONLY — every caller that needs "which account is the proxy right now" reads
 * this constant rather than typing the address again, so the day this migrates a second time (as it
 * already has once, crs → gdc — `docs/superpowers/specs/2026-09-19-service-account-migration-crs-to-
 * gdc-runbook.md`) there is exactly one line to change rather than a grep across the codebase.
 */
export const CURRENT_PROXY_ACCOUNT_EMAIL = "gdc@sdguthrie.com";

/**
 * The CURRENT proxy account's display name — same "one literal" reasoning as
 * `CURRENT_PROXY_ACCOUNT_EMAIL` above, and deliberately NOT derived from
 * `canonicalServiceAccountName("gdc")`: that lookup exists to RECOGNISE a service account among
 * many possible stored shapes, which is a different job from stating what the current one is
 * called. Used by the Audit Log's "Who" column (2026-09-22, client: every row should show this,
 * with no exceptions — see `AuditLog.tsx`'s own comment above the Who cell for the decision).
 */
export const CURRENT_PROXY_ACCOUNT_NAME = "Guthrie Document Centre";

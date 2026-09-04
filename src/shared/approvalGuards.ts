// The checks that must pass before a document is approved — in ONE place, for every route that approves.
//
// ── Why this module exists (2026-08-25) ─────────────────────────────────────────────────────────
// Until now these lived inside `ApprovalDocument.tsx`, which was fine while that page was the only
// way to approve. The client then asked for bulk approval from the library command bar, and a second
// approval path that did NOT run these checks would not be a convenience — it would be a faster way
// to do the exact damage they exist to prevent:
//
//   * approving into a unit folder that does not exist yet makes Auto-route create the whole chain
//     from the library root, INHERITING the library ACL — every user reads that unit's documents;
//   * approving over a same-named file used to REPLACE it silently — `Copy file` was
//     `nameConflictBehavior: 1` until 2026-08-25, when both routing flows were set to "Copy with a
//     new name". That is the backstop now; this check still runs, because catching a clash HERE
//     names the file and warns the approver, where the rename is silent and leaves two documents to
//     reconcile later;
//   * approving something you hold no `ApproveItems` on fails at the last step, after the requester
//     has been told it went through.
//
// So the rule is: anything that flips `OData__ModerationStatus` calls these first. There is no
// second implementation to drift.
//
// ⚠ THIS DOES NOT COVER SharePoint's OWN Approve/Reject command, and cannot. That is a native
// control setting the field directly, with no awareness any of this exists (see CLAUDE.md, "THE
// NATIVE SHAREPOINT APPROVE COMMAND BYPASSES BOTH OVERWRITE GUARDS"). The real fix for that is the
// same check inside Auto-route itself. Adding a bulk path here narrows the reason anyone would reach
// for the native command; it does not close it.
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { unitFolderPath, isProblem } from "./approvalDestination";
import { swapLibrarySegment } from "./hcRouting";

/** A check's verdict. `reason` reads as the tail of "…was NOT approved because <reason>". */
export type GuardResult = {
  ok: boolean;
  reason?: string;
  /**
   * ⚠ TRUE ONLY FOR A CONFIRMED COLLISION, and absent for every other refusal.
   *
   * `ok: false` covers two very different situations — "a document of that name IS there" and "the
   * check could not be answered". They were interchangeable while both simply refused; they stopped
   * being interchangeable on 2026-08-28, when the approval page began letting an approver proceed
   * over a clash (client: *"the overwritten is needed"*). Proceeding over an UNVERIFIED check is a
   * different decision from proceeding over a known one, and the approver has to be told which.
   *
   * Absent means "not a confirmed clash" — never "verified clean". Only `ok: true` means that.
   */
  clash?: boolean;
};

/** Whether the caller may approve one item. Three states — see `checkApproveRight`. */
export type ApproveRight = "granted" | "denied" | "unknown";

/**
 * Per-segment encoding with OData quote doubling.
 *
 * NOT `encodeURIComponent` over the whole path: that turns every `/` into `%2F`, and a deep path
 * built that way returns HTTP 400 rather than 404 (CLAUDE.md gotcha #9) — a malformed request that
 * reads exactly like "the folder is missing".
 */
const encodePath = (p: string): string =>
  p.split("/").map(encodeURIComponent).join("/").replace(/'/g, "''");

/**
 * Is the destination unit folder there, and locked down?
 *
 * FAILS CLOSED, including on an inconclusive read. A retry costs the approver seconds; a wrong
 * "proceed" publishes a unit's documents to everyone and nobody finds out.
 *
 * `permissionedTiers` comes from the segment's `Levels` chain and MUST NOT be derived from the
 * document's own fields: every tier column has a `<Base>Tid` twin, and a below-Unit tier such as
 * SubUnit carries one exactly like a permissioned tier, so counting those lands on a folder that
 * inherits and the guard refuses for a new reason. `undefined` means unknown, and unknown refuses.
 */
export async function checkUnitFolderReady(o: {
  sp: SPHttpClient;
  webUrl: string;
  /** The web's server-relative URL, e.g. `/sites/Example`. */
  webSru: string;
  /** The pending file's server-relative URL. */
  fileSru: string;
  /** URL segment of the library the file is in now, e.g. `ApprovalDocument`. */
  sourceSegment: string;
  /** URL segment of the library it will be routed to, e.g. `Shared Documents`. Blank ⇒ refuse. */
  destSegment: string;
  /** Display title of that destination library, for the message an approver acts on. */
  destLibTitle: string;
  permissionedTiers: number | undefined;
}): Promise<GuardResult> {
  const dest = unitFolderPath({
    fileSru: o.fileSru,
    webSru: o.webSru,
    sourceSegment: o.sourceSegment,
    destSegment: o.destSegment,
    permissionedTiers: o.permissionedTiers,
  });
  if (isProblem(dest)) return { ok: false, reason: dest.error };

  const res: SPHttpClientResponse = await o.sp.get(
    `${o.webUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields` +
      `?$select=HasUniqueRoleAssignments&@f='${encodePath(dest.path)}'`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );

  // 404 covers two cases SharePoint does not separate: the folder is absent, or the caller cannot
  // resolve it because they hold nothing on the destination library. Both refuse; name both, because
  // the fixes differ and an approver sent to the wrong one looks at a library that is working.
  if (res.status === 404) {
    return {
      ok: false,
      reason: `it does not exist yet, or your account has no access to the ${o.destLibTitle} library at all`,
    };
  }
  if (!res.ok) {
    return { ok: false, reason: `it could not be verified (HTTP ${res.status})` };
  }

  const d = await res.json().catch(() => undefined);
  // `{"odata.null": true}` means the folder RESOLVED but its list item was security-trimmed away.
  // That is a PASS, and the reasoning is worth stating because the obvious reading is the opposite:
  // a folder you cannot read the item of is one you do not have blanket rights over, which is the
  // state a locked-down unit folder puts a non-member in.
  if (!d || d["odata.null"] === true) return { ok: true };
  if (d.HasUniqueRoleAssignments === true) return { ok: true };
  if (d.HasUniqueRoleAssignments === false) {
    return { ok: false, reason: "it is not locked down — it still inherits the library's permissions" };
  }
  return { ok: false, reason: "its permissions could not be read from your account" };
}

/**
 * Would approving this REPLACE a document already sitting at the destination?
 *
 * The destination is the FULL file path with only the library segment swapped — never the unit
 * folder `checkUnitFolderReady` computes, which deliberately stops at Unit. Everything below Unit
 * (Year / Document Type / Archive…) must be preserved exactly, because that is what Auto-route's own
 * `Compose_1` expression preserves when it builds the copy destination.
 *
 * FAILS CLOSED on an unreadable answer, and deliberately UNLIKE the upload-time clash check in
 * `Form.tsx`, which fails open: that one runs on every upload site-wide, so an unanswerable read
 * there would take the whole form out of service. This runs once, at one approval.
 */
export async function checkDestinationClash(o: {
  sp: SPHttpClient;
  webUrl: string;
  fileSru: string;
  fileName: string;
  sourceSegment: string;
  /** Blank ⇒ refuse: a blank destination would otherwise mean "no clash" for an unresolved HC pair. */
  destSegment: string;
}): Promise<GuardResult> {
  const destFull = swapLibrarySegment(o.fileSru, o.sourceSegment, o.destSegment);
  if (!destFull) return { ok: false, reason: "its destination in the library could not be worked out" };

  const lastSlash = destFull.lastIndexOf("/");
  const destFolder = lastSlash > -1 ? destFull.slice(0, lastSlash) : destFull;

  const res: SPHttpClientResponse = await o.sp.get(
    `${o.webUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Files('${encodeURIComponent(o.fileName)}')` +
      `?$select=Exists&@f='${encodePath(destFolder)}'`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );

  /* 404 here is the NORMAL case and means something different from the 404 above. That one is about
     the UNIT folder; this is a below-Unit folder, ensure-created ON DEMAND by Auto-route's own
     `Create new folder` step, so it very often does not exist before the FIRST approval into a given
     Year / Document Type / Archive combination. Absent folder and absent file mean the same thing
     here: nothing to clash with. */
  if (res.status === 404) return { ok: true };
  if (!res.ok) {
    return { ok: false, reason: `it could not be checked for an existing document (HTTP ${res.status})` };
  }

  const d = await res.json().catch(() => undefined);
  if (!d || typeof d.Exists !== "boolean") {
    // A 200 that confirms neither way is exactly as uninformative as a failed read.
    return { ok: false, reason: "it could not be confirmed whether a document of that name is already there" };
  }
  if (d.Exists) {
    /* ⚠ `clash: true` MARKS THE ONE ANSWER THAT IS A CONFIRMED COLLISION. The three refusals above
       are "could not verify", which is a different thing, and since 2026-08-28 the approval page
       treats them differently: a confirmed clash warns that approving REPLACES the filed document
       and lets the approver proceed (client: *"the overwritten is needed"*), while an unverifiable
       check says so honestly rather than claiming a replacement that may not happen.

       Additive on purpose. `ok` is unchanged for every existing caller, so the bulk approve panel
       keeps SKIPPING a clashing file rather than gaining an override - see the note in
       `ApprovalDocument.tsx` for why bulk must stay strict. */
    return {
      ok: false,
      clash: true,
      reason: "a document of that name is already in the destination library, and approving would replace it",
    };
  }
  return { ok: true };
}

/**
 * May THIS caller approve THIS item?
 *
 * Asks the ITEM, never a role table: `EffectiveBasePermissions` accounts for the folder ACL,
 * inheritance and site admin in one read, with no need to know which folder the file is in or which
 * persona the caller holds. A role lookup can say "this persona approves" while reconciliation has
 * not actually granted it, or while the group was renamed; an ACL read cannot be wrong that way.
 *
 * `ApproveItems` is PermissionKind 5, so bit index 4. Read ARITHMETICALLY — JS bitwise coerces to a
 * signed 32-bit int, and Full Control returns `Low = "4294967295"`, which `&` gets wrong.
 *
 * THREE states, and callers must treat them differently: `denied` is a definite no, `unknown` is not
 * an answer at all. Hiding a control on `denied` is right; refusing to act on `unknown` is not,
 * because the write itself still fails safely with a 403.
 */
export async function checkApproveRight(o: {
  sp: SPHttpClient;
  webUrl: string;
  /** Display title of the library the item is in. */
  listTitle: string;
  itemId: number;
}): Promise<ApproveRight> {
  try {
    const res: SPHttpClientResponse = await o.sp.get(
      `${o.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(o.listTitle)}')` +
        `/items(${o.itemId})/EffectiveBasePermissions`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return "unknown";
    const d = await res.json();
    const low = Number(d.Low);
    if (!isFinite(low)) return "unknown";
    return Math.floor(low / Math.pow(2, 4)) % 2 === 1 ? "granted" : "denied";
  } catch {
    return "unknown";
  }
}

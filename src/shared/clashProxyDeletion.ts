import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { readFileETag } from "./deletionGuard";
import { primeNames } from "./spNaming";
import { cachedListTitle, LIST_SUFFIX } from "./naming";
import { encodeServerRelativePath } from "./dmsFolderMap";

/**
 * The clashing draft's identity, however much of it could be resolved.
 *
 * ⚠ EITHER MAY BE BLANK, NEVER BOTH — every caller refuses before reaching this function if neither
 * resolved. `uniqueId` is preferred (exact); `path` is the fallback for a colleague's HIDDEN pending
 * draft, whose own properties Draft Item Security trims away — see `deleteClashingDraftByProxy`'s own
 * comment for the one gap this still leaves.
 */
export interface ClashingDraftIdentity {
  uniqueId?: string;
  path?: string;
}

/**
 * Recycle a clashing PENDING draft BY PROXY, via `CRS — Execute approved deletion`, rather than
 * overwriting it directly.
 *
 * ⚠ EXTRACTED FROM `Form.tsx` (2026-09-23) SO BULK UPLOAD CAN SHARE IT, RATHER THAN GROWING A SECOND
 * COPY. This is exactly the risk `decideClash`'s own doc comment (`uploadBatches.ts`) already names
 * for the rule beside it — "a hand-written second copy is how the stricter of the two drifts, and the
 * copy that drifts is the one nobody exercises" — and it applies at least as strongly here: this
 * function decides whether a document is genuinely destroyed or safely proxied through an
 * approval-free deletion request, which is not a rule either upload screen should reimplement from
 * memory. Client, 2026-09-23: "it will follow the same flow as the normal document upload" — this is
 * that flow, shared rather than duplicated.
 *
 * WHY THIS EXISTS AT ALL: PIC/HoU no longer hold `Edit Items` on the approval libraries (the
 * 2026-09-18 tag/approve-by-proxy migration), so a direct `Files/Add(overwrite=true)` on an EXISTING
 * item 403s. The fix mirrors the 2026-09-17 proxy-deletion design already used for the ordinary delete
 * button: write a SELF-APPROVED `CRS Requests` row (`Status: "Approved"` from the moment it is
 * written, since nobody needs to decide it), and let `CRS — Execute approved deletion` — running as
 * the service account — do the actual `recycle()`.
 *
 * ⚠⚠ THIS IS A LIVE FLOW, CONFIRMED IN THE TENANT'S FLOW LIST (2026-09-19). So a poll that never
 * resolves is not explained by "the flow does not exist" — it means one of: (a) the service account
 * was never added to `CRS Owners` (runbook §0 prerequisite — the recycle then 403s), (b) the flow
 * inherited a stale trigger condition from a Save-As and never fires at all (runbook §1's own
 * warning), or (c) something in the condition/action wiring does not match what this code writes.
 * Check the `CRS Requests` row this call creates: `Status` staying `Approved` forever means the flow
 * never fired; `Status` flipping to `Failed` with a `DecisionNote` naming an HTTP code means it fired
 * and the recycle itself failed.
 *
 * Returns `true` once `GetFileById` on the old id starts 404ing (confirmed gone), `false` if the
 * deletion request itself could not be written, or if it never resolves within the poll budget below.
 */
export async function deleteClashingDraftByProxy(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  me: string,
  clashing: ClashingDraftIdentity,
  clashingName: string,
): Promise<boolean> {
  const now = new Date().toISOString();
  /* The baseline `CRS — Execute approved deletion` checks before recycling — the exact call site
     where the race was confirmed live, 2026-09-20 (three uploads of one filename sharing one
     `ItemUniqueId`, because a direct overwrite fallback reuses the SAME identity in place, and the
     proxy flow's own recycle landed a second after the newest upload). See `deletionGuard.ts` and
     `2026-09-20-etag-guarded-proxy-deletion-design.md`.
     `undefined` on any read failure or for the `hidden` no-GUID case, deliberately — a safety net for
     the FLOW's read, never a gate on the upload proceeding. */
  const targetEtag = clashing.uniqueId
    ? await readFileETag(spHttpClient, siteUrl, clashing.uniqueId)
    : undefined;
  try {
    await primeNames(spHttpClient, siteUrl);
    const listTitle = encodeURIComponent(
      cachedListTitle(LIST_SUFFIX.requests),
    );
    const itemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;
    const requestHeaders = {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "odata-version": "",
    };
    const body: Record<string, string> = {
      Title: `Replace on upload — ${clashingName}`.slice(0, 255),
      RequestType: "Deletion",
      Status: "Approved",
      // ⚠ EITHER MAY BE BLANK, NEVER BOTH — see `ClashingDraftIdentity`'s own comment. `ItemUrl` is
      // written for the `hidden`-clash case (a colleague's draft this account cannot resolve a GUID
      // for), but the flow as documented in `2026-09-17-execute-approved-deletion-flow-runbook.md`
      // §3 acts ONLY via `GetFileById(guid'...')` — it has no path-based fallback of its own. So a
      // `hidden` clash with a blank `ItemUniqueId` will make the flow's own recycle call malformed
      // (`GetFileById(guid'')`) and it will land in the Failure branch. Not the cause for an ordinary
      // self-replace (this account's own file always resolves a real GUID first), but a real,
      // separate gap for the colleague's-draft case — flag it if that specific scenario is what is
      // actually being tested.
      ItemUniqueId: clashing.uniqueId ?? "",
      ItemUrl: clashing.path ?? "",
      ItemName: clashingName,
      RequestedBy: me,
      RequestedAt: now,
      Reason: "",
      DecidedBy: me,
      DecidedAt: now,
      DecisionNote:
        "No approval needed — the uploader replaced this draft directly.",
    };
    // See `targetEtag`'s own comment above — absent on any read failure or the `hidden` case.
    if (targetEtag !== undefined) body.TargetETag = targetEtag;
    const send = (
      payload: Record<string, string>,
    ): Promise<SPHttpClientResponse> =>
      spHttpClient.post(itemsUrl, SPHttpClient.configurations.v1, {
        headers: requestHeaders,
        body: JSON.stringify(payload),
      });
    let res = await send(body);
    // ⚠ `TargetETag` IS THE NEWEST COLUMN, so it drops FIRST on a 400 — one unknown field name fails
    // the WHOLE write (gotcha #11), and this write has no other fallback: if it fails outright, the
    // caller falls straight to the risky direct overwrite this guard exists to make safe, which would
    // defeat the point on any site that has not yet added the column.
    if (res.status === 400 && body.TargetETag !== undefined) {
      const without = { ...body };
      delete without.TargetETag;
      res = await send(without);
    }
    if (!res.ok) return false;
  } catch {
    return false;
  }
  // Poll budget: ~10 attempts, 3s apart — about 30s. Bounded and short because it blocks the upload
  // button, unlike the tag/decision flows' own background ~1-minute poll.
  for (let attempt = 0; attempt < 10; attempt++) {
    await new Promise<void>((resolve) => setTimeout(resolve, 3000));
    try {
      // ⚠ THE GUID PROBE IS PREFERRED — it is exact, where a path can in principle be reused by
      // something else the instant it frees up. Falls back to the path only when no GUID was ever
      // resolved, which is exactly the `hidden`-clash case this whole fallback exists for.
      const check = clashing.uniqueId
        ? await spHttpClient.get(
            `${siteUrl}/_api/web/GetFileById(guid'${clashing.uniqueId}')?$select=Exists`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          )
        : await spHttpClient.get(
            // OData alias form, never an inline literal — gotcha #9: an inline path 400s once deep
            // enough, which reads as a malformed request rather than a missing file.
            `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)?$select=Exists&@f='${encodeServerRelativePath(clashing.path ?? "")}'`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
      if (check.status === 404) return true;
    } catch {
      // Transient — keep polling within the budget rather than giving up on one failed check.
    }
  }
  return false;
}

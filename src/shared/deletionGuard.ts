import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

/**
 * A tiny, single-purpose read: the current SharePoint `ETag` of one file, used to guard the proxy
 * deletion flow against acting on content that has changed since a deletion was approved.
 *
 * Spec: docs/superpowers/specs/2026-09-20-etag-guarded-proxy-deletion-design.md
 *
 * WHY THIS EXISTS: `CRS — Execute approved deletion` (Power Automate, running as the service
 * account) resolves its target purely by `ItemUniqueId`, and it is a POLLING trigger — its actual
 * recycle can run tens of seconds after the deletion was approved. If, in that window, the SAME
 * identity gets REUSED (`Files/Add(overwrite=true)` keeps the same `UniqueId` when overwriting in
 * place — confirmed of this project's own replace-clash fallback), the flow's eventual recycle
 * deletes whatever content is THERE NOW, not the content that was actually approved for deletion.
 * Confirmed live, 2026-09-20 — see `docs/2026-09-20-session-handoff.md`, items 23-28.
 *
 * The fix: stamp the target's ETag onto the `CRS Requests` row at the moment deletion is
 * authorised, and have the flow re-check it immediately before recycling — skip if it has changed.
 *
 * ⚠ FAILS OPEN, DELIBERATELY. A failed read here must never block writing the deletion request
 * itself — this is a safety net for the FLOW's own read, not a gate on raising a request.
 * `undefined` means "could not capture a baseline", which every call site treats identically to a
 * row written before this guard existed: no baseline, no guard, proceed as before.
 */
export async function readFileETag(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  uniqueId: string,
): Promise<string | undefined> {
  const id = (uniqueId ?? "").trim();
  if (id.length === 0) return undefined;
  try {
    const res: SPHttpClientResponse = await spHttpClient.get(
      `${siteUrl}/_api/web/GetFileById(guid'${id}')?$select=ETag`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Cache-Control": "no-cache",
          Pragma: "no-cache",
        },
      },
    );
    if (!res.ok) return undefined;
    const body = (await res.json()) as { ETag?: string };
    const tag = typeof body.ETag === "string" ? body.ETag.trim() : "";
    return tag.length > 0 ? tag : undefined;
  } catch {
    return undefined;
  }
}

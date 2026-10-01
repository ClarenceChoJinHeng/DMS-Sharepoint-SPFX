import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AbbrevRowDraft } from "./abbreviationDraft";
import { abbrevListTitle } from "./folderAbbreviation";

/**
 * PERMISSIONED-tier abbreviation rows for one segment (Department/Unit, or whatever this segment
 * calls them) — walked to exactly `levelCount` deep, joined with the segment's saved codes.
 *
 * Deliberately narrower than `AbbreviationManager.tsx`'s own tree-walking effect, which ALSO reads
 * coded below-Unit ("shared") levels. Nothing below Unit ever has a group mapped to it at all — no
 * group and no Group Map row is ever created that deep — so a caller building group-name CHAINS
 * (the only reason this function exists) has no use for them. Reusing the richer walk here would
 * mean pulling in that screen's "shared term set" handling for no benefit, and touching the most
 * heavily-scrutinised tree-walk in the project to serve an unrelated screen is the wrong risk.
 *
 * Returns `undefined` on ANY read failure — never a partial tree. A caller comparing a group's
 * CURRENT name against one computed from an incomplete chain would compute a WRONG expected name
 * and could offer to rename a group to something it should not be — "could not check" must never
 * be silently treated as "nothing to check".
 */
export async function loadPermissionedAbbreviationRows(
  context: WebPartContext,
  siteUrl: string,
  termSetGuid: string,
  levelCount: number,
): Promise<AbbrevRowDraft[] | undefined> {
  interface TermNode {
    id: string;
    label: string;
    parentId: string;
    depth: number; // 1-based
  }
  try {
    const nodes: TermNode[] = [];
    const walk = async (parentId: string, depth: number): Promise<void> => {
      if (depth > levelCount) return;
      const url = parentId
        ? `${siteUrl}/_api/v2.1/termStore/sets/${termSetGuid}/terms/${parentId}/children?$select=id,labels`
        : `${siteUrl}/_api/v2.1/termStore/sets/${termSetGuid}/children?$select=id,labels`;
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) throw new Error(`the term store returned HTTP ${res.status}`);
      const kids = (
        ((await res.json()).value ?? []) as Array<{
          id?: string;
          labels?: Array<{ name?: string }>;
        }>
      ).map((t) => ({
        id: t.id ?? "",
        label: (t.labels ?? [])[0]?.name ?? "",
        parentId,
        depth,
      }));
      for (const k of kids) {
        if (!k.id) continue;
        nodes.push(k);
        await walk(k.id, depth + 1);
      }
    };
    await walk("", 1);

    const existing: Record<string, string> = {};
    const list = encodeURIComponent(abbrevListTitle());
    const cur: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${list}')/items?$select=TermGuid,Abbreviation&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!cur.ok) throw new Error(`${abbrevListTitle()} returned HTTP ${cur.status}`);
    (
      ((await cur.json()).value ?? []) as Array<{ TermGuid?: string; Abbreviation?: string }>
    ).forEach((r) => {
      const key = (r.TermGuid ?? "").trim().toLowerCase();
      if (key) existing[key] = (r.Abbreviation ?? "").trim();
    });

    return nodes.map((n) => {
      const code = existing[n.id.toLowerCase()] ?? "";
      return {
        termGuid: n.id,
        label: n.label,
        level: `Level ${n.depth}`, // never read by `codeChain`/`plannedGroupRenames` — a placeholder
        parentGuid: n.parentId,
        abbreviation: code,
        original: code,
      };
    });
  } catch {
    return undefined;
  }
}

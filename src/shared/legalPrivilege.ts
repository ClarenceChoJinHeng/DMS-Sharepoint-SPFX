// Which confidentiality levels offer the Legally Privileged tick.
//
// ONE LEVEL WAS NOT ENOUGH (client, 2026-08-19: "we need the tickbox legally privillege to show for
// HC"). `legallyPrivilegedFor` was a single string compared with `===`, so naming Highly Confidential
// would have REMOVED the tick from Confidential — the setting could express "one level" and nothing
// else. It now reads as a list, and a single value behaves exactly as it always did.
//
// Pure, because it decides whether a legal marker is offered at all, and the failure is silent in both
// directions: a level that should offer it and does not loses the marker on every document filed there,
// while a level that should not offer it stamps a legal claim nobody authorised.
//
// FOUR call sites share this rule — the write derivation and the render guard, in both upload web parts
// — which is why it is here and not inline. The two derivations were already copies of each other, and
// BulkUpload's carried a comment saying "Mirrors Form.tsx": a rule maintained by hand in two places,
// deciding a legal marker.

/**
 * Split the `legallyPrivilegedFor` config value into levels.
 *
 * Semicolon OR comma separated, because an administrator typing a list into a text cell will use
 * whichever they think of first, and a value silently read as one long level name would offer the tick
 * nowhere — the same outcome as leaving it blank, while looking configured.
 *
 * Blank yields NO levels, which is the deliberate default: the tick appears only once a site asks for
 * it. A level named here that does not exist in the term set simply never matches.
 */
export function privilegedLevels(configured: string | undefined): string[] {
  const out: string[] = [];
  for (const part of (configured ?? "").split(/[;,]/)) {
    const v = part.trim();
    if (v.length > 0 && out.indexOf(v) === -1) out.push(v);
  }
  return out;
}

/**
 * Does this confidentiality level offer the Legally Privileged tick?
 *
 * Compared TRIMMED and CASE-INSENSITIVELY, unlike the old exact `===`. The value is typed into a
 * config cell by hand while reading a label off a term-store panel, so `Highly Confidential ` with a
 * trailing space, or `highly confidential`, would otherwise offer the tick nowhere while looking
 * correct — and nothing anywhere would report it.
 *
 * A blank level never matches, even against a blank entry in the list: `privilegedLevels` drops blanks,
 * so a document with no confidentiality set cannot inherit the marker.
 */
export function offersLegalPrivilege(
  level: string | undefined,
  configured: string | undefined,
): boolean {
  const l = (level ?? "").trim().toLowerCase();
  if (l.length === 0) return false;
  return privilegedLevels(configured).some((x) => x.toLowerCase() === l);
}

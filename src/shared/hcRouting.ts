/**
 * Highly Confidential routing — which library a confidentiality level files into.
 *
 * Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md
 *
 * Pure, SPFx-free and tested, because every rule here has a wrong version that looks identical on
 * screen. The whole feature is the difference between a document landing in the library its unit
 * reads and one landing in the library only cleared people read — and nothing in the UI shows which
 * happened. The file uploads, the toast is green, and the mistake surfaces when the wrong person
 * opens it weeks later.
 *
 * THE ONE RULE THIS MODULE EXISTS TO HOLD: when anything is unknown, Highly Confidential is NOT
 * OFFERED. That is the opposite of this codebase's usual habit, where an unreadable list means "show
 * everything" so a transient error cannot take a form out of service (gotchas 10b and 11, the
 * provisioned-segment filters). The cost of being wrong is what flips it: an over-hidden segment
 * blocks an upload someone retries a minute later, while an over-offered HC level publishes a secret
 * and nobody finds out.
 */

/** The four libraries, as logical keys. `Staging` is historical and stays — see naming.ts. */
export type LibKey = "Staging" | "Documents" | "StagingHC" | "DocumentsHC";

/** What the caller knows when it asks where a document should go. */
export interface RoutingContext {
  /**
   * The level label that routes to HC, from the `hcConfidentialityLevel` config row.
   *
   * BLANK MEANS NO LEVEL EVER ROUTES TO HC — the correct default for the many sites that will never
   * have HC libraries, and the reason this is not "undefined means Highly Confidential". A site that
   * has not configured HC has not opted into it.
   */
  hcLevel: string;
  /** Whether BOTH HC libraries resolved. False covers "absent" and "not probed yet" alike. */
  hcAvailable: boolean;
  /**
   * Whether this user can actually WRITE into their HC folder — the answer to a probe, not to
   * "are they in the group". Undefined means the probe was inconclusive, which counts as no.
   */
  canWriteHc?: boolean;
}

function norm(v: string | undefined): string {
  return (v ?? "").trim().toLowerCase();
}

/** The label assumed when a site has HC libraries but has not named the level. */
export const DEFAULT_HC_LEVEL = "Highly Confidential";

/**
 * The level that actually routes, given the config row and whether the libraries exist.
 *
 * THE LIBRARIES DECIDE WHETHER HC ROUTING IS ON AT ALL; the config row only renames the level. Both
 * simpler rules are wrong, in opposite and instructive ways:
 *
 * - Defaulting to `Highly Confidential` unconditionally would HIDE that level on every site that has
 *   not set up HC — and there it is an ordinary metadata label any PIC has always been free to
 *   choose. A feature nobody enabled would quietly remove an existing option.
 * - Defaulting to blank unconditionally would mean a site that created the libraries but forgot the
 *   config row files Highly Confidential documents into the NORMAL library. That is the exact
 *   failure this design exists to prevent, arriving through a setting nobody knew they had to write.
 *
 * So: no libraries, no routing, and the level stays the plain label it has always been. Libraries
 * present, routing on, with a sane default the client can rename.
 */
export function effectiveHcLevel(configured: string | undefined, libsAvailable: boolean): string {
  const v = (configured ?? "").trim();
  if (v.length > 0) return v;
  return libsAvailable ? DEFAULT_HC_LEVEL : "";
}

/**
 * Is this the level that routes to Highly Confidential?
 *
 * Compared case-insensitively and trimmed, because the label is typed into a config row by an
 * administrator reading it off a term-store panel, and `Highly Confidential ` with a trailing space
 * would otherwise route nothing while looking correct.
 *
 * A blank configured level matches NOTHING, including a blank level on the document. Without that
 * guard a document with no confidentiality set would route to HC on every site that has not
 * configured the row — which is every site, today.
 */
export function isHcLevel(level: string | undefined, ctx: RoutingContext): boolean {
  const configured = norm(ctx?.hcLevel);
  if (configured.length === 0) return false;
  return norm(level) === configured;
}

/**
 * Where does a document at this confidentiality level go?
 *
 * `approval` is the library the file is uploaded into; `documents` is where auto-route moves it after
 * approval. Returned together so no caller can pick one half of a pair — the same reasoning as
 * resolving a library's title and URL segment together.
 *
 * FALLS BACK TO THE NORMAL PAIR ONLY WHEN THE LEVEL IS NOT HC. An HC level with no available HC
 * library does NOT fall back: `canOfferHc` keeps that level off the form in the first place, and a
 * caller that reaches here anyway gets `undefined` rather than a normal-library destination. Falling
 * back is the single worst thing this module could do.
 */
export function routeFor(
  level: string | undefined,
  ctx: RoutingContext,
): { approval: LibKey; documents: LibKey } | undefined {
  if (!isHcLevel(level, ctx)) return { approval: "Staging", documents: "Documents" };
  if (!ctx.hcAvailable) return undefined;
  return { approval: "StagingHC", documents: "DocumentsHC" };
}

/**
 * May this user be offered the Highly Confidential level at all?
 *
 * THREE conditions, all of which must be positively true:
 *   1. a level is configured — otherwise the site has not opted into HC;
 *   2. both HC libraries resolved — otherwise there is nowhere to file it;
 *   3. the write probe said yes — not "they are in the group".
 *
 * Condition 3 is the one that is easy to get wrong. Holding `GHO_GF_CORU_UPL_HC` and being able to
 * write into that unit's HC folder are different questions: reconciliation grants folder ACLs in a
 * separate pass from group creation, so a group can exist for days before it grants anything. That
 * gap is the origin of the upload form's original HTTP 403, and a membership check would reproduce it
 * here — except that here the symptom is not a refused upload but an offered level that files nowhere.
 */
export function canOfferHc(ctx: RoutingContext): boolean {
  if (norm(ctx?.hcLevel).length === 0) return false;
  if (!ctx.hcAvailable) return false;
  return ctx.canWriteHc === true;
}

/**
 * The confidentiality levels this user may choose.
 *
 * Everything except the HC level, plus the HC level when `canOfferHc`. Order is preserved: these come
 * from the term set in the order the client arranged them, and re-sorting would move a level
 * somewhere nobody expects.
 *
 * HIDDEN, NOT DISABLED — the client's choice. A greyed-out "Highly Confidential" tells everyone that
 * the level exists and that some documents in their unit are filed under it, which is information an
 * uncleared uploader has no need for.
 */
export function selectableLevels(levels: readonly string[], ctx: RoutingContext): string[] {
  const offer = canOfferHc(ctx);
  return (levels ?? []).filter((l) => !isHcLevel(l, ctx) || offer);
}

/**
 * Was an HC level chosen that this user cannot actually file?
 *
 * The last-moment check, run immediately before the write. A stale page is the case it exists for:
 * settings and clearances are read once at mount, so a tab left open across a clearance change would
 * otherwise submit against the old answer — the same failure as gotcha 10b, where a stale chain filed
 * into the old folder shape and nothing looked wrong.
 *
 * Both messages say NOTHING HAS BEEN UPLOADED, because the alternative reading — that a Highly
 * Confidential file is now sitting somewhere unknown — is the one that sends someone hunting through
 * libraries they should not be opening.
 *
 * Returns the message to show, or undefined when the upload may proceed.
 */
export function refuseReason(level: string | undefined, ctx: RoutingContext): string | undefined {
  if (!isHcLevel(level, ctx)) return undefined;
  if (!ctx.hcAvailable) {
    return (
      "Highly Confidential documents cannot be filed on this site — the Highly Confidential library " +
      "was not found. Nothing has been uploaded. Tell an administrator before trying again."
    );
  }
  if (ctx.canWriteHc !== true) {
    return (
      "You are not cleared to file Highly Confidential documents in this unit, or your access has " +
      "changed since this page was opened. Nothing has been uploaded. Reload the page, and ask an " +
      "administrator if it still refuses."
    );
  }
  return undefined;
}

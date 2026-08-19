// Which business segments a reconciliation run covers.
//
// Spec: docs/superpowers/specs/2026-08-19-selective-reconciliation-design.md (register #18)
//
// The client, 2026-08-19: *"if client only added new subunit or new unit under two business segment's
// department and they have 5 Business segment in the future, this will slow them down if folder recon
// have to go through everything again… we ask client which business segment they want to run."*
//
// Measured on the rehearsal site: a check run over ONE segment is 4,552 log lines and about forty
// minutes. Adding one unit costs that same forty minutes to change one folder in four libraries, and
// the one line that matters sits among four thousand saying `already there`. At five segments it is
// over three hours — and the client stops running it, which is worse than slow, because
// reconciliation is the step that makes access real.
//
// ⚠ THIS MODULE DECIDES ONLY WHAT THE ADMIN ASKED FOR. It does NOT decide what "needs" running.
// Detecting changed segments is deliberately out of scope here (spec §2): detection compares INPUTS
// — terms, abbreviation rows, map rows — while reconciliation asserts OUTCOMES: folder ACLs, broken
// inheritance, content types, page grants. An ACL someone removed by hand changes no term and no row,
// and that is exactly the failure reconciliation exists to repair. So a "nothing changed" verdict
// would be a statement about the lists, never about the site.
//
// Pure, because the alternative to getting this right is a segment silently excluded from a run —
// folders unbuilt, ACLs ungranted, uploads refused, and a green log saying nothing needed doing.

/** The fields this module needs from a reconciliation mode row. */
export interface ScopeSegment {
  /** The mode row's key — `mode_gho`. Stable, and what a selection is stored against. */
  key: string;
  /** The segment's top folder name — `GHO`. */
  stagingFolder: string;
  /** Display label, when the mode row carries one. Falls back to `stagingFolder`. */
  label?: string;
}

export interface RunScope {
  /** The segments to run, in the order given. Empty only when `refused`. */
  segments: ScopeSegment[];
  /**
   * Why the run may not proceed, or `undefined` when it may. An empty tick list is a mis-click, not
   * an instruction to do nothing — so it is refused with a reason rather than running over zero
   * segments and reporting success.
   */
  refused?: string;
  /**
   * One line naming the coverage, for the run log and the audit row: `all 5 segments`, or
   * `2 of 5 segments (Group Head Office, Minamas Head Office)`.
   *
   * A run record that does not say what it covered is unreadable a month later — somebody will
   * compare a two-segment run against a five-segment one and conclude something broke.
   */
  label: string;
}

const nameOf = (s: ScopeSegment): string => ((s.label ?? "").trim() || s.stagingFolder);

/**
 * Resolve the run's scope from what the administrator ticked.
 *
 * `selected` is the set of mode keys. `undefined` means "not chosen yet" and is treated as ALL — the
 * safe default is today's behaviour, and a default of "only what changed" would make the first run
 * after an unseen manual edit skip the one segment that needed it.
 *
 * A ticked key matching no mode row is dropped silently: it means a segment was retired between the
 * page loading and Run being pressed, and refusing the whole run over it would be worse than covering
 * one segment fewer. If that leaves nothing, the empty-selection refusal catches it.
 */
export function resolveRunScope(
  all: ScopeSegment[] | null | undefined,
  selected: Set<string> | null | undefined,
): RunScope {
  const every = (all ?? []).filter((s) => ((s ?? {}) as ScopeSegment).key !== undefined)
    .filter((s) => (s.key ?? "").trim().length > 0);
  if (every.length === 0) {
    return { segments: [], refused: "No business segments are configured on this site.", label: "no segments" };
  }
  const plural = (n: number): string => `all ${n} segment${n === 1 ? "" : "s"}`;
  if (!selected) return { segments: every, label: plural(every.length) };

  const picked = every.filter((s) => selected.has(s.key));
  if (picked.length === 0) {
    return { segments: [], refused: "Tick at least one business segment to run.", label: "no segments" };
  }
  if (picked.length === every.length) return { segments: picked, label: plural(every.length) };
  return {
    segments: picked,
    label: `${picked.length} of ${every.length} segments (${picked.map(nameOf).join(", ")})`,
  };
}

/**
 * The passes that must run in FULL regardless of scope.
 *
 * ⚠ SITE ENTRY, LIBRARY STATE, PAGE ACCESS, THE ADMIN-PAGE LOCKDOWN AND THE HC GATING CHECK ASSERT
 * STATE THAT HAS NO SEGMENT. Scoping them to the ticked segments would recreate precisely the class
 * of bug this codebase has now found four times — a mechanism driven by grant rows, and the thing
 * needing protection having none. Their cost is a handful of requests; their absence is an exposed
 * library or a locked-out administrator.
 *
 * Exported as documentation with a test behind it, so the list cannot quietly shrink.
 */
export const ALWAYS_FULL_PASSES: string[] = [
  "site entry",
  "library state",
  "page access",
  "admin page lockdown",
  "HC gating check",
];

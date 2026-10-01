// Highly Confidential CLEARANCE — the one implementation, shared by both upload screens.
//
// Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md
//
// ⚠ THIS EXISTS BECAUSE TWO COPIES DRIFTED AND BOTH BROKE (2026-08-27). The upload form owned the
// probe and Bulk Upload had none; the form gained an `hcReady` re-render trigger on 2026-08-22 and
// Bulk did not, so the Highly Confidential level appeared intermittently on one screen and for the
// wrong reason on the other. The client's report was a single sentence covering both. A second copy
// of this logic will drift again — mount this instead.
import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  lookupFolderMapping,
  resolveMappedFolder,
  probeFolderUploadAccessByPath,
} from "./dmsFolderMap";
import { cachedHcLibraries, hcAvailable, libraryUrlSegment } from "./naming";
import { swapLibrarySegment } from "./hcRouting";

/** Genuine failed attempts per unit before giving up. RENDERS ARE NOT ATTEMPTS — see below. */
const HC_PROBE_ATTEMPTS = 3;

/**
 * Which units this person can actually file Highly Confidential documents into.
 *
 * CLEARANCE IS A WRITE PROBE, NEVER GROUP MEMBERSHIP. Reconciliation grants folder ACLs in a pass
 * separate from group creation, so a `..._UPL_HIGHLY_CONFIDENTIAL` group can exist for days before it
 * grants anything. Asking the folder is the only question whose answer cannot be stale.
 *
 * ⚠ FAILS CLOSED. An absent, denied or unanswerable verdict leaves the unit out, so the level is not
 * offered. Against this codebase's usual fail-open habit, and deliberately: an over-offered HC level
 * publishes a secret, where an over-hidden one costs an upload somebody retries.
 *
 * ⚠ ATTEMPTS ARE COUNTED ON GENUINE FAILURES ONLY, and counting renders instead was the 2026-08-27
 * bug. The caller's `leafTerm` changes with every cascade step, so a render-counted budget was spent
 * in milliseconds whether or not any probe had finished — and the leaf was then denied for the life of
 * the page, telling a cleared uploader they were not cleared.
 *
 * ⚠ `hcReady` IS REQUIRED, not optional. `hcAvailable()` reads a MODULE CACHE that React cannot
 * observe, and the HC pair resolves asynchronously at mount — so without a state flag threaded in, an
 * effect that ran before priming returned early and never ran again.
 */
export function useHcClearance(
  sp: SPHttpClient,
  siteUrl: string,
  leafTerm: string,
  /** True once `primeNames` has resolved. Ties this to an async fact React cannot see. */
  hcReady: boolean,
): Record<string, boolean> {
  const [cleared, setCleared] = React.useState<Record<string, boolean>>({});
  /** Genuine failures per leaf. */
  const failures = React.useRef<Record<string, number>>({});
  /** Leaves with a probe in flight, so a re-render reuses it rather than starting another. */
  const inFlight = React.useRef<Record<string, boolean>>({});

  React.useEffect(() => {
    const leaf = leafTerm;
    if (!hcAvailable() || !leaf) return;
    if (cleared[leaf] === true) return;            // answered yes; nothing to re-ask
    if (inFlight.current[leaf] === true) return;   // already running; its result re-renders us
    if ((failures.current[leaf] ?? 0) >= HC_PROBE_ATTEMPTS) return;

    /** Records ONE genuine failure. Called from every bail path and nowhere else. */
    const failed = (): void => {
      failures.current[leaf] = (failures.current[leaf] ?? 0) + 1;
    };

    const ask = async (): Promise<void> => {
      const mapping = await lookupFolderMapping(sp, siteUrl, leaf).catch((e) => {
        console.warn("HC probe: folder-map lookup failed for term", leaf, e);
        return null;
      });
      if (!mapping?.folderUniqueId) {
        console.warn("HC probe: no Folder Map row (or no folder id) for term", leaf);
        failed();
        return;
      }
      const unit = await resolveMappedFolder(sp, siteUrl, mapping.folderUniqueId, mapping.folderUrl);
      /* The HC twin of a folder already resolved. The trees mirror each other by design, so there is
         no second Folder Map to read — but the path is only where the folder WOULD be, and the probe
         is what decides whether it exists and is writable. */
      const hcPath = swapLibrarySegment(
        unit.serverRelativeUrl, libraryUrlSegment(), cachedHcLibraries()?.approval.urlSegment,
      );
      if (!hcPath) {
        console.warn(
          "HC probe: could not mirror the unit path into the HC library.",
          "unit:", unit.serverRelativeUrl,
          "from:", libraryUrlSegment(),
          "to:", cachedHcLibraries()?.approval.urlSegment,
        );
        failed();
        return;
      }
      const verdict = await probeFolderUploadAccessByPath(sp, siteUrl, hcPath);
      if (verdict === "granted" || verdict === "denied" || verdict === "missing") {
        // Conclusive: settled for this page, whichever way it went.
        failures.current[leaf] = HC_PROBE_ATTEMPTS;
      } else {
        failed(); // `unknown` — a throttle or a slow paint. One attempt, not the budget.
      }
      if (verdict !== "granted") console.warn("HC probe:", verdict, "for", hcPath);
      if (verdict === "granted") setCleared((c) => ({ ...c, [leaf]: true }));
    };

    inFlight.current[leaf] = true;
    ask()
      .catch((e) => { console.warn("HC probe threw", e); failed(); })
      .then(() => { inFlight.current[leaf] = false; })
      .catch(() => undefined);
  }, [leafTerm, hcReady, cleared, sp, siteUrl]);

  return cleared;
}

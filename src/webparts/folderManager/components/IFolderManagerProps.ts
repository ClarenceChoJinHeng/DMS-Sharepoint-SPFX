import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFolderManagerProps {
  context: WebPartContext;
  /**
   * Which tab to open on, as a deep-link slug (`abbreviations`, `structure`, `migrate`,
   * `reconciliation`, `newsegment`).
   *
   * Added 2026-08-14 so the guided flows in `FolderAdmin` can drive this component one step at a time
   * without reconciliation being extracted from it. Absent = the URL hash decides, then Term
   * Abbreviations. See `DEEP_LINK_TABS` in FolderManager for the slug contract.
   *
   * Ignored by `FolderMap`, which shares this interface.
   */
  initialTab?: string;
  /**
   * Hide the tab bar and the page heading.
   *
   * Set when a flow is showing one step: the flow's own rail is the navigation, and a second row of tabs
   * beside it would offer a way out of the flow that looks like part of it.
   */
  hideTabs?: boolean;
  /**
   * Fired when the New segment tab actually writes a segment, with its derived key and label.
   *
   * Passed straight through to `SegmentCreator`. It exists so the guided flow can close the creation
   * form and select the new segment on the creation ITSELF, rather than inferring it from a re-read plus
   * a question to the admin — which is why the form used to sit open underneath its own success message.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onSegmentCreated?: (key: string, label: string) => void;
  /**
   * Hide **Delete** on the New segment tab's list of existing segments.
   *
   * Set by the "Add a new segment" flow only. NOT keyed off `hideTabs`, because the **Retire** flow mounts
   * this same tab with `hideTabs` set and Delete is the whole point of that step — so the two cases must
   * be told apart explicitly. The list itself is never hidden: seeing what exists is what stops someone
   * creating a segment twice.
   */
  hideSegmentDelete?: boolean;
  /**
   * Passed through to `AbbreviationManager`: how many terms have no folder code, or `undefined` when that
   * is not knowable.
   *
   * The guided flow gates its Next button on this and cannot work it out itself — the count needs a walk
   * of the whole term tree (~115 requests for GHO), which the abbreviations screen has already paid for.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onAbbreviationsMissingChange?: (missing: number | undefined) => void;
  /**
   * True while the abbreviations screen is reading — the segment list or the term tree.
   *
   * The count alone cannot answer this: it is `undefined` both while loading and when the read
   * failed, and only the first should hold the flow's Next button. Ignored by `FolderMap`, which
   * shares this interface.
   */
  onAbbreviationsLoadingChange?: (loading: boolean) => void;
  /**
   * True while RECONCILIATION is running, so a host can hold its own navigation.
   *
   * The run executes entirely in this page — there is no server-side job — so a step change unmounts
   * it and stops it mid-operation. `beforeunload` already guards the tab; this guards the guided
   * flow's rail, which sits right beside a run that takes an hour at the client's scale. Same
   * report-upward shape as `onBusyChange` on the bulk group provisioner, and it drives the SAME
   * padlock in FolderAdmin — one reason to hold navigation, one implementation of holding it.
   */
  onReconRunningChange?: (running: boolean) => void;
  /**
   * True while the subtree MIGRATION screen is scanning or moving.
   *
   * Same reason as `onReconRunningChange`: the work happens entirely in this page, so a step change
   * unmounts it and stops it part-way. The scan is the case the client hit — Next stayed live while
   * the button read "Checking…", and pressing it would have thrown the scan away silently.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onMigrateRunningChange?: (running: boolean) => void;
  /**
   * True while the Folder levels screen holds UNSAVED edits.
   *
   * `FolderManager` already tracks this to block its own tab switches; surfacing it lets the guided
   * flow hold Next for the same reason. Walking to the migrate step with an unsaved chain is worse
   * than a tab switch: the next step reads `PendingLevels`, which the unsaved edit has not written,
   * so the flow reports "nothing to move" for a change the admin believes they made.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onStructureDirtyChange?: (dirty: boolean) => void;
  /**
   * Fired after the Folder levels screen SAVES a structure change.
   *
   * ⚠ THE GUIDED FLOW MUST RE-READ ITS SEGMENT LIST ON THIS, or the very next step insists there is
   * no pending change to move to — about the change just made. That flow locks its migrate step on
   * `pendingLevels`, which it copies out of a segment list read once at mount; re-reading the
   * DERIVED facts on step change (which it already does) cannot help, because the stale value lives
   * in the upstream read. Found live 2026-09-08.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onStructureSaved?: () => void;
  /**
   * True when the migration screen has SCANNED work that has not been run yet.
   *
   * Walking past that step with moves outstanding leaves the segment half-changed — `PendingLevels`
   * still set, folders still in the old shape — and the step after it turns uploads back on over
   * exactly that state. False when a scan found nothing, which is a legitimate end state.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onMigratePendingChange?: (pending: boolean) => void;
  /**
   * The staged chain has been switched on.
   *
   * ⚠ The guided flow's `migrate` gate reads `pendingLevels` off a segment list read earlier, so
   * without this it would hold Next for ever after a successful migration. Mirrors
   * `onStructureSaved`: the host re-reads the list that owns the fact.
   */
  onMigrateApplied?: () => void;
  /**
   * A migration CHECK has finished — whatever it found, and whether or not it succeeded.
   *
   * ⚠ FINISHED, NOT SUCCEEDED, and the distinction matters here more than usual: in the structure
   * flow this step sits before the one that turns uploads back ON, so a success requirement would
   * keep the site refusing documents for as long as the scan kept failing. Same rule as
   * `onReconRunningChange`.
   */
  onMigrateScanned?: () => void;
  /**
   * Pre-selects the migration screen's segment picker, so a flow that already asked does not ask
   * again. Client, 2026-08-20: *"After selecting I got to select again which is weird."*
   * Matches on the mode row's `Title`, the key BOTH screens build their options from.
   */
  migrateInitialSegmentKey?: string;
  /**
   * Whether uploads are paused site-wide, passed through to the migration screen's warning banner.
   *
   * ⚠ ONLY `true` SUPPRESSES IT. The flow knows this already — it is the fact `blocksNext` gates step
   * 1 on — while the migration screen has no reader of its own, so its banner was rendered
   * unconditionally and went on demanding uploads be turned off after they had been (reported on
   * site 2026-09-07). Omitted from the standalone Migrate tab on purpose: no step 1 exists there, so
   * the warning is the only thing that says it.
   */
  migrateUploadsPaused?: boolean;
  /**
   * Same fix as `migrateInitialSegmentKey`, extended to the Abbreviations tab — found live 2026-08-26
   * on the "Rename or re-code a folder" flow: the flow already asks which segment (it's in the step
   * header), but the Abbreviations screen still opened on its own blank "Select a segment..." and made
   * the admin pick again. Every flow that carries `needsSegment: true` and includes an ABBREVIATIONS
   * step should pass this.
   */
  abbreviationsInitialSegmentKey?: string;
  /**
   * True while the Abbreviations screen holds UNSAVED edits — same shape as `onStructureDirtyChange`,
   * extended to the sibling screen it was never applied to. Found live 2026-08-26: an admin can type a
   * new code, then click Next straight into Folder Reconciliation without saving. Reconciliation reads
   * the SAVED rows, so it runs clean and reports success while doing nothing for the unsaved edit — the
   * admin believes the rename happened, and the client is the one who finds out it did not.
   */
  onAbbreviationsDirtyChange?: (dirty: boolean) => void;
  /**
   * Hands the abbreviation screen's own `save()` up to a guided flow, so **Next** can save before
   * advancing (client, 2026-09-06 — that screen's Save button is gone).
   *
   * Passed straight through to `AbbreviationManager.registerSave`; see the note there for why the
   * host must call THAT function rather than write the rows itself.
   */
  onAbbreviationsRegisterSave?: (save: (() => Promise<boolean>) | undefined) => void;
  /**
   * Hides the create-a-segment form on the NewSegment tab, leaving only the segments list and its
   * Delete buttons. For the Retire flow's "Segments -> Delete" step, which has no use for it — client,
   * 2026-08-26: "Retiring a segment is just to delete, not to create." Mirrors `hideSegmentDelete`,
   * inverted; defaults to shown so the standalone Segments tab and "Add a new segment" are unaffected.
   */
  hideSegmentCreate?: boolean;
  /**
   * Passed through to `SegmentCreator`'s `allowRecode`, inverted — hides **Re-code folder** from the
   * segments list. Set true only by the Retire flow (client, 2026-09-13: found it sitting on the
   * retire screen and asked why it was there — it was never a deliberate choice, just a gap between
   * Re-code being built and Retire's own "hide everything but delete" wiring already existing).
   * Defaults to shown, mirroring `hideSegmentDelete`/`hideSegmentCreate`, so the standalone Segments
   * tab and "Add a new segment" are unaffected.
   */
  hideSegmentRecode?: boolean;
}

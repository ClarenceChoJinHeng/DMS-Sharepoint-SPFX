import * as React from "react";

/**
 * The one definition of "clicking the dark area behind a dialog closes it".
 *
 * ⚠⚠ IT MUST BE WIRED TO `onMouseDown`, NEVER `onClick`, AND THAT IS THE WHOLE POINT OF THIS FILE.
 *
 * Every dialog in this codebase used to close on the backdrop's `onClick`, with
 * `onClick={(e) => e.stopPropagation()}` on the dialog itself to stop a click INSIDE it bubbling
 * out. That looks airtight and is not: **a browser dispatches `click` on the nearest COMMON ANCESTOR
 * of where the mouse went down and where it came up.** So selecting text — press inside the dialog,
 * drag, release a few pixels outside it — dispatches the click on the BACKDROP. The dialog's
 * `stopPropagation` never runs, because the event never passes through the dialog at all, and the
 * dialog shuts.
 *
 * Reported on the share-request dialog (2026-09-07): *"Each time I highlight something on the popup
 * it closes."* Highlighting a reason to re-read it, or dragging across an email address to correct
 * it, is ordinary behaviour — and on the two dialogs that hold typed text (the approver's rejection
 * note, and the typed `DELETE` confirmation on bulk group delete) it silently destroyed that text.
 *
 * `onMouseDown` cannot be fooled the same way: a drag that starts inside the dialog has its mousedown
 * INSIDE the dialog, so `e.target === e.currentTarget` is false and nothing closes. No ref, no state,
 * no timing window.
 *
 * `e.target === e.currentTarget` is doing the work rather than `stopPropagation`, so it also survives
 * a caller who forgets the inner handler — the failure mode being guarded against here is precisely
 * a handler that was assumed to run and did not.
 *
 * Cost, accepted: the dialog now closes on PRESS rather than release, so pressing the backdrop and
 * dragging back into the dialog still closes it. That gesture has no meaning; the one being protected
 * — selecting text — does.
 *
 * @param onClose Called only when the press landed on the backdrop itself. Callers gate their own
 *   "a write is in flight" checks inside this, exactly as they did with the old `onClick`.
 */
export function closeOnBackdrop(
  onClose: () => void,
): (e: React.MouseEvent<HTMLDivElement>) => void {
  return (e) => {
    if (e.target === e.currentTarget) onClose();
  };
}

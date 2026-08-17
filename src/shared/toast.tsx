// A floating result message, for admin screens taller than the viewport.
//
// Client, 2026-08-17: *"This success message, can you instead show a toast? Client is not going to go up
// and they wont know if it is succeeded."* The abbreviations page runs to 70-odd rows, so a banner at the
// top is off-screen at the moment it is written — the admin presses Save, nothing appears to happen, and
// the reasonable conclusion is that it failed. The same reasoning moved the New segment form's collision
// errors down onto their own fields.
//
// FIXED to the viewport, so scroll position cannot hide it.
//
// Note there is an older component-local toast inside BulkUpload.tsx. It is not reused here and this does
// not replace it: that file is over the lint ceiling and sits on a site-critical upload path, so moving it
// is its own reviewable change. It also auto-dismisses errors, which is the one behaviour deliberately
// not copied — see below.
import * as React from "react";
import { useEffect } from "react";

export type ToastKind = "ok" | "err" | "warn";

/** How long a SUCCESS stays up. Long enough to read a sentence, short enough to stay out of the way. */
const OK_MS = 6000;

const base: React.CSSProperties = {
  position: "fixed",
  zIndex: 1000000, // above SharePoint's own command bars and panels
  right: 20,
  bottom: 20,
  maxWidth: 460,
  boxSizing: "border-box",
  padding: "12px 14px",
  borderRadius: 8,
  fontFamily: '"Segoe UI", system-ui, sans-serif',
  fontSize: 13,
  lineHeight: 1.5,
  boxShadow: "0 6px 18px rgba(0,0,0,.18)",
  display: "flex",
  alignItems: "flex-start",
  gap: 10,
};

const tone: Record<ToastKind, React.CSSProperties> = {
  ok: { background: "#f1f8f4", border: "1px solid #9fd3b5", color: "#0f6c3f" },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { background: "#fff4e5", border: "1px solid #f0d9b5", color: "#7a4f00" },
};

const closeBtn: React.CSSProperties = {
  border: "none",
  background: "transparent",
  color: "inherit",
  font: "inherit",
  fontWeight: 700,
  cursor: "pointer",
  padding: "0 2px",
  lineHeight: 1.2,
};

export interface ToastProps {
  kind: ToastKind;
  text: string;
  /** Called when it dismisses itself, or when the reader closes it. */
  onDismiss: () => void;
}

/**
 * One floating message. Render it only while there is something to say.
 *
 * **A SUCCESS AUTO-DISMISSES; A FAILURE AND A WARNING DO NOT.** That asymmetry is the point: nobody needs
 * to keep reading "saved", but a message saying the save failed — or half-succeeded, which is exactly what
 * the abbreviations page reports when the rows were written and the audit row was not — must still be on
 * screen when the admin looks back. A toast that clears itself after an error is WORSE than the banner it
 * replaces, because it can be missed entirely rather than merely be out of view.
 *
 * `aria-live` is `assertive` for a failure and `polite` otherwise, so a screen reader interrupts only for
 * the one that changes what the user should do next.
 */
export function Toast({ kind, text, onDismiss }: ToastProps): React.ReactElement {
  useEffect(() => {
    if (kind !== "ok") return undefined;
    const t = setTimeout(onDismiss, OK_MS);
    return () => clearTimeout(t);
    // `text` is in the deps so a second success restarts the clock instead of inheriting the first's
    // remaining time — two saves in quick succession would otherwise flash the second one away.
  }, [kind, text, onDismiss]);

  return (
    <div
      style={{ ...base, ...tone[kind] }}
      role={kind === "err" ? "alert" : "status"}
      aria-live={kind === "err" ? "assertive" : "polite"}
    >
      <span style={{ flex: 1, minWidth: 0 }}>{text}</span>
      <button style={closeBtn} onClick={onDismiss} title="Dismiss" aria-label="Dismiss">
        ✕
      </button>
    </div>
  );
}

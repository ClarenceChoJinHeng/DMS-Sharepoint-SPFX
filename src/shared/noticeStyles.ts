/**
 * The palette every ATTENTION banner uses — one definition, for all of them.
 *
 * ── Why this exists ─────────────────────────────────────────────────────────────────────────────
 * Client, 2026-09-03, looking at Bulk Upload's amber banner beside Folder Management's red
 * "Retire a segment" one: *"change the background color to follow the Retire a Segment, any banner
 * that uses yellow pls follow the Retire a Segment color."*
 *
 * Before this, fifteen banner boxes across thirteen files each carried their own hexes, in FOUR
 * different ambers (`#fff4e5`, `#fff4ce`, `#fff8e6`, `#fff8e1`) with four different text colours —
 * because each was written where it was needed rather than taken from anywhere. So a request to
 * change "the banner colour" was a thirteen-file sweep with no way to tell whether it had been
 * finished. It is one line now, which matters: this is the SECOND colour change the client has asked
 * for (the replace-clash popup's icon went amber to red the same week).
 *
 * ── What this is NOT for ────────────────────────────────────────────────────────────────────────
 * ⚠ BADGES, CHIPS AND STATUS PILLS KEEP THEIR OWN COLOURS, and must not be pointed here. On
 * My Submissions an amber `Pending` sits beside a red `Rejected`, a slate `Archived` and a pale-red
 * `Failed`, and those four colours ARE the information — `Rejected` means an approver said no and
 * `Failed` means the action broke, which is not the same fact and must not look like it. The client
 * asked about banners, and a banner is a full-width box explaining a state, not a word next to a
 * filename.
 *
 * ⚠ TOASTS ARE ALSO EXCLUDED (`shared/toast.tsx`). A toast is transient and already has its own
 * `error` variant; making every warning toast look like an error would leave nothing to distinguish
 * "this did not work" from "this worked, with a caveat".
 *
 * ── The cost, stated ────────────────────────────────────────────────────────────────────────────
 * ⚠ AMBER USED TO MEAN "DOUBT" AND RED "FAILURE", AND THAT DISTINCTION IS NOW GONE FROM BANNERS.
 * Several of these say *could not be checked* / *could not be read*, where this codebase deliberately
 * treats an unanswerable read as different from a refused one — `unknown` is the amber one on
 * purpose. With one palette, a banner reporting doubt looks exactly like one reporting a failure.
 * The WORDING still distinguishes them, and every one of those banners already names which it is;
 * only the colour stopped carrying it. Accepted on the client's instruction — worth re-raising if
 * they ever ask why a "not checked" notice looks alarming.
 */

/** Taken verbatim from `FolderAdmin.tsx`'s `danger` style, which is the banner the client pointed at. */
export const NOTICE_ATTENTION_BG = "#fdf3f4";
export const NOTICE_ATTENTION_BORDER = "#f1b0b3";
export const NOTICE_ATTENTION_TEXT = "#a4262c";

/**
 * Spread into a style object: `warn: { ...NOTICE_ATTENTION, padding: "10px 12px" }`.
 *
 * Deliberately a plain object rather than a `React.CSSProperties`, so this module stays free of any
 * framework import and can be used by the two web parts whose styles are a CSS template literal as
 * readily as by the ones using inline style objects.
 */
export const NOTICE_ATTENTION = {
  background: NOTICE_ATTENTION_BG,
  border: "1px solid " + NOTICE_ATTENTION_BORDER,
  color: NOTICE_ATTENTION_TEXT,
};

/**
 * The same three, as CSS declarations, for a `<style>` template literal.
 *
 * ⚠ NO BACKTICKS ANYWHERE NEAR THIS when it is interpolated into one — a backtick inside such a
 * literal ENDS it, and the resulting error names neither the cause nor the line. That has broken
 * `Form.tsx` four times, so this string is built with concatenation rather than interpolation.
 */
export const NOTICE_ATTENTION_CSS =
  "background: " + NOTICE_ATTENTION_BG + "; border: 1px solid " + NOTICE_ATTENTION_BORDER + "; " +
  "color: " + NOTICE_ATTENTION_TEXT + ";";

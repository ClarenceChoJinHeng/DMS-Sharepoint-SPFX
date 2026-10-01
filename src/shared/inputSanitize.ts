/**
 * Characters refused in the fields that name and describe a document.
 *
 * Client QA, 2026-08-30: *"PLEASE REMOVE / DISABLE USER TO KEY IN SPECIAL CHARACTER FOR ANY FIELDS,
 * EXCEPT ADDING NEW USERS (EMAIL ADDRESS) INTO THE GROUP"*, then, asked to define it: *"Only block
 * the neccesary ones, I suppose lets follow your recommendation, also remove !, [] @ #"*.
 *
 * ⚠ NOT `sanitizeFolderSegment`, which already exists and is a different job. That one names a
 * FOLDER from a term label — it strips SharePoint's illegal set, collapses whitespace and trims, all
 * of which are right for a derived name and wrong for a box somebody is still typing in. This runs
 * on a keystroke, blocks a wider set, and must leave spacing alone.
 *
 * ⚠ TWO GROUPS OF CHARACTERS, BLOCKED FOR DIFFERENT REASONS.
 *
 * `" * : < > ? / \ |` are the characters **SharePoint itself rejects in a file name**, so these are
 * not a matter of taste: `composeUploadBase` builds the uploaded file's name out of Project Name,
 * Vendor/Customer, Document Name and the date, so one of them typed into any of those fields
 * produces an upload SharePoint refuses — and the refusal names the FILE, not the field, leaving the
 * uploader to guess which box was at fault.
 *
 * `! [ ] @ #` are the client's own additions. SharePoint tolerates them; the client does not want
 * them in document names. `#` is worth knowing about on its own — it starts the fragment in a URL,
 * so a file named with one is awkward to link to even where it is legal.
 *
 * ⚠ WHAT IS DELIBERATELY NOT BLOCKED, and why the list stops here:
 * - **`&`, and especially the fullwidth `＆`.** GHO's own department is *Group Legal, Risk ＆
 *   Compliance* and the term store requires the fullwidth form. Blocking it would refuse the
 *   client's own data.
 * - **`'` and `-`**, which appear in real values (`O'Brien`, `Q1-2026`).
 * - **`.` in the middle of a value.** A leading dot is removed separately below, because SharePoint
 *   refuses a name that starts with one.
 */
export const BLOCKED_INPUT_CHARS = '"*:<>?/\\|![]@#';

/** The blocked characters actually present, de-duplicated, in the order they appear. */
export function blockedCharsIn(value: string): string[] {
  const out: string[] = [];
  for (const ch of (value ?? "").split("")) {
    if (BLOCKED_INPUT_CHARS.indexOf(ch) !== -1 && out.indexOf(ch) === -1) out.push(ch);
  }
  return out;
}

/**
 * The value with those characters removed.
 *
 * ⚠ ALSO REMOVES A LEADING DOT, which is not in the blocked list because a dot is legal everywhere
 * else in a value. SharePoint refuses a file or folder name that STARTS with one, and that failure
 * arrives as a rejected upload rather than as anything about the field.
 *
 * Whitespace is left exactly as typed — including trailing, because somebody is probably still
 * typing. Trimming happens where the value is USED, not while it is being entered.
 */
export function stripBlockedChars(value: string): string {
  let out = "";
  for (const ch of (value ?? "").split("")) {
    if (BLOCKED_INPUT_CHARS.indexOf(ch) === -1) out += ch;
  }
  return out.replace(/^\.+/, "");
}

/**
 * What to tell somebody whose keystroke was refused, or `undefined` if nothing was.
 *
 * ⚠ IT NAMES THE CHARACTERS. "Invalid characters" sends the uploader hunting through their own
 * text; the whole value of catching this at the keystroke is that they can see what was rejected
 * while they still remember typing it.
 */
export function blockedCharsMessage(value: string): string | undefined {
  const found = blockedCharsIn(value);
  if (found.length === 0) return undefined;
  return `${found.join(" ")} cannot be used here.`;
}

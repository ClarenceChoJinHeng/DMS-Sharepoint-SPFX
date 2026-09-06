// A single-choice dropdown with a filter box, for option lists too long to scan.
//
// Client, 2026-09-04: *"Also the dropdown can you use a customize dropdown? Is so long."* — the Audit
// Log's Action filter offers every event type the product can write (26 and growing), which as a
// native `<select>` is a menu nobody reads to the end of.
//
// ⚠ REPLACING A NATIVE `<select>` MEANS TAKING ON WHAT IT GAVE FOR FREE. Keyboard support is not
// optional politeness here: the native control could be opened, arrowed through, type-ahead searched
// and closed with Escape, and a custom one answering only to a mouse is a real regression for anyone
// who does not use one. Arrow keys, Enter, Escape and click-outside are all implemented below. What
// is NOT is native type-ahead — the filter box replaces it and is better at it.
//
// ⚠ THE POPUP IS ABSOLUTELY POSITIONED, so a scroll container ANYWHERE above it clips the list. That
// has already broken three screens in this project — Group Management's people picker, the upload
// form's info panels, the member-add dropdown. `AuditLog.tsx` has no `overflow` rule at all (checked
// before this was written); any future host has to be checked the same way.
import * as React from "react";
import { useEffect, useRef, useState } from "react";

export interface FilterSelectOption {
  value: string;
  label: string;
}

const s: Record<string, React.CSSProperties> = {
  /* ⚠ EVERY KEY USED BELOW MUST EXIST IN THIS OBJECT. It is a `Record<string, CSSProperties>`, so a
     key that does not yields `undefined` and the element renders UNSTYLED with a green build — which
     is exactly how the Requests accordion header shipped wearing the browser's default button
     chrome, and how four other screens here have been caught. */
  wrap: { position: "relative" },
  button: {
    width: "100%", boxSizing: "border-box", display: "flex", alignItems: "center", gap: 8,
    padding: "7px 10px", fontSize: 13, fontFamily: "inherit", textAlign: "left",
    border: "1px solid #c7c7c7", borderRadius: 8, background: "#fff", cursor: "pointer",
    color: "#1b1b1b",
  },
  buttonOff: {
    width: "100%", boxSizing: "border-box", display: "flex", alignItems: "center", gap: 8,
    padding: "7px 10px", fontSize: 13, fontFamily: "inherit", textAlign: "left",
    border: "1px solid #e6e6e6", borderRadius: 8, background: "#f6f6f6", cursor: "not-allowed",
    color: "#a6a6a6",
  },
  /* The chosen label takes the free space so the chevron sits at the right edge, and truncates rather
     than wrapping — a two-line button would shift every control beside it. */
  value: { flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  placeholder: {
    flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
    color: "#8a8886",
  },
  chevron: { flexShrink: 0, fontSize: 10, color: "#605e5c" },
  panel: {
    position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 30,
    background: "#fff", border: "1px solid #c7c7c7", borderRadius: 8,
    boxShadow: "0 4px 14px rgba(0,0,0,.14)", overflow: "hidden",
  },
  search: {
    width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: 13,
    fontFamily: "inherit", border: "none", borderBottom: "1px solid #edebe9", outline: "none",
  },
  /* Capped so the panel never runs off the screen. This scroll box is INSIDE the popup, so it clips
     nothing but its own rows. */
  list: { maxHeight: 260, overflowY: "auto" },
  option: {
    padding: "7px 10px", fontSize: 13, cursor: "pointer", color: "#1b1b1b",
    borderBottom: "1px solid #f5f4f4",
  },
  optionOn: {
    padding: "7px 10px", fontSize: 13, cursor: "pointer", color: "#0f6c3f",
    borderBottom: "1px solid #f5f4f4", background: "#f2f8f4", fontWeight: 600,
  },
  none: { padding: "10px", fontSize: 12, color: "#605e5c" },
};

/**
 * @param value    the chosen option's `value`; `""` means none.
 * @param onChange called with the new value, `""` when the "any" row is chosen.
 * @param options  the list, in the order it should appear. NOT sorted here — the caller's order is
 *                 usually meaningful (the Audit Log's groups events by what writes them).
 * @param anyLabel the top row that clears the choice. Omit to make a choice mandatory.
 */
export function FilterSelect({
  value,
  onChange,
  options,
  anyLabel,
  placeholder = "Select…",
  searchPlaceholder = "Type to filter…",
  disabled,
  id,
}: {
  value: string;
  onChange: (next: string) => void;
  options: readonly FilterSelectOption[];
  anyLabel?: string;
  placeholder?: string;
  searchPlaceholder?: string;
  disabled?: boolean;
  id?: string;
}): React.ReactElement {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  /** Which row the keyboard is on. An index into `shown` below, NOT into `options`. */
  const [active, setActive] = useState(0);
  const wrapRef = useRef<HTMLDivElement | undefined>(undefined);
  const searchRef = useRef<HTMLInputElement | undefined>(undefined);

  /* The "any" row is prepended as an ORDINARY option so one list serves both mouse and keyboard. A
     separately rendered row would need its own click handler, its own highlight and its own place in
     the arrow-key order — three chances to disagree with the rest of the list. */
  const all: FilterSelectOption[] =
    anyLabel === undefined
      ? options.slice()
      : [{ value: "", label: anyLabel }].concat(options.slice());

  const q = query.trim().toLowerCase();
  const shown = q.length === 0 ? all : all.filter((o) => o.label.toLowerCase().indexOf(q) !== -1);

  /* ⚠ FALLS BACK TO THE RAW VALUE, NEVER TO BLANK. An option that has since left the list is still a
     filter that is still APPLIED, and printing nothing would tell the reader the filter is off while
     the rows on screen are narrowed by it. */
  const chosen = all.filter((o) => o.value === value)[0];
  const chosenLabel = chosen ? chosen.label : value;

  /* ⚠ CLOSE ON `mousedown`, NOT `click`. A press on another control fires mousedown → that control's
     handler → click; closing on `click` lets this panel swallow the first press, so the next control
     needs two. */
  useEffect(() => {
    if (!open) return undefined;
    const onDown = (e: MouseEvent): void => {
      const el = wrapRef.current;
      if (el && e.target instanceof Node && !el.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, [open]);

  // Opening puts the caret in the filter box — typing to narrow is the whole point of this control.
  useEffect(() => {
    if (open && searchRef.current) searchRef.current.focus();
  }, [open]);

  const choose = (v: string): void => {
    onChange(v);
    setOpen(false);
    setQuery("");
    setActive(0);
  };

  const onKey = (e: React.KeyboardEvent): void => {
    if (e.key === "Escape") {
      setOpen(false);
      setQuery("");
      return;
    }
    if (!open && (e.key === "ArrowDown" || e.key === "Enter" || e.key === " ")) {
      e.preventDefault();
      setOpen(true);
      return;
    }
    if (!open) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActive((n) => Math.min(n + 1, shown.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActive((n) => Math.max(n - 1, 0));
    } else if (e.key === "Enter") {
      e.preventDefault();
      // Guarded: filtering down to nothing leaves `active` pointing at a row that is not there.
      if (shown[active]) choose(shown[active].value);
    }
  };

  return (
    <div style={s.wrap} ref={(el) => { wrapRef.current = el ?? undefined; }} onKeyDown={onKey}>
      <button
        type="button"
        id={id}
        style={disabled ? s.buttonOff : s.button}
        disabled={disabled}
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        <span style={chosenLabel ? s.value : s.placeholder}>{chosenLabel || placeholder}</span>
        <span style={s.chevron} aria-hidden="true">{open ? "▲" : "▼"}</span>
      </button>

      {open && (
        <div style={s.panel}>
          <input
            ref={(el) => { searchRef.current = el ?? undefined; }}
            style={s.search}
            value={query}
            placeholder={searchPlaceholder}
            aria-label={searchPlaceholder}
            onChange={(e) => {
              setQuery(e.target.value);
              // The highlight must follow the list, or Enter picks a row that has scrolled out of it.
              setActive(0);
            }}
          />
          <div style={s.list} role="listbox">
            {shown.length === 0 ? (
              <div style={s.none}>Nothing matches &ldquo;{query.trim()}&rdquo;.</div>
            ) : (
              shown.map((o, i) => (
                <div
                  key={o.value || "__any__"}
                  role="option"
                  aria-selected={o.value === value}
                  style={i === active || o.value === value ? s.optionOn : s.option}
                  onMouseEnter={() => setActive(i)}
                  onClick={() => choose(o.value)}
                >
                  {o.label}
                </div>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
}

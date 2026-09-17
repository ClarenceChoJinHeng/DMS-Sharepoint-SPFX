// A single-date picker with a calendar popup we can style, replacing the native `<input type="date">`.
//
// Client: *"use custom date dropdown that we can stlyle for the date popup"* — the native calendar
// is drawn by the OS/browser (Chrome, Edge, Firefox all render it differently) and cannot be
// restyled to match the rest of the page, which is exactly the complaint `FilterSelect` was already
// built to answer for the Action dropdown beside it.
//
// ⚠ THE VALUE CONTRACT IS UNCHANGED FROM THE NATIVE CONTROL IT REPLACES: a plain `YYYY-MM-DD` string,
// `""` for none. Every existing caller (`AuditLog`'s `dateFrom`/`dateTo`, fed straight into
// `currentQuery`'s `new Date(v)` parse) needs no change beyond swapping the element — this is a
// control swap, not a data-shape change.
//
// ⚠ THE POPUP IS ABSOLUTELY POSITIONED, exactly like `FilterSelect` beside it — a scroll container
// ANYWHERE above it would clip the calendar. That has already broken three screens in this project
// (Group Management's people picker, the upload form's info panels, the member-add dropdown).
// `AuditLog.tsx` has no `overflow` rule around this control; any future host must be checked the
// same way before mounting this.
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import {
  calendarGrid,
  displayIso,
  isoToday,
  monthLabel,
  parseIso,
  shiftMonth,
  WEEKDAY_HEADERS,
} from "./dateCalendar";

const s: Record<string, React.CSSProperties> = {
  /* ⚠ EVERY KEY USED BELOW MUST EXIST IN THIS OBJECT — see the identical warning on `FilterSelect`,
     which this file otherwise mirrors: a key that does not exist yields `undefined` and the element
     renders unstyled with a green build. */
  wrap: { position: "relative", display: "inline-block" },
  button: {
    width: "100%", boxSizing: "border-box", display: "flex", alignItems: "center", gap: 6,
    padding: "7px 9px", fontSize: 13, fontFamily: "inherit", textAlign: "left",
    border: "1px solid #c8c8c8", borderRadius: 8, background: "#fff", cursor: "pointer",
    color: "#1b1b1b",
  },
  buttonOff: {
    width: "100%", boxSizing: "border-box", display: "flex", alignItems: "center", gap: 6,
    padding: "7px 9px", fontSize: 13, fontFamily: "inherit", textAlign: "left",
    border: "1px solid #e6e6e6", borderRadius: 8, background: "#f6f6f6", cursor: "not-allowed",
    color: "#a6a6a6",
  },
  value: { flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  placeholder: {
    flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
    color: "#8a8886",
  },
  icon: { flexShrink: 0, fontSize: 12, lineHeight: 1, color: "#605e5c" },
  panel: {
    position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 30, width: 232,
    boxSizing: "border-box",
    background: "#fff", border: "1px solid #c7c7c7", borderRadius: 8,
    boxShadow: "0 4px 14px rgba(0,0,0,.14)", padding: 10,
  },
  nav: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 },
  navBtn: {
    border: "none", background: "transparent", cursor: "pointer", fontSize: 14,
    color: "#0f6c3f", fontWeight: 700, padding: "2px 8px", borderRadius: 4, lineHeight: 1,
  },
  monthLabel: { fontSize: 13, fontWeight: 700, color: "#1b1b1b" },
  weekRow: { display: "grid", gridTemplateColumns: "repeat(7, 1fr)", marginBottom: 2 },
  weekday: {
    fontSize: 10.5, fontWeight: 600, color: "#8a8886", textAlign: "center", padding: "2px 0",
  },
  grid: { display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 2 },
  day: {
    fontSize: 12.5, textAlign: "center", padding: "6px 0", borderRadius: 6, cursor: "pointer",
    border: "1px solid transparent", background: "transparent", color: "#1b1b1b",
    fontFamily: "inherit",
  },
  dayOut: {
    fontSize: 12.5, textAlign: "center", padding: "6px 0", borderRadius: 6, cursor: "pointer",
    border: "1px solid transparent", background: "transparent", color: "#c7c7c7",
    fontFamily: "inherit",
  },
  dayToday: {
    fontSize: 12.5, textAlign: "center", padding: "6px 0", borderRadius: 6, cursor: "pointer",
    border: "1px solid #0f6c3f", background: "transparent", color: "#0f6c3f", fontWeight: 700,
    fontFamily: "inherit",
  },
  daySelected: {
    fontSize: 12.5, textAlign: "center", padding: "6px 0", borderRadius: 6, cursor: "pointer",
    border: "1px solid #0f6c3f", background: "#0f6c3f", color: "#fff", fontWeight: 700,
    fontFamily: "inherit",
  },
  footer: {
    display: "flex", justifyContent: "space-between", marginTop: 8, paddingTop: 8,
    borderTop: "1px solid #edebe9",
  },
  footBtn: {
    border: "none", background: "transparent", cursor: "pointer", fontSize: 12,
    color: "#0f6c3f", fontWeight: 600, padding: "4px 6px", fontFamily: "inherit",
  },
};

/**
 * @param value       `YYYY-MM-DD`, or `""` for no date chosen — same shape `<input type="date">` used.
 * @param onChange     called with the new `YYYY-MM-DD` value, or `""` when Clear is pressed.
 * @param placeholder  shown in the button when `value` is empty.
 * @param style        merged onto the outer wrapper — the caller's width, not this file's concern.
 */
export function DatePicker({
  value,
  onChange,
  placeholder = "Select date",
  disabled,
  id,
  style,
}: {
  value: string;
  onChange: (next: string) => void;
  placeholder?: string;
  disabled?: boolean;
  id?: string;
  style?: React.CSSProperties;
}): React.ReactElement {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement | undefined>(undefined);

  // The month currently ON SCREEN. Seeded once from whatever `value` holds at mount (or today, if
  // blank) — re-derived every time the popup OPENS below, so it does not go stale while it is closed.
  const seed = parseIso(value) ?? parseIso(isoToday());
  const [viewYear, setViewYear] = useState(seed ? seed.year : new Date().getFullYear());
  const [viewMonth, setViewMonth] = useState(seed ? seed.month : new Date().getMonth() + 1);

  // Re-centre on the picked value every time the popup opens, so a filter set to last month does not
  // silently reopen on today's page.
  useEffect(() => {
    if (!open) return;
    const p = parseIso(value) ?? parseIso(isoToday());
    if (!p) return;
    setViewYear(p.year);
    setViewMonth(p.month);
    // `[open]` only, deliberately — only the OPEN edge should re-centre. Depending on `value` too
    // would snap the calendar back to the selected month on every click while it is already open,
    // fighting the Prev/Next month buttons.
  }, [open]);

  /* Close on `mousedown`, not `click` — same reasoning as `FilterSelect`: closing on `click` lets
     this panel swallow the first press on the next control, so it needs two clicks instead of one. */
  useEffect(() => {
    if (!open) return undefined;
    const onDown = (e: MouseEvent): void => {
      const el = wrapRef.current;
      if (el && e.target instanceof Node && !el.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, [open]);

  const choose = (iso: string): void => {
    onChange(iso);
    setOpen(false);
  };

  const goMonth = (delta: number): void => {
    const n = shiftMonth(viewYear, viewMonth, delta);
    setViewYear(n.year);
    setViewMonth(n.month);
  };

  const onKey = (e: React.KeyboardEvent): void => {
    if (e.key === "Escape") {
      setOpen(false);
      return;
    }
    if (!open && (e.key === "Enter" || e.key === " ")) {
      e.preventDefault();
      setOpen(true);
    }
  };

  const cells = calendarGrid(viewYear, viewMonth);
  const todayIso = isoToday();

  return (
    <div
      style={style ? { ...s.wrap, ...style } : s.wrap}
      ref={(el) => {
        wrapRef.current = el ?? undefined;
      }}
      onKeyDown={onKey}
    >
      <button
        type="button"
        id={id}
        style={disabled ? s.buttonOff : s.button}
        disabled={disabled}
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        <span style={value ? s.value : s.placeholder}>
          {value ? displayIso(value) : placeholder}
        </span>
        <span style={s.icon} aria-hidden="true">
          📅
        </span>
      </button>

      {open && (
        <div style={s.panel} role="dialog" aria-label="Choose a date">
          <div style={s.nav}>
            <button
              type="button"
              style={s.navBtn}
              aria-label="Previous month"
              onClick={() => goMonth(-1)}
            >
              &lsaquo;
            </button>
            <span style={s.monthLabel}>{monthLabel(viewYear, viewMonth)}</span>
            <button
              type="button"
              style={s.navBtn}
              aria-label="Next month"
              onClick={() => goMonth(1)}
            >
              &rsaquo;
            </button>
          </div>

          <div style={s.weekRow}>
            {WEEKDAY_HEADERS.map((w) => (
              <span key={w} style={s.weekday}>
                {w}
              </span>
            ))}
          </div>

          <div style={s.grid} role="grid">
            {cells.map((c) => {
              const cellStyle =
                c.iso === value
                  ? s.daySelected
                  : c.iso === todayIso
                    ? s.dayToday
                    : c.inMonth
                      ? s.day
                      : s.dayOut;
              return (
                <button
                  key={c.iso}
                  type="button"
                  role="gridcell"
                  aria-selected={c.iso === value}
                  style={cellStyle}
                  onClick={() => choose(c.iso)}
                >
                  {c.day}
                </button>
              );
            })}
          </div>

          <div style={s.footer}>
            <button type="button" style={s.footBtn} onClick={() => choose(todayIso)}>
              Today
            </button>
            {value && (
              <button type="button" style={s.footBtn} onClick={() => choose("")}>
                Clear
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

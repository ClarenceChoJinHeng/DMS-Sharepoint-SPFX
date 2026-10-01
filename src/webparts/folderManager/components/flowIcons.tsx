// The Folder Management flow-card icons.
//
// Drawn from the client's artwork supplied 2026-08-30 as PNGs — `New-Folder.png`,
// `Group 633982.png`, `Edit-Folder.png`, `Rename-Folder.png`.
//
// ⚠ SVG, NOT THE PNGs THEMSELVES, for two reasons.
//   1. **Provisioning image assets does not work on this tenant.** `elements.xml` has failed to
//      provision twice, both times silently (the `+ New Folder` customizer, then the bulk-approve
//      command set). A packaged image would render as a broken box with nothing explaining it —
//      which is exactly why the bulk-approve extension ships its icon as a base64 data URI.
//   2. The CRS Settings cards next door are already SVG (`crsSettings/components/icons.tsx`), drawn
//      to one recipe. A raster tile beside them reads as a different family, and these glyphs are
//      simple enough that nothing is lost by redrawing them.
//
// Same recipe as that file, so the two pages look like one system: a 54×54 canvas, an `rx=10`
// pastel tile, one flat glyph in a darker shade of the same hue. Here the hue is the client's green
// throughout, because these are one group of related actions rather than five unrelated
// destinations.
//
// ⚠ Base64 data URIs remain the fallback if pixel-exactness is ever wanted — roughly 7 KB for all
// four, which is affordable. Not chosen because raster art does not scale on a high-DPI screen, and
// these tiles are the largest thing on the card.
import * as React from "react";

export type FlowIconName = "newSegment" | "addUnit" | "structure" | "rename" | "reconcile";

const GREEN = "#00684A";

/** The pale tile every glyph sits on. One definition, so the five cannot drift apart. */
function Tile(): React.ReactElement {
  return <rect width="54" height="54" rx="10" fill="#D5EBD2" />;
}

/** Add a new segment — a folder with a plus. */
function NewSegment(): React.ReactElement {
  return (
    <svg width="54" height="54" viewBox="0 0 54 54" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <Tile />
      <path
        d="M14 20.5c0-1.1.9-2 2-2h6.2c.6 0 1.2.3 1.6.8l1.8 2.4c.2.2.4.3.7.3H32c1.1 0 2 .9 2 2v3.2M14 20.5v13c0 1.1.9 2 2 2h11"
        stroke={GREEN}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <circle cx="35" cy="34" r="6.5" stroke={GREEN} strokeWidth="2" />
      <path d="M35 31v6M32 34h6" stroke={GREEN} strokeWidth="2" strokeLinecap="round" />
    </svg>
  );
}

/** Add a department or unit — the hierarchy, one box above two. */
function AddUnit(): React.ReactElement {
  return (
    <svg width="54" height="54" viewBox="0 0 54 54" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <Tile />
      <rect x="22" y="13" width="10" height="8" rx="2" stroke={GREEN} strokeWidth="2" />
      <rect x="13" y="33" width="10" height="8" rx="2" stroke={GREEN} strokeWidth="2" />
      <rect x="31" y="33" width="10" height="8" rx="2" stroke={GREEN} strokeWidth="2" />
      {/* The connectors: down from the parent, across, then down into each child. */}
      <path
        d="M27 21v6M18 33v-6h18v6"
        stroke={GREEN}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

/** Change the folder structure — a folder with a pencil. */
function Structure(): React.ReactElement {
  return (
    <svg width="54" height="54" viewBox="0 0 54 54" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <Tile />
      <path
        d="M14 20.5c0-1.1.9-2 2-2h6.2c.6 0 1.2.3 1.6.8l1.8 2.4c.2.2.4.3.7.3H32c1.1 0 2 .9 2 2v2M14 20.5v13c0 1.1.9 2 2 2h10"
        stroke={GREEN}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M39.6 26.4a2 2 0 0 1 0 2.8l-8.8 8.8-3.7.9.9-3.7 8.8-8.8a2 2 0 0 1 2.8 0z"
        stroke={GREEN}
        strokeWidth="2"
        strokeLinejoin="round"
      />
    </svg>
  );
}

/** Rename or re-code a folder — a folder with "Aa". */
function Rename(): React.ReactElement {
  return (
    <svg width="54" height="54" viewBox="0 0 54 54" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <Tile />
      <path
        d="M14 20.5c0-1.1.9-2 2-2h6.2c.6 0 1.2.3 1.6.8l1.8 2.4c.2.2.4.3.7.3H32c1.1 0 2 .9 2 2v2M14 20.5v13c0 1.1.9 2 2 2h8"
        stroke={GREEN}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      {/* "Aa" as PATHS, not a <text> element: a font is not guaranteed to resolve inside an SVG on
          every machine, and a substituted face would sit differently against the folder. */}
      <path
        d="M27.5 40l4-11 4 11M28.9 36.4h5.2"
        stroke={GREEN}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M42.5 34.2a3 3 0 1 0 0 4.6V40M42.5 36.5V40"
        stroke={GREEN}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

/**
 * Run folder reconciliation — arrows in a circle.
 *
 * ⚠ NOT IN THE CLIENT'S ARTWORK, and its card is not in their mock either. Kept at the client's own
 * instruction (2026-08-30: *"don't drop that card specifically for Folder Reconciliation"*), because
 * it is the only route to a reconciliation run — and since the rail was locked forward the same day,
 * an admin without this card would have to start an unrelated flow and walk every step of it. Drawn
 * to the same recipe so it does not read as a stray.
 */
function Reconcile(): React.ReactElement {
  return (
    <svg width="54" height="54" viewBox="0 0 54 54" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <Tile />
      <path d="M38 27a11 11 0 1 1-3.2-7.8" stroke={GREEN} strokeWidth="2" strokeLinecap="round" />
      <path d="M35.5 12.5v7h-7" stroke={GREEN} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

const ICONS: Record<FlowIconName, () => React.ReactElement> = {
  newSegment: NewSegment,
  addUnit: AddUnit,
  structure: Structure,
  rename: Rename,
  reconcile: Reconcile,
};

/**
 * One flow-card icon.
 *
 * An unrecognised name renders NOTHING rather than throwing. A card whose icon is missing is a
 * cosmetic problem; a card that takes the page down with it is not — and the name arrives from a
 * flow definition, which is data this component does not control.
 */
export function FlowIcon({ name }: { name: string }): React.ReactElement | null {
  const Draw = ICONS[name as FlowIconName];
  return Draw ? <Draw /> : null;
}

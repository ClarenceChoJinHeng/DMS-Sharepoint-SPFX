# Folder Abbreviation Naming Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Name Staging/Documents folders from short per-term abbreviations (`GHO/GCA/GMB_STRATCOMMS`) instead of full term labels, keeping the term store and the upload dropdown on full names, and surfacing the full name in the details pane via a `Full Name` column.

**Architecture:** A new `DMS Term Abbreviation` list maps term GUID → abbreviation. `FolderManager.tsx` is the only component that derives folder names, so the change is contained to its `buildProvisionTargets` plus a new rename step. Uploads are unaffected — `Form.tsx` resolves folders by UniqueId. Safety comes from three layers: collision detection before writing, rename-conflict reporting, and reuse of the existing `incomplete` guard so orphan pruning can never act on a partial term-store read.

**Tech Stack:** SPFx 1.23.0, React 17, TypeScript 5.8, Heft + Jest, SharePoint REST via `SPHttpClient`.

**Spec:** `docs/superpowers/specs/2026-07-30-folder-abbreviation-naming-design.md`
**Data:** `docs/term-store-import/10-per-level-abbreviations.csv` (178 rows, validated)

---

## Phasing

**Phase 1 (this plan, Tasks 1-8):** naming, rename, `Full Name`, orphan prune, messages, guidance. Shippable and testable on its own — the list is seeded by a one-off browser script rather than a UI.

**Phase 2 (separate plan, not written yet):** the admin web part UI — bulk grid for onboarding and a single-term prompt for steady state, per spec §7. Deferred deliberately: it is a substantial piece of UI, and Phase 1 can be exercised end-to-end without it.

---

## Background the engineer needs

**Why abbreviations.** The worst-case encoded path is currently **257 characters** against a ~330 threshold where `GetFolderByServerRelativeUrl` starts returning HTTP 400 (CLAUDE.md gotcha #9) — about 73 characters of headroom. Encoding is the cost: the term store requires **fullwidth ＆**, which is **9 encoded characters** (`%EF%BC%86`), and each space is 3. Abbreviating takes the worst case to **129**.

**The risk this introduces.** Today collisions are impossible because folder names come from term labels and the term store forbids two siblings sharing a name. A hand-maintained abbreviation list has no such constraint. Two units abbreviating identically under the same parent would become **one folder with one ACL holding two units' documents** — breaking the isolation the whole permission model rests on. That is why collision detection is in Task 1, not bolted on later.

**Project constraints:**

1. **No ES2017+ library methods** — no `Array.prototype.includes`, `Promise.allSettled`, `Object.entries`. Use `indexOf`, `Map`, per-item `try/catch`. CLAUDE.md gotcha #3.
2. **Tests:** `npx heft test --test-path-pattern <name>`. There is no `npm test` script. A test importing a not-yet-written module fails as a compile error — that is the expected TDD red.
3. **Always pass paths as an OData parameter alias**, never an inline quoted literal: `GetFolderByServerRelativeUrl(@f)?@f='<encoded>'`. Gotcha #9.
4. **`sanitizeFolderSegment` still applies** to abbreviations. A `/` in a name is both rejected by SharePoint and read as a path separator.

**Commit hygiene, every task:** `git add` only the files named in that task. Never `git add -A`, `git add .`, or `git commit -a` — the working tree carries unrelated staged files. Stay on the current branch.

---

## Task 1: Abbreviation model and collision detection

**Files:**
- Create: `src/shared/folderAbbreviation.ts`
- Test: `src/shared/folderAbbreviation.test.ts`

- [ ] **Step 1: Write the failing test**

Create `src/shared/folderAbbreviation.test.ts`:

```ts
import {
  buildAbbrevIndex,
  lookupAbbrev,
  findCollisions,
  AbbrevRow,
} from "./folderAbbreviation";

const rows: AbbrevRow[] = [
  { termGuid: "AAA", abbreviation: "GCA" },
  { termGuid: "BBB", abbreviation: "GMB_STRATCOMMS" },
];

describe("buildAbbrevIndex / lookupAbbrev", () => {
  it("looks up by term guid, case-insensitively", () => {
    const ix = buildAbbrevIndex(rows);
    expect(lookupAbbrev(ix, "aaa")).toBe("GCA");
    expect(lookupAbbrev(ix, "AAA")).toBe("GCA");
  });

  it("returns undefined for an unmapped term", () => {
    expect(lookupAbbrev(buildAbbrevIndex(rows), "ZZZ")).toBeUndefined();
  });

  it("ignores rows with a blank guid or blank abbreviation", () => {
    const ix = buildAbbrevIndex([
      { termGuid: "", abbreviation: "X" },
      { termGuid: "CCC", abbreviation: "   " },
    ]);
    expect(lookupAbbrev(ix, "CCC")).toBeUndefined();
  });

  it("trims surrounding whitespace on the abbreviation", () => {
    const ix = buildAbbrevIndex([{ termGuid: "DDD", abbreviation: "  GF  " }]);
    expect(lookupAbbrev(ix, "DDD")).toBe("GF");
  });
});

describe("findCollisions", () => {
  // The failure this whole guard exists to prevent: two units under one parent
  // resolving to the same folder means one ACL over two units' documents.
  it("reports two siblings sharing an abbreviation", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "STRATCOMMS", label: "GMB - Strategic Communications" },
        { parentPath: "/GHO/GCA", termGuid: "B", abbreviation: "STRATCOMMS", label: "GC - Strategic Communications" },
      ]),
    ).toEqual([
      {
        parentPath: "/GHO/GCA",
        abbreviation: "STRATCOMMS",
        labels: ["GMB - Strategic Communications", "GC - Strategic Communications"],
      },
    ]);
  });

  it("allows the same abbreviation under different parents", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "PM", label: "x" },
        { parentPath: "/MHO/GCA", termGuid: "B", abbreviation: "PM", label: "y" },
      ]),
    ).toEqual([]);
  });

  it("compares case-insensitively — SharePoint folder names are not case-unique", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "cors", label: "x" },
        { parentPath: "/GHO/GCA", termGuid: "B", abbreviation: "CORS", label: "y" },
      ]).length,
    ).toBe(1);
  });

  it("returns an empty array when everything is unique", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "IR", label: "x" },
        { parentPath: "/GHO/GCA", termGuid: "B", abbreviation: "YG", label: "y" },
      ]),
    ).toEqual([]);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

Run: `npx heft test --test-path-pattern folderAbbreviation`
Expected: FAIL — cannot resolve `./folderAbbreviation`.

- [ ] **Step 3: Implement**

Create `src/shared/folderAbbreviation.ts`:

```ts
/**
 * Per-term folder abbreviations, from the `DMS Term Abbreviation` list.
 *
 * Folder names used to come from term labels, which the term store guarantees are
 * unique among siblings. Abbreviations are hand-maintained and carry no such
 * guarantee, so `findCollisions` restores it: two siblings resolving to one folder
 * would mean one ACL over two units' documents.
 *
 * See docs/superpowers/specs/2026-07-30-folder-abbreviation-naming-design.md section 4.
 */
export interface AbbrevRow {
  termGuid: string;
  abbreviation: string;
}

/** A folder target awaiting its name, with enough context to detect a clash. */
export interface AbbrevTarget {
  parentPath: string;
  termGuid: string;
  abbreviation: string;
  label: string;
}

export interface AbbrevCollision {
  parentPath: string;
  abbreviation: string;
  labels: string[];
}

/** Term GUIDs are compared lowercased, matching DMS Group Map and DMS Folder Map. */
export function buildAbbrevIndex(rows: readonly AbbrevRow[]): Map<string, string> {
  const ix = new Map<string, string>();
  rows.forEach((r) => {
    const guid = (r.termGuid ?? "").trim().toLowerCase();
    const abbrev = (r.abbreviation ?? "").trim();
    if (guid.length === 0 || abbrev.length === 0) return;
    ix.set(guid, abbrev);
  });
  return ix;
}

export function lookupAbbrev(
  index: Map<string, string>,
  termGuid: string,
): string | undefined {
  return index.get((termGuid ?? "").trim().toLowerCase());
}

/**
 * Siblings sharing an abbreviation, compared case-insensitively because SharePoint
 * folder names are not unique by case — `CORS` and `cors` collide in one parent.
 */
export function findCollisions(
  targets: readonly AbbrevTarget[],
): AbbrevCollision[] {
  const groups = new Map<string, AbbrevTarget[]>();
  targets.forEach((t) => {
    const key = `${t.parentPath} ${t.abbreviation.toLowerCase()}`;
    const bucket = groups.get(key);
    if (bucket) bucket.push(t);
    else groups.set(key, [t]);
  });
  const out: AbbrevCollision[] = [];
  groups.forEach((bucket) => {
    if (bucket.length < 2) return;
    out.push({
      parentPath: bucket[0].parentPath,
      abbreviation: bucket[0].abbreviation,
      labels: bucket.map((b) => b.label),
    });
  });
  return out;
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `npx heft test --test-path-pattern folderAbbreviation`
Expected: PASS, 9 tests.

- [ ] **Step 5: Commit**

```bash
git add src/shared/folderAbbreviation.ts src/shared/folderAbbreviation.test.ts
git commit -m "feat: add term abbreviation lookup with sibling collision detection"
```

---

## Task 2: Create and seed the `DMS Term Abbreviation` list

No web part UI in Phase 1 — the list is created by hand and seeded by a browser console script, the same way the term store and config lists have been verified all along.

**Files:**
- Create: `scripts/seed-term-abbreviations.js`

- [ ] **Step 1: Create the list in SharePoint**

On `/sites/ClarenceDMSTesting`, new **Custom List** named `DMS Term Abbreviation` with:

| Column | Type | Notes |
|---|---|---|
| `Title` | Text | the term's full label, for humans reading the list |
| `TermGuid` | Text | the key |
| `Abbreviation` | Text | the folder name segment |
| `Level` | Choice | `Segment` \| `Department` \| `Unit` |

- [ ] **Step 2: Write the seeding script**

Create `scripts/seed-term-abbreviations.js`. It is pasted into DevTools on a page of the site — it is not bundled or imported by the solution.

```js
/**
 * Seeds DMS Term Abbreviation from docs/term-store-import/10-per-level-abbreviations.csv.
 *
 * The CSV is keyed by label path, not term GUID, so this walks the term store and
 * matches labels to resolve GUIDs. Paste the CSV text into CSV_TEXT below, then run
 * the whole file in the DevTools console on any page of the site.
 *
 * Re-runnable: updates the row when the TermGuid already exists.
 */
const SITE = "/sites/ClarenceDMSTesting";
const LIST = "DMS Term Abbreviation";
const SETS = {
  "Group Head Office": "08dd94cb-f76c-431c-9b37-e9c98f739ffc",
  "Minamas Head Office": "9ad00b00-a43c-4a8b-a39a-d0efa89ba706",
  "NBPOL Head Office": "77c3993b-0c3c-4a18-89d9-d69209886322",
};
const CSV_TEXT = `PASTE THE CONTENTS OF 10-per-level-abbreviations.csv HERE`;

const j = async (url) => {
  const r = await fetch(url, { headers: { Accept: "application/json;odata=nometadata" } });
  if (!r.ok) throw new Error(`${r.status} on ${url}: ${await r.text()}`);
  return r.json();
};

// Minimal CSV reader: handles quoted fields containing commas, which the
// department "Group Legal, Risk & Compliance" requires.
function parseCsv(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (c === '"') inQuotes = false;
      else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ",") { row.push(field); field = ""; }
    else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
    else if (c !== "\r") field += c;
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  const head = rows.shift();
  return rows.filter((r) => r.length === head.length)
             .map((r) => { const o = {}; head.forEach((h, i) => (o[h.trim()] = r[i])); return o; });
}

(async () => {
  // 1. Build label -> guid for every term in every set.
  const guidByPath = new Map();
  for (const [segName, setId] of Object.entries(SETS)) {
    const walk = async (parentUrl, prefix) => {
      const d = await j(`${SITE}/_api/v2.1/termStore/sets/${setId}/${parentUrl}`);
      for (const t of d.value ?? []) {
        const label = t.labels[0].name;
        const path = prefix ? `${prefix}|${label}` : `${segName}|${label}`;
        guidByPath.set(path, t.id);
        await walk(`terms/${t.id}/children`, path);
      }
    };
    await walk("children", "");
  }
  console.log(`resolved ${guidByPath.size} terms`);

  // 2. Read existing rows so the script is re-runnable.
  const existing = new Map();
  const cur = await j(`${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items?$select=Id,TermGuid&$top=5000`);
  (cur.value ?? []).forEach((r) => existing.set((r.TermGuid || "").toLowerCase(), r.Id));

  // 3. Write.
  const digest = (await (await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST", headers: { Accept: "application/json;odata=nometadata" },
  })).json()).FormDigestValue;

  let written = 0; const missing = [];
  for (const r of parseCsv(CSV_TEXT)) {
    if (r.Level === "Segment") continue; // segments use StagingFolder in DMS Config
    const path = r.Level === "Department"
      ? `${r.BusinessSegment}|${r.Department}`
      : `${r.BusinessSegment}|${r.Department}|${r.Unit}`;
    const guid = guidByPath.get(path);
    if (!guid) { missing.push(path); continue; }
    const id = existing.get(guid.toLowerCase());
    const url = id
      ? `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items(${id})`
      : `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items`;
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        ...(id ? { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" } : {}),
      },
      body: JSON.stringify({
        Title: r.FullName, TermGuid: guid, Abbreviation: r.Abbreviation, Level: r.Level,
      }),
    });
    if (!res.ok) { console.error(`FAILED ${path}: ${res.status} ${await res.text()}`); continue; }
    written++;
  }
  console.log(`wrote ${written} rows`);
  if (missing.length) console.warn(`could not resolve ${missing.length} label paths:`, missing);
})();
```

- [ ] **Step 3: Run it and check the output**

Expected: `resolved 175 terms` (42 departments + 133 units), `wrote 175 rows`, and **no unresolved label paths**. Any entry in the `missing` list means the CSV label does not match the live term label exactly — investigate rather than editing the CSV to fit.

- [ ] **Step 4: Set the segment abbreviations**

In `DMS Config`, set `StagingFolder` on each `mode` row: `mode_gho` → `GHO`, `mode_minamas_ho` → `MHO`, `mode_nbpol_ho` → `NBPOLHO`.

- [ ] **Step 5: Commit**

```bash
git add scripts/seed-term-abbreviations.js
git commit -m "chore: add one-off seeder for the term abbreviation list"
```

---

## Task 3: Name folders from abbreviations

**Files:**
- Modify: `src/webparts/folderManager/components/FolderManager.tsx` — imports, a new loader, and `buildProvisionTargets` (lines 1035-1087)

- [ ] **Step 1: Add the import**

```ts
import {
  AbbrevRow,
  AbbrevTarget,
  buildAbbrevIndex,
  findCollisions,
  lookupAbbrev,
} from "../../../shared/folderAbbreviation";
```

Match the relative depth to the existing `../../../shared/formModel` import in the same file.

- [ ] **Step 2: Load the list**

Add beside `loadReconModes` (around line 942):

```ts
  const ABBREV_LIST = "DMS Term Abbreviation";

  const loadAbbreviations = async (): Promise<Map<string, string>> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(ABBREV_LIST)}')/items?$select=TermGuid,Abbreviation&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`${ABBREV_LIST} read failed: HTTP ${res.status}. ${body}`);
    }
    const data = await res.json();
    const rows: AbbrevRow[] = (data.value ?? []).map(
      (r: { TermGuid?: string; Abbreviation?: string }) => ({
        termGuid: r.TermGuid ?? "",
        abbreviation: r.Abbreviation ?? "",
      }),
    );
    return buildAbbrevIndex(rows);
  };
```

Do **not** wrap this in a `.catch` that returns an empty map. An unreadable list must abort the run — silently proceeding would skip every folder.

- [ ] **Step 3: Use abbreviations for path segments**

In `buildProvisionTargets`, the return type gains two report channels and the naming rule changes.

Change the signature and add the index plus collectors at the top of the function:

```ts
  const buildProvisionTargets = async (): Promise<{
    targets: ProvTarget[];
    incomplete: string[];
    missingAbbrev: Array<{ termGuid: string; label: string }>;
    collisions: ReturnType<typeof findCollisions>;
  }> => {
    const out: ProvTarget[] = [];
    const incomplete: string[] = [];
    const missingAbbrev: Array<{ termGuid: string; label: string }> = [];
    const abbrevTargets: AbbrevTarget[] = [];
    const abbrevIndex = await loadAbbreviations();
    const modes = await loadReconModes();
```

Replace the `seg` helper (line 1054) — it currently sanitises the term label. It must now resolve the abbreviation, and report rather than guess when there is none:

```ts
        // Folder names come from the abbreviation list, NOT the term label — the
        // labels are long and their fullwidth ampersands cost 9 encoded characters
        // each. Spec 2026-07-30. sanitizeFolderSegment still applies: a "/" in a
        // name is both rejected by SharePoint and read as a path separator.
        // A term with no abbreviation is SKIPPED and reported, never guessed at —
        // guessing would produce a folder nobody can find and uploads cannot reach.
        const seg = (termGuid: string, label: string): string | undefined => {
          const abbrev = lookupAbbrev(abbrevIndex, termGuid);
          if (abbrev === undefined) {
            missingAbbrev.push({ termGuid, label });
            return undefined;
          }
          return sanitizeFolderSegment(abbrev) || abbrev;
        };
```

Then at each use site, skip the subtree when the abbreviation is absent. For the top level (lines 1055-1057):

```ts
        const topSeg = seg(top.id, top.label);
        if (topSeg === undefined) continue; // reported; its children are unreachable
        const topTarget: ProvTarget = { termGuid: top.id, assignTerm: top.id, relPath: `/${mode.stagingFolder}/${topSeg}`, label: `${mode.stagingFolder} > ${top.label}`, section: mode.stagingFolder, isLeaf: false };
        out.push(topTarget);
        abbrevTargets.push({ parentPath: `/${mode.stagingFolder}`, termGuid: top.id, abbreviation: topSeg, label: top.label });
```

And inside `walk`, replacing lines 1064-1077:

```ts
          for (const child of children) {
            const childSeg = seg(child.id, child.label);
            if (childSeg === undefined) continue; // reported; skip this subtree
            const chain = [...ancestors, child.label];
            const pathChain = [...pathAncestors, childSeg];
            const parentPath = `/${mode.stagingFolder}/${topSeg}${pathAncestors.length ? "/" + pathAncestors.join("/") : ""}`;
            abbrevTargets.push({ parentPath, termGuid: child.id, abbreviation: childSeg, label: child.label });
            const childTarget: ProvTarget = {
              termGuid: child.id,
              assignTerm: child.id,
              relPath: `/${mode.stagingFolder}/${topSeg}/${pathChain.join("/")}`,
              label: `${mode.stagingFolder} > ${top.label} > ${chain.join(" > ")}`,
              section: mode.stagingFolder,
              isLeaf: false,
            };
            out.push(childTarget);
            childTarget.isLeaf = !(await walk(child.id, chain, pathChain));
          }
```

Finally, replace the return (line 1086):

```ts
    return { targets: out, incomplete, missingAbbrev, collisions: findCollisions(abbrevTargets) };
```

- [ ] **Step 4: Abort the run on a collision**

At the call site of `buildProvisionTargets`, before any folder is created: if `collisions.length > 0`, stop and report each one as
`COLLISION: <parentPath> — "<abbreviation>" used by: <labels joined by " | ">`.

Do **not** create any folder when a collision is present. Two siblings resolving to one path is the ACL-merging failure this guard exists for, and a partial run would create the merged folder before anyone reads the log.

Report `missingAbbrev` as a "needs attention" section — the run continues for every other term, but each skipped term is listed as
`SKIPPED (no abbreviation): <label>`.

- [ ] **Step 5: Verify the build**

Run: `npx heft test`
Expected: PASS, and no new lint warnings beyond the 16 pre-existing ones.

- [ ] **Step 6: Commit**

```bash
git add src/webparts/folderManager/components/FolderManager.tsx
git commit -m "feat: name folders from term abbreviations, skipping unmapped terms"
```

---

## Task 4: Rename a folder whose abbreviation changed

**Files:**
- Modify: `src/shared/dmsFolderMap.ts` — new `renameFolder` beside `ensureFolder`
- Modify: `src/webparts/folderManager/components/FolderManager.tsx` — the mapped-term branch around line 1264

Reconciliation already verifies the stored UniqueId rather than the label path, and deliberately leaves a renamed folder alone. This task turns that tolerance into correction.

- [ ] **Step 1: Add the rename helper**

Add to `src/shared/dmsFolderMap.ts`, beside `ensureFolder`:

```ts
export interface RenameResult {
  ok: boolean;
  conflict: boolean;
  status: number;
  body: string;
}

/**
 * Rename a folder in place, preserving its UniqueId, ACL and contents.
 *
 * A conflict (something already occupies the target name) is reported, never
 * forced: forcing risks merging two folders with different ACLs, and skipping
 * silently leaves a real folder beside an empty decoy with nobody told.
 */
export async function renameFolder(
  client: SPHttpClient,
  siteUrl: string,
  serverRelativeUrl: string,
  newLeafName: string,
): Promise<RenameResult> {
  const parent = serverRelativeUrl.slice(0, serverRelativeUrl.lastIndexOf("/"));
  const target = `${parent}/${newLeafName}`;
  const res = await client.post(
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/MoveTo(newUrl=@n)?@f='${encodeServerRelativePath(serverRelativeUrl)}'&@n='${encodeServerRelativePath(target)}'`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  const body = res.ok ? "" : await res.text();
  // SharePoint answers 400 with an "already exists" message rather than 409.
  const conflict = !res.ok && body.toLowerCase().indexOf("already exists") !== -1;
  return { ok: res.ok, conflict, status: res.status, body };
}
```

Note the OData alias form on both parameters — an inline quoted literal returns HTTP 400 on long paths (gotcha #9).

- [ ] **Step 2: Call it when the name is stale**

In the mapped-term branch, after the existing `probeFolderById` verification confirms the stored UniqueId still resolves, compare the folder's current leaf name to the expected one and rename when they differ:

```ts
                // The abbreviation changed, so the folder's name is now stale.
                // Uploads are unaffected either way — Form.tsx resolves by
                // UniqueId — so this is cosmetic correction, not a repair.
                const currentLeaf = probe.serverRelativeUrl.slice(
                  probe.serverRelativeUrl.lastIndexOf("/") + 1,
                );
                const expectedLeaf = t.relPath.slice(t.relPath.lastIndexOf("/") + 1);
                if (currentLeaf !== expectedLeaf) {
                  const r = await renameFolder(
                    context.spHttpClient, siteUrl, probe.serverRelativeUrl, expectedLeaf,
                  );
                  if (r.ok) {
                    entries.push({ msg: `  ↳ renamed ${currentLeaf} → ${expectedLeaf}`, ok: true });
                  } else if (r.conflict) {
                    entries.push({
                      msg: `  ↳ CANNOT RENAME ${currentLeaf} → ${expectedLeaf}: a folder of that name already exists here. Resolve manually.`,
                      ok: false,
                    });
                  } else {
                    entries.push({
                      msg: `  ↳ rename failed ${currentLeaf} → ${expectedLeaf}: HTTP ${r.status} ${r.body}`,
                      ok: false,
                    });
                  }
                }
```

On success also refresh the row's stored `folderUrl` via `updateFolderMapping` — the UniqueId is unchanged, but the recorded URL is now stale.

Adjust `probe`, `entries`, `existingRow` and the `updateFolderMapping` signature to the names actually in scope at that point — read the surrounding block first. The rule, not the identifiers, is what matters: rename when stale, refresh the stored URL on success, report a conflict distinctly from any other failure, and never retry or force.

- [ ] **Step 3: Verify the build**

Run: `npx heft test`
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add src/shared/dmsFolderMap.ts src/webparts/folderManager/components/FolderManager.tsx
git commit -m "feat: rename folders whose abbreviation changed, reporting conflicts"
```

---

## Task 5: Write `Full Name` on every folder

**Files:**
- Modify: `src/webparts/folderManager/components/FolderManager.tsx` — folder creation branch

- [ ] **Step 1: Add the column to both libraries**

On `Staging` and `Documents`: single line of text, display name `Full Name`. The internal name will most likely be `Full_x0020_Name` — **verify** against `/_api/web/lists/getbytitle('Staging')/fields` rather than assuming. CLAUDE.md is explicit that internal names must not be guessed from display names.

Then remove it from the default **view** — not from the form. The details pane renders form fields, so a column hidden on the form disappears from the pane, while one merely absent from the view still shows.

- [ ] **Step 2: Write it at creation**

Where a folder is created, set the field on the folder's list item to the term's full label. `t.label` already carries the raw term text — the code deliberately keeps the display label unsanitised while the path segment is sanitised, so use the raw label here.

For the segment container, use the mode's label rather than the abbreviation.

- [ ] **Step 3: Verify**

Reconcile into an empty tree, then select a folder and open the details pane. Expect `Name` = the abbreviation and `Full Name` = the full label. Confirm the grid does not show the column.

- [ ] **Step 4: Commit**

```bash
git add src/webparts/folderManager/components/FolderManager.tsx
git commit -m "feat: write the full term label to Full Name on each folder"
```

---

## Task 6: Prune orphaned abbreviation rows

**Files:**
- Modify: `src/webparts/folderManager/components/FolderManager.tsx` — the existing prune step

- [ ] **Step 1: Extend prune to the abbreviation list**

An abbreviation row whose term no longer exists is an orphan. Report it in the same place `DMS Folder Map` orphans are reported, and delete only under the same conditions.

**Reuse the existing `incomplete` guard.** The comment at line 1029 already states the rule: prune decides a row is orphaned by asking whether its term is in the enumerated set, so a segment that silently returned empty would make every one of its rows look deleted. Abbreviation rows carry exactly the same hazard — a transient term-store failure must never be read as "the client deleted 133 terms."

So: skip abbreviation pruning entirely whenever `incomplete` is non-empty, exactly as folder-map pruning already does.

- [ ] **Step 2: Report the folder too**

An orphaned abbreviation usually means the term was deleted while its folder and documents still exist. Report the folder alongside the row so an admin can decide what to do with the contents — deleting the row does not delete the folder.

- [ ] **Step 3: Verify the build and commit**

```bash
npx heft test
git add src/webparts/folderManager/components/FolderManager.tsx
git commit -m "feat: prune orphaned abbreviation rows under the incomplete guard"
```

---

## Task 7: Messages and client guidance

**Files:**
- Modify: `src/webparts/form/components/Form.tsx`, `src/webparts/bulkUpload/components/BulkUpload.tsx` — upload failure text
- Modify: `CLAUDE.md`
- Create: `docs/client/folder-abbreviations-guide.md`

- [ ] **Step 1: Correct the upload failure message**

Both web parts currently tell the user to re-run reconciliation when no `DMS Folder Map` row resolves. Re-running does not help when the cause is a missing abbreviation. The message must name both possibilities and say who fixes it — for example: *"This unit has no folder yet. Your DMS administrator needs to add an abbreviation for it in DMS Term Abbreviation and re-run folder reconciliation."*

Same class of defect as the "not fully provisioned" banner: the code changed, the message did not.

- [ ] **Step 2: Update CLAUDE.md**

Under folder routing, record that folder names come from `DMS Term Abbreviation` keyed by term GUID (segments from `StagingFolder` in `DMS Config`), that term labels stay full and drive the dropdown, that a missing abbreviation skips the term, and that abbreviations must be unique among siblings or two units merge into one ACL.

- [ ] **Step 3: Write the client guide**

Create `docs/client/folder-abbreviations-guide.md` covering: what an abbreviation controls; that it must be unique among its siblings; that a unit abbreviation must include its sub-group prefix (with the two Strategic Communications units as the worked example); that changing one renames a live folder; that a missing one blocks that unit's uploads entirely; and that the full name stays visible in the details panel.

Keep it short and non-technical — the reader is a client administrator, not a developer.

- [ ] **Step 4: Commit**

```bash
git add src/webparts/form/components/Form.tsx src/webparts/bulkUpload/components/BulkUpload.tsx CLAUDE.md docs/client/folder-abbreviations-guide.md
git commit -m "docs: correct upload messages and add the client abbreviation guide"
```

---

## Task 8: Live verification

Not automatable — needs the tenant. Every check needs a hard refresh; config is read once on mount.

- [ ] **Step 1** — Deploy: bump the version in `config/package-solution.json`, `npm run build`, upload the `.sppkg`, hard-refresh.
- [ ] **Step 2** — Delete the existing Staging and Documents folder trees. The project is pre-production with no real documents, so a clean rebuild is the migration.
- [ ] **Step 3** — Reconcile. Expect `GHO/GCA/GMB_STRATCOMMS` shaped paths, and `Full Name` populated on every folder.
- [ ] **Step 4** — Upload into a unit and confirm the file lands under the abbreviated path.
- [ ] **Step 5** — **Rename test.** Change one abbreviation in the list, re-reconcile, confirm the folder is renamed and its documents come with it. Then upload again to confirm the map row still resolves.
- [ ] **Step 6** — **Conflict test.** Create an empty folder by hand at the target name, then change an abbreviation to match it. Re-reconcile and confirm the conflict is reported, not skipped, and that nothing was created or merged.
- [ ] **Step 7** — **Missing-abbreviation test.** Delete one row, re-reconcile, confirm the term is reported as skipped and no folder is created. Then upload to that unit and confirm the message names the real cause.
- [ ] **Step 8** — **Collision test.** Set two sibling units to the same abbreviation and confirm the run aborts before creating anything.

---

## Out of scope

- The admin UI (spec §7) — Phase 2, separate plan.
- Migrating a populated tree. Pre-production only; **must be revisited before anything ships to `sdguthrie`.**
- Restructuring the term store to make sub-group a real fourth level.

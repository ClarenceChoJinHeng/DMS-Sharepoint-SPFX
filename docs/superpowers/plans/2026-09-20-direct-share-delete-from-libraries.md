# Direct Share/Delete from the Four Approved-Side Libraries — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let someone click a file's Name column in `Documents`, `HC Documents`, `Archive`, or
`HC Archive` and land on a read-only file view — with direct Share (Approver, System Admin) and
Delete (System Admin only) buttons when their rights allow it — reusing the existing `Requests`
page and `FileDetailPanel` component rather than building a new page.

**Architecture:** Add a `?file=<UniqueId>` mode to `Requests.tsx`, sibling to its existing
`?request=<Id>` mode. `GetFileById` (web-scoped) resolves the file regardless of which of the four
libraries it is in, so no library parameter is needed. `probeFileRights` gates the two buttons;
Archive always overrides to read-only. Widen `Requests.aspx`'s page-access policy to include `UPL`
and `UPLHC` so a plain uploader can reach the new mode at all.

**Tech Stack:** TypeScript, React (SPFx web part), SharePoint REST (`_api/web`), Jest for the one
new pure module.

**Spec:** `docs/superpowers/specs/2026-09-20-direct-share-delete-from-libraries-design.md`

---

### Task 1: `directActionsFor` — the rights-tier + archive-override rule

**Files:**
- Create: `src/shared/directFileActions.ts`
- Test: `src/shared/directFileActions.test.ts`

This is the one piece of genuinely new decision logic in the feature, and the only part of it this
project's existing testing style covers (pure modules, no UI tests anywhere in the codebase).

- [ ] **Step 1: Write the failing tests**

```typescript
// src/shared/directFileActions.test.ts
import { directActionsFor } from "./directFileActions";
import { FileRights } from "./dmsFolderMap";

const rights = (remove: FileRights["remove"], share: FileRights["share"]): FileRights => ({
  remove,
  share,
});

describe("directActionsFor", () => {
  it("offers both when the probe grants both and the file is not archived", () => {
    expect(directActionsFor(rights("granted", "granted"), false)).toEqual({
      canDelete: true,
      canShare: true,
    });
  });

  it("offers share only when the probe denies delete but grants share (the ordinary Approver)", () => {
    expect(directActionsFor(rights("denied", "granted"), false)).toEqual({
      canDelete: false,
      canShare: true,
    });
  });

  it("offers neither when the probe denies both (a plain uploader)", () => {
    expect(directActionsFor(rights("denied", "denied"), false)).toEqual({
      canDelete: false,
      canShare: false,
    });
  });

  it("offers neither on 'missing' or 'unknown', same as 'denied' — never assume a right exists", () => {
    expect(directActionsFor(rights("missing", "unknown"), false)).toEqual({
      canDelete: false,
      canShare: false,
    });
  });

  it("ARCHIVE OVERRIDES EVERYTHING — both suppressed even when the probe grants both", () => {
    // This is the whole reason the function exists rather than a one-line inline check:
    // probeFileRights only answers permission, not appropriateness, and an archived document
    // is read-only to everybody by design regardless of what Full Control would otherwise allow.
    expect(directActionsFor(rights("granted", "granted"), true)).toEqual({
      canDelete: false,
      canShare: false,
    });
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx jest src/shared/directFileActions.test.ts`
Expected: FAIL — `Cannot find module './directFileActions'`

- [ ] **Step 3: Write the implementation**

```typescript
// src/shared/directFileActions.ts
import { FileRights } from "./dmsFolderMap";

/**
 * Should THIS viewer see a direct Delete/Share button for THIS document, reached by clicking a
 * file's Name column in Documents/HC Documents/Archive/HC Archive?
 *
 * Spec: docs/superpowers/specs/2026-09-20-direct-share-delete-from-libraries-design.md
 *
 * ⚠ ARCHIVE OVERRIDES THE PROBE OUTRIGHT, REGARDLESS OF WHAT IT ANSWERS. `probeFileRights`'s own
 * documentation says it answers permission only, never appropriateness — "an archived document is
 * read-only to everybody by design, and the caller must enforce that." A system admin's Full
 * Control would otherwise make the probe say "granted" on an archived file; this function is what
 * stops that from ever reaching a button.
 *
 * Only a literal "granted" offers an action — "denied", "missing", and "unknown" are all treated
 * identically as "no", never as "probably fine." Understating what someone can do costs them one
 * extra click through the ordinary request flow elsewhere; overstating it would offer a delete or
 * share that then fails at the write, which is the worse of the two directions.
 */
export interface DirectActions {
  canDelete: boolean;
  canShare: boolean;
}

export function directActionsFor(
  rights: FileRights,
  isArchived: boolean,
): DirectActions {
  if (isArchived) return { canDelete: false, canShare: false };
  return {
    canDelete: rights.remove === "granted",
    canShare: rights.share === "granted",
  };
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx jest src/shared/directFileActions.test.ts`
Expected: PASS — 5 tests

- [ ] **Step 5: Typecheck and commit**

Run: `npx tsc --noEmit`
Expected: no errors

```bash
git add src/shared/directFileActions.ts src/shared/directFileActions.test.ts
git commit -m "feat: add directActionsFor — rights-tier + archive-override rule for direct file actions"
```

---

### Task 2: Widen the Requests page's role policy to include `UPL`/`UPLHC`

**Files:**
- Modify: `src/shared/pageAccessPolicy.ts`
- Modify: `src/shared/pageAccessPolicy.test.ts`

Found during spec self-review: `Requests.aspx`'s current policy is `["APR", "APRHC", "DEPTVIEW"]`,
with `UPL` deliberately removed on 2026-08-21 because a plain uploader's queue there was always
empty. This feature gives them a genuine reason to reach the page for the first time (the read-only
`?file=` view), so the policy must widen or they get AccessDenied outright.

- [ ] **Step 1: Find the exact current block to edit**

Read `src/shared/pageAccessPolicy.ts` around the `/request/i` rule (it currently reads, in full):

```typescript
  {
    // BEFORE the approver rule, deliberately: a name like "Approval-Requests.aspx" would otherwise
    // land on /approv/i, and this page's audience is wider than that rule's by one role. It would
    // otherwise fall to DEFAULT_POLICY, which is APR + UPL + DELS; explicit and narrow beats
    // right-by-accident.
    //
    // ⚠ `UPL` WAS HERE UNTIL 2026-08-21 AND ITS REMOVAL IS THE POINT. The 2026-08-15 design gave this
    // page two audiences — the uploader raising a request and the Head of Unit deciding it — and that
    // reasoning went stale on 2026-08-20, when the requester's own view moved to My Submissions →
    // Requests (with Cancel). A PIC opening this page now gets a screen filtered to units where they
    // hold APR, i.e. none: an empty page on their menu.
    //
    // `DEPTVIEW` joins because a Head of Department holds DEL and SHARE since 1.0.197.0 and can
    // therefore carry out an approved-document deletion or share outright (client, 2026-08-21: *"I
    // also include HOD is because they literally have Share and Deletion power"*). They see their
    // department's APPROVED-stage requests only — see `ViewerScope` in shared/requests.ts for why
    // pending ones are hidden rather than merely disabled.
    //
    // Widening this list cannot grant anyone a new place to act: the page grant opens the SCREEN,
    // while `canDecide` decides each row, and the approval itself runs in the viewer's own session
    // and fails loudly if their permissions do not cover it.
    // Spec: docs/superpowers/specs/2026-08-21-requests-page-hod-access-design.md
    match: /request/i,
    policy: {
      roles: ["APR", "APRHC", "DEPTVIEW"],
      adminOnly: false,
      reason: "Approver and Head of Department groups are listed — the Head of Unit decides deletion and share requests, and a Head of Department can carry out those on approved documents. Uploaders raise requests on My Submissions, not here.",
    },
  },
```

- [ ] **Step 2: Replace it with the widened version**

```typescript
  {
    // BEFORE the approver rule, deliberately: a name like "Approval-Requests.aspx" would otherwise
    // land on /approv/i, and this page's audience is wider than that rule's by one role. It would
    // otherwise fall to DEFAULT_POLICY, which is APR + UPL + DELS; explicit and narrow beats
    // right-by-accident.
    //
    // ⚠ `UPL`/`UPLHC` WERE REMOVED 2026-08-21 AND ADDED BACK 2026-09-20 — SUPERSEDING THAT REMOVAL,
    // NOT CONTRADICTING IT BY ACCIDENT. The 2026-08-21 removal was correct for what was true then: a
    // plain uploader's queue here was always empty, since their own requests had just moved to My
    // Submissions. This page has since gained a SECOND job — `?file=<UniqueId>` opens a read-only
    // file view reachable by clicking a Name column in Documents/HC Documents/Archive/HC Archive,
    // which every uploader can browse — so leaving them off this list means AccessDenied on the
    // whole page instead of the intended read-only view.
    // Spec: docs/superpowers/specs/2026-09-20-direct-share-delete-from-libraries-design.md
    //
    // `DEPTVIEW` joins because a Head of Department holds DEL and SHARE since 1.0.197.0 and can
    // therefore carry out an approved-document deletion or share outright (client, 2026-08-21: *"I
    // also include HOD is because they literally have Share and Deletion power"*). They see their
    // department's APPROVED-stage requests only — see `ViewerScope` in shared/requests.ts for why
    // pending ones are hidden rather than merely disabled.
    //
    // Widening this list cannot grant anyone a new place to act: the page grant opens the SCREEN,
    // while `canDecide`/`directActionsFor` decide each row, and every action runs in the viewer's
    // own session and fails loudly if their permissions do not cover it.
    // Spec: docs/superpowers/specs/2026-08-21-requests-page-hod-access-design.md
    match: /request/i,
    policy: {
      roles: ["APR", "APRHC", "DEPTVIEW", "UPL", "UPLHC"],
      adminOnly: false,
      reason: "Approver and Head of Department groups can decide deletion and share requests here; uploader groups can open a file's read-only detail view when reached by clicking it directly in a library.",
    },
  },
```

- [ ] **Step 3: Update the near-top comment tracking this pattern**

Find (near the top of the file, above the `RULES` array):
```typescript
// (upload form → APR, Requests → DEPTVIEW, My Submissions → UPLHC, and this).
```

Replace with:
```typescript
// (upload form → APR, Requests → DEPTVIEW then UPL/UPLHC again, My Submissions → UPLHC, and this).
```

- [ ] **Step 4: Update the pinned tests**

In `src/shared/pageAccessPolicy.test.ts`, find this block:

```typescript
describe("the Requests page", () => {
  // Spec: docs/superpowers/specs/2026-08-21-requests-page-hod-access-design.md
  it("offers approvers and Heads of Department — the two roles that can carry a request out", () => {
    expect(policyForPage("Requests.aspx").roles).toEqual(["APR", "APRHC", "DEPTVIEW"]);
    expect(policyForPage("CRS-Requests.aspx").roles).toEqual(["APR", "APRHC", "DEPTVIEW"]);
  });

  it("NO LONGER offers uploaders — their view moved to My Submissions on 2026-08-20", () => {
    // Regression guard for the stale-reasoning bug this rule carried for six days: a PIC granted
    // this page gets a queue filtered to units where they hold APR, i.e. none. An empty page on
    // their menu, granted by a comment describing a design that had been superseded.
    for (const name of ["Requests.aspx", "CRS-Requests.aspx", "Approval-Requests.aspx"]) {
      expect(policyForPage(name).roles).not.toContain("UPL");
    }
  });

  it("beats the approver rule, so a name carrying 'approval' still reaches a Head of Department", () => {
    // Ordering, pinned: on the /approv/ rule this page would list approver groups ONLY, and a Head
    // of Department — who holds DEL and SHARE and can act on approved documents — could not open it.
    expect(policyForPage("Approval-Requests.aspx").roles).toEqual(["APR", "APRHC", "DEPTVIEW"]);
  });

  it("does not disturb the approver's own page — DEPTVIEW must never reach the approval queue", () => {
    // The asymmetry is deliberate. A Head of Department is view-and-act-on-approved-documents;
    // letting them into the queue would let them publish into a unit they do not run.
    expect(policyForPage("ApprovalDocument.aspx").roles).toEqual(["APR", "APRHC"]);
  });

  it("is not an administrator tool — a Head of Unit must be able to be granted it", () => {
    expect(policyForPage("Requests.aspx").adminOnly).toBe(false);
  });
});
```

Replace with:

```typescript
describe("the Requests page", () => {
  // Spec: docs/superpowers/specs/2026-08-21-requests-page-hod-access-design.md
  it("offers approvers, Heads of Department, and uploaders", () => {
    const expected = ["APR", "APRHC", "DEPTVIEW", "UPL", "UPLHC"];
    expect(policyForPage("Requests.aspx").roles).toEqual(expected);
    expect(policyForPage("CRS-Requests.aspx").roles).toEqual(expected);
  });

  it("OFFERS UPLOADERS AGAIN, 2026-09-20 — reversing the 2026-08-21 removal for a new reason", () => {
    // The 2026-08-21 removal was right for what was true then: a plain uploader's queue here was
    // always empty. This page has since gained a second job — the ?file=<UniqueId> read-only view
    // reachable from a library click, which every uploader can trigger — so they need the page
    // grant again, for a genuinely different purpose than the one that got them removed.
    // Spec: docs/superpowers/specs/2026-09-20-direct-share-delete-from-libraries-design.md
    for (const name of ["Requests.aspx", "CRS-Requests.aspx", "Approval-Requests.aspx"]) {
      expect(policyForPage(name).roles).toContain("UPL");
      expect(policyForPage(name).roles).toContain("UPLHC");
    }
  });

  it("beats the approver rule, so a name carrying 'approval' still reaches a Head of Department", () => {
    // Ordering, pinned: on the /approv/ rule this page would list approver groups ONLY, and a Head
    // of Department — who holds DEL and SHARE and can act on approved documents — could not open it.
    expect(policyForPage("Approval-Requests.aspx").roles).toEqual([
      "APR",
      "APRHC",
      "DEPTVIEW",
      "UPL",
      "UPLHC",
    ]);
  });

  it("does not disturb the approver's own page — DEPTVIEW must never reach the approval queue", () => {
    // The asymmetry is deliberate. A Head of Department is view-and-act-on-approved-documents;
    // letting them into the queue would let them publish into a unit they do not run.
    expect(policyForPage("ApprovalDocument.aspx").roles).toEqual(["APR", "APRHC"]);
  });

  it("is not an administrator tool — a Head of Unit must be able to be granted it", () => {
    expect(policyForPage("Requests.aspx").adminOnly).toBe(false);
  });
});
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `npx jest src/shared/pageAccessPolicy.test.ts`
Expected: PASS — no failures

- [ ] **Step 6: Run the FULL suite to catch any other test asserting the old array**

Run: `npx jest`
Expected: PASS. If any OTHER test file asserts `Requests.aspx`'s roles equal exactly
`["APR", "APRHC", "DEPTVIEW"]` (search for it with `npx jest -t "Requests"` first if unsure), update
it the same way — append `"UPL", "UPLHC"` to the expected array.

- [ ] **Step 7: Typecheck and commit**

Run: `npx tsc --noEmit`
Expected: no errors

```bash
git add src/shared/pageAccessPolicy.ts src/shared/pageAccessPolicy.test.ts
git commit -m "fix: widen Requests page policy to include UPL/UPLHC for the new read-only file view"
```

---

### Task 3: URL param plumbing — detect `?file=<UniqueId>` on `Requests.tsx`

**Files:**
- Modify: `src/webparts/requests/components/Requests.tsx`

This step only adds the ref/state scaffolding that reads and strips the URL parameter, mirroring
the existing `linkedRequest`/`linkRead` pattern exactly. No loading or rendering yet — that comes
in Tasks 4-6. This keeps each commit small and independently verifiable.

- [ ] **Step 1: Add the new imports**

Find this import block near the top of `Requests.tsx`:

```typescript
import { closeOnBackdrop } from "../../../shared/backdropClose";
import { displayNameFor } from "../../../shared/displayName";
// The file view is SHARED with My Submissions (client, 2026-09-10) - one component, two mounts.
import { FileDetailPanel } from "../../../shared/fileDetailPanel";
```

Replace with:

```typescript
import { closeOnBackdrop } from "../../../shared/backdropClose";
import { displayNameFor } from "../../../shared/displayName";
// The file view is SHARED with My Submissions (client, 2026-09-10) - one component, two mounts.
import { FileDetailPanel } from "../../../shared/fileDetailPanel";
// ── The direct file view (2026-09-20) — see the module's own comment for why this is a SIBLING to
// loadFileView rather than a reuse of it: that one resolves a file BEHIND A REQUEST, with a stamp
// fallback for a target that may have moved since the request was raised. A file reached by
// clicking it directly in Documents/HC Documents/Archive/HC Archive has not moved since a request
// was raised, because no request is involved at all.
import { probeFileRights, FileRights } from "../../../shared/dmsFolderMap";
import { documentUnit } from "../../../shared/documentDetails";
import { directActionsFor, DirectActions } from "../../../shared/directFileActions";
import { searchTenantPeople, PersonPick } from "../../../shared/spGroups";
```

- [ ] **Step 2: Add the ref/state scaffolding**

Find this block (right after `linkRead`'s `if` block):

```typescript
  const [viewId, setViewId] = useState<number | undefined>(undefined);
  const [fileView, setFileView] = useState<FileView | undefined>(undefined);
  /* A superseded read must not win - two can be in flight when the viewed row's status changes. */
  const viewSeq = useRef(0);
  /* Read ONCE from the address bar, and stripped, so a refresh does not re-open a closed view. */
  const linkedRequest = useRef<number | undefined>(undefined);
  const linkRead = useRef(false);
  if (!linkRead.current) {
    linkRead.current = true;
    const n = Number(
      new URLSearchParams(window.location.search).get("request") ?? "",
    );
    if (isFinite(n) && n > 0) linkedRequest.current = n;
  }
```

Replace with:

```typescript
  const [viewId, setViewId] = useState<number | undefined>(undefined);
  const [fileView, setFileView] = useState<FileView | undefined>(undefined);
  /* A superseded read must not win - two can be in flight when the viewed row's status changes. */
  const viewSeq = useRef(0);
  /* Read ONCE from the address bar, and stripped, so a refresh does not re-open a closed view. */
  const linkedRequest = useRef<number | undefined>(undefined);
  const linkRead = useRef(false);
  if (!linkRead.current) {
    linkRead.current = true;
    const n = Number(
      new URLSearchParams(window.location.search).get("request") ?? "",
    );
    if (isFinite(n) && n > 0) linkedRequest.current = n;
  }

  /* ── The direct file view (2026-09-20, no request involved) ──
     Reached by clicking a file's Name column in Documents/HC Documents/Archive/HC Archive, via
     column formatting pointing at `?file=<UniqueId>` — see the design doc for why no library
     parameter is needed: GetFileById is web-scoped and a UniqueId is unique site-wide.
     Same read-once-and-strip pattern as `linkedRequest` above, and the same reason: a refresh must
     not re-trigger opening something that was already read from the address bar once. */
  const linkedFile = useRef<string | undefined>(undefined);
  const fileLinkRead = useRef(false);
  if (!fileLinkRead.current) {
    fileLinkRead.current = true;
    const f = (new URLSearchParams(window.location.search).get("file") ?? "").trim();
    if (f.length > 0) linkedFile.current = f;
  }
  const [directFileId, setDirectFileId] = useState<string | undefined>(undefined);
  useEffect(() => {
    const id = linkedFile.current;
    if (id === undefined) return;
    linkedFile.current = undefined;
    setDirectFileId(id);
    try {
      const u = new URL(window.location.href);
      u.searchParams.delete("file");
      window.history.replaceState(window.history.state, "", u.toString());
    } catch {
      /* an address bar we cannot tidy costs nothing */
    }
  }, []);
```

- [ ] **Step 3: Typecheck**

Run: `npx tsc --noEmit`
Expected: unused-variable warnings for `probeFileRights`, `documentUnit`, `directActionsFor`,
`searchTenantPeople`, `PersonPick`, `FileRights`, `DirectActions` — these are consumed in Tasks 4-8.
No TYPE errors. If your linter fails the build on unused imports (check by running the next step),
skip lint until Task 8 is done, or temporarily prefix each with `_` — do NOT delete the imports.

Run: `npx eslint src/webparts/requests/components/Requests.tsx`
Expected: warnings only for the unused imports listed above (same pre-existing pattern this
project already tracks as a documented baseline elsewhere) — no errors.

- [ ] **Step 4: Commit**

```bash
git add src/webparts/requests/components/Requests.tsx
git commit -m "feat: read ?file=<UniqueId> from the address bar on the Requests page"
```

---

### Task 4: The file resolution loader

**Files:**
- Modify: `src/webparts/requests/components/Requests.tsx`

- [ ] **Step 1: Add the loader function**

Find the end of the existing `loadFileView` function (it ends with):

```typescript
    } catch (e) {
      settle({
        state: "gone",
        message: `The document could not be read: ${(e as Error).message}`,
      });
    }
  };
```

Immediately after that closing `};`, add:

```typescript

  /**
   * Resolve a file reached DIRECTLY (no request involved) — `?file=<UniqueId>` from a library's
   * Name-column click. Mirrors `loadFileView`'s shape exactly, minus the stamp fallback: a file
   * already sitting in one of the four target libraries has not moved since the click.
   */
  const [directFileView, setDirectFileView] = useState<FileView | undefined>(
    undefined,
  );
  const directViewSeq = useRef(0);
  const loadDirectFileView = async (uniqueId: string): Promise<void> => {
    const seq = ++directViewSeq.current;
    const settle = (v: FileView): void => {
      if (directViewSeq.current === seq) setDirectFileView(v);
    };
    setDirectFileView({ state: "loading" });
    const read = (url: string): Promise<SPHttpClientResponse> =>
      context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: GET,
      });
    try {
      const res = await read(
        `${siteUrl}/_api/web/GetFileById(guid'${encodeURIComponent(uniqueId)}')?$select=ServerRelativeUrl${bust()}`,
      );
      if (!res.ok) {
        settle({
          state: "gone",
          message:
            res.status === 404
              ? "This document is no longer in the library - it may already have been deleted."
              : `The document could not be read (HTTP ${res.status}). Refresh and try again.`,
        });
        return;
      }
      const path = ((await res.json()) as { ServerRelativeUrl?: string })
        .ServerRelativeUrl;
      if (!path) {
        settle({
          state: "gone",
          message: "This document could not be located.",
        });
        return;
      }
      /* Parameter alias, never an inline literal - a deep path answers 400 otherwise (gotcha #9). */
      const alias = `@f='${encodeServerRelativePath(path)}'`;
      const base = `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)`;
      const fr = await read(
        `${base}?$select=Name,Length,TimeLastModified,ServerRelativeUrl&${alias}${bust()}`,
      );
      if (!fr.ok) {
        settle({
          state: "gone",
          message:
            fr.status === 404
              ? "This document is no longer in the library - it may already have been deleted."
              : `The document could not be read (HTTP ${fr.status}). Refresh and try again.`,
        });
        return;
      }
      const f = (await fr.json()) as {
        Name?: string;
        Length?: string;
        TimeLastModified?: string;
        ServerRelativeUrl?: string;
      };
      let fieldText: Record<string, string> = {};
      try {
        const t = await read(
          `${base}/ListItemAllFields/FieldValuesAsText?${alias}${bust()}`,
        );
        if (t.ok) fieldText = (await t.json()) as Record<string, string>;
      } catch {
        /* keep {} */
      }
      /* Document Date re-read RAW and formatted locally, as My Submissions does. */
      try {
        const raw = await read(
          `${base}/ListItemAllFields?$select=DocumentDate&${alias}${bust()}`,
        );
        if (raw.ok) {
          const iso = ((await raw.json()) as { DocumentDate?: string })
            .DocumentDate;
          const d = iso ? new Date(iso) : undefined;
          if (d && !isNaN(d.getTime()))
            fieldText.DocumentDate = formatSubmittedOn(d);
        }
      } catch {
        /* keep SharePoint's own string */
      }
      const modified = f.TimeLastModified
        ? new Date(f.TimeLastModified)
        : undefined;
      settle({
        state: "ready",
        name: f.Name || "document",
        fileRef: f.ServerRelativeUrl || path,
        size: f.Length,
        modified: modified && !isNaN(modified.getTime()) ? modified : undefined,
        fieldText,
      });
    } catch (e) {
      settle({
        state: "gone",
        message: `The document could not be read: ${(e as Error).message}`,
      });
    }
  };

  /* ONE ROUTE IN, ONE LOADER — mirrors the existing rule right above for the request-based mode. */
  useEffect(() => {
    if (directFileId === undefined) return;
    loadDirectFileView(directFileId).catch(() => undefined);
  }, [directFileId]);
```

- [ ] **Step 2: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors (the `probeFileRights`/`documentUnit`/`directActionsFor`/`searchTenantPeople`/
`PersonPick`/`FileRights`/`DirectActions` imports are still unused — that's expected until Task 5-8)

- [ ] **Step 3: Commit**

```bash
git add src/webparts/requests/components/Requests.tsx
git commit -m "feat: resolve a file directly by UniqueId, no request involved"
```

---

### Task 5: Rights probing — three tiers, Archive always read-only

**Files:**
- Modify: `src/webparts/requests/components/Requests.tsx`

- [ ] **Step 1: Add the rights-probing effect**

Immediately after the `useEffect` you added at the end of Task 4 (the one that calls
`loadDirectFileView`), add:

```typescript

  /**
   * Once the file resolves, probe what THIS viewer can do to it directly, and classify it as
   * archived or not — both feed `directActionsFor`. Declared as its own effect, keyed on the
   * resolved file's own identity, so it re-probes if a different file is opened without a full
   * page reload.
   */
  const [directRights, setDirectRights] = useState<FileRights | undefined>(
    undefined,
  );
  const [directIsArchived, setDirectIsArchived] = useState(false);
  useEffect(() => {
    if (directFileId === undefined || directFileView?.state !== "ready") {
      setDirectRights(undefined);
      return;
    }
    setDirectRights(undefined);
    const archiveSegs = {
      normal: cachedArchiveLibraries()?.normal.urlSegment,
      hc: cachedArchiveLibraries()?.hc?.urlSegment,
    };
    const allSegs = libraryTargets().map((t) => t.urlSegment);
    const seg = librarySegmentOf(directFileView.fileRef, allSegs);
    setDirectIsArchived(isArchivedRow(seg ?? "", archiveSegs));
    probeFileRights(context.spHttpClient, siteUrl, directFileId)
      .then(setDirectRights)
      .catch(() =>
        setDirectRights({ remove: "unknown", share: "unknown" }),
      );
  }, [directFileId, directFileView?.state]);

  const directActions: DirectActions | undefined =
    directRights === undefined
      ? undefined
      : directActionsFor(directRights, directIsArchived);
```

- [ ] **Step 2: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors. `searchTenantPeople`/`PersonPick`/`documentUnit` remain unused until Task 8.

- [ ] **Step 3: Commit**

```bash
git add src/webparts/requests/components/Requests.tsx
git commit -m "feat: probe direct file rights, with Archive always overriding to read-only"
```

---

### Task 6: The render branch — read-only view, "Checking…" state, Back

**Files:**
- Modify: `src/webparts/requests/components/Requests.tsx`

- [ ] **Step 1: Add the render branch**

Find the existing `?request=<Id>` render branch — it starts with:

```typescript
  /* ── The file view ── */
  if (viewId !== undefined) {
```

Immediately BEFORE that line, add a new branch for the direct-file mode:

```typescript
  /* ── The direct file view (no request) ── */
  if (directFileId !== undefined) {
    const tenantRoot = siteUrl.replace(/^(https?:\/\/[^/]+).*$/, "$1");
    const libSegments = libraryTargets().map((t) => t.urlSegment);
    const hcSegs = {
      approval: cachedHcLibraries()?.approval.urlSegment,
      documents: cachedHcLibraries()?.documents.urlSegment,
    };
    return (
      <section style={s.wrap}>
        <div style={s.backBand}>
          <button
            type="button"
            style={s.backLink}
            onClick={() => window.history.back()}
          >
            ‹ Back
          </button>
        </div>
        {notice && (
          <div ref={noticeRef} style={noticeBad ? s.warn : s.ok}>
            {notice}
          </div>
        )}
        {directFileView === undefined || directFileView.state === "loading" ? (
          <p style={s.quiet}>Reading the document&hellip;</p>
        ) : directFileView.state === "gone" ? (
          <div style={s.warn}>{directFileView.message}</div>
        ) : (
          <>
            <div style={s.rowTop}>
              <span style={s.name}>{directFileView.name}</span>
              {isHcRow(
                librarySegmentOf(directFileView.fileRef, libSegments) ?? "",
                hcSegs,
              ) && <span style={s.hcTag}>HC</span>}
            </div>
            <div style={{ marginTop: 12, marginBottom: 12 }}>
              {directActions === undefined ? (
                <p style={s.quiet}>Checking what you can do with this file&hellip;</p>
              ) : (
                <div style={s.askBar}>
                  {directActions.canDelete && (
                    <button
                      style={s.askBtn}
                      onClick={() => setDirectDeleteConfirm(true)}
                    >
                      Delete
                    </button>
                  )}
                  {directActions.canShare && (
                    <button
                      style={s.askBtn}
                      onClick={() => {
                        setDirectShareRecipients([]);
                        setDirectShareQuery("");
                        setDirectShareResults([]);
                        setDirectShareError(undefined);
                        setDirectShareOpen(true);
                      }}
                    >
                      Share
                    </button>
                  )}
                  {!directActions.canDelete && !directActions.canShare && (
                    <span style={{ fontSize: 12, color: "#605e5c" }}>
                      You do not have permission to delete or share this
                      document directly.
                    </span>
                  )}
                </div>
              )}
            </div>
            <div style={{ marginTop: 20 }}>
              <FileDetailPanel
                name={directFileView.name}
                fileRef={directFileView.fileRef}
                tenantRoot={tenantRoot}
                siteUrl={siteUrl}
                fieldText={directFileView.fieldText}
                location={trailText(
                  folderTrail(directFileView.fileRef, libSegments),
                )}
                size={directFileView.size}
                modified={directFileView.modified}
              />
            </div>
          </>
        )}
        {directDeleteDialog}
        {directShareDialog}
      </section>
    );
  }

```

- [ ] **Step 2: Add the four new state variables the JSX above references**

These need to exist before the render branch runs. Add them right next to `directRights`/
`directIsArchived` from Task 5 (immediately after the `directActions` derivation you added there):

```typescript
  const [directDeleteConfirm, setDirectDeleteConfirm] = useState(false);
  const [directDeleteBusy, setDirectDeleteBusy] = useState(false);
  const [directDeleteError, setDirectDeleteError] = useState<
    string | undefined
  >(undefined);
  const [directShareOpen, setDirectShareOpen] = useState(false);
  const [directShareQuery, setDirectShareQuery] = useState("");
  const [directShareResults, setDirectShareResults] = useState<PersonPick[]>(
    [],
  );
  const [directShareRecipients, setDirectShareRecipients] = useState<
    string[]
  >([]);
  const [directShareBusy, setDirectShareBusy] = useState(false);
  const [directShareError, setDirectShareError] = useState<
    string | undefined
  >(undefined);
  const [directShareDone, setDirectShareDone] = useState(false);
  /* Placeholders until Tasks 7-8 define the real dialogs — an undefined value renders nothing. */
  const directDeleteDialog: React.ReactNode = undefined;
  const directShareDialog: React.ReactNode = undefined;
```

- [ ] **Step 3: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors. `documentUnit` remains unused until Task 7.

- [ ] **Step 4: Commit**

```bash
git add src/webparts/requests/components/Requests.tsx
git commit -m "feat: render the direct file view — read-only, with a Checking state for the buttons"
```

---

### Task 7: Direct delete — writer + confirm dialog

**Files:**
- Modify: `src/webparts/requests/components/Requests.tsx`

- [ ] **Step 1: Add the writer function**

Add this immediately after `loadDirectFileView` (from Task 4), before the "ONE ROUTE IN, ONE
LOADER" `useEffect`:

```typescript

  /**
   * Write a self-approved deletion request for a file opened directly (no My Submissions record,
   * since a file already sitting in Documents/HC Documents/Archive/HC Archive was not necessarily
   * uploaded through this app). Mirrors `MySubmissions.tsx`'s `writeApprovedDeletionRequest`
   * exactly in shape — Stage is always "approved" here, since every file reachable through this
   * mode already is. `CRS — Execute approved deletion` performs the actual recycle.
   */
  const writeDirectDeletionRequest = async (
    view: Extract<FileView, { state: "ready" }>,
    uniqueId: string,
  ): Promise<string | undefined> => {
    const meEmail = (context.pageContext.user.email ?? "").toLowerCase();
    const where =
      Object.keys(view.fieldText).length > 0
        ? documentUnit(view.fieldText)
        : undefined;
    const now = new Date().toISOString();
    const targetEtag = await readFileETag(
      context.spHttpClient,
      siteUrl,
      uniqueId,
    );
    const body: Record<string, string> = {
      Title: `Deletion — ${view.name}`.slice(0, 255),
      RequestType: "Deletion",
      Status: "Approved",
      Stage: "approved",
      ItemUniqueId: uniqueId,
      ItemName: view.name,
      ItemUrl: view.fileRef,
      Segment: where?.segment ?? "",
      Unit: where?.unit ?? "",
      UnitTermGuid: where?.unitTermGuid ?? "",
      RequestedBy: meEmail,
      RequestedAt: now,
      Reason: "",
      DecidedBy: meEmail,
      DecidedAt: now,
      DecisionNote: "No approval needed — carried out automatically.",
    };
    if ((view.fieldText.SubmissionFileId ?? "").trim().length > 0) {
      body.SubmissionFileId = view.fieldText.SubmissionFileId.trim();
    }
    if (targetEtag !== undefined) {
      body.TargetETag = targetEtag;
    }
    try {
      const send = (
        payload: Record<string, string>,
      ): Promise<SPHttpClientResponse> => post(`${listUrl()}/items`, payload);
      let res = await send(body);
      // ⚠ Same drop-newest-optional-column-first order as MySubmissions.tsx's writer — one
      // unknown field name fails the WHOLE write (gotcha #11).
      if (res.status === 400 && body.TargetETag !== undefined) {
        const without = { ...body };
        delete without.TargetETag;
        res = await send(without);
      }
      if (res.status === 400 && body.SubmissionFileId !== undefined) {
        const without = { ...body };
        delete without.SubmissionFileId;
        delete without.TargetETag;
        res = await send(without);
      }
      if (res.ok) return undefined;
      return res.status === 404
        ? "The requests list does not exist yet — ask an administrator to open the Requests page, which creates it."
        : `The deletion could not be recorded (HTTP ${res.status}).`;
    } catch (e) {
      return `Could not record the deletion: ${(e as Error).message}`;
    }
  };
```

- [ ] **Step 2: Replace the placeholder `directDeleteDialog` from Task 6**

Find (added in Task 6, Step 2):

```typescript
  /* Placeholders until Tasks 7-8 define the real dialogs — an undefined value renders nothing. */
  const directDeleteDialog: React.ReactNode = undefined;
  const directShareDialog: React.ReactNode = undefined;
```

Replace with (keeping `directShareDialog` as a placeholder for one more task):

```typescript
  const directDeleteDialog: React.ReactNode = directDeleteConfirm && (
    <div
      style={s.modalBg}
      onMouseDown={closeOnBackdrop(() => {
        if (!directDeleteBusy) setDirectDeleteConfirm(false);
      })}
    >
      <div style={s.modal} onClick={(e) => e.stopPropagation()}>
        <p style={{ fontSize: 16, fontWeight: 600, margin: "0 0 6px" }}>
          Delete this file?
        </p>
        <p
          style={{
            fontSize: 12.5,
            color: "#605e5c",
            margin: "0 0 8px",
            lineHeight: 1.5,
          }}
        >
          This deletes it straight away — no approver decides this. It moves
          to the recycle bin and can be restored within 93 days.
        </p>
        {directDeleteError && (
          <p style={{ fontSize: 12.5, color: "#a4262c", margin: "0 0 8px" }}>
            {directDeleteError}
          </p>
        )}
        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          <button
            style={s.ghost}
            disabled={directDeleteBusy}
            onClick={() => setDirectDeleteConfirm(false)}
          >
            Cancel
          </button>
          <button
            style={s.primary}
            disabled={directDeleteBusy}
            onClick={() => {
              if (directFileId === undefined) return;
              if (
                directFileView === undefined ||
                directFileView.state !== "ready"
              ) {
                return;
              }
              const view = directFileView;
              const id = directFileId;
              setDirectDeleteBusy(true);
              setDirectDeleteError(undefined);
              writeDirectDeletionRequest(view, id)
                .then((err) => {
                  if (err) {
                    setDirectDeleteError(err);
                    return;
                  }
                  setDirectDeleteConfirm(false);
                  setNoticeBad(false);
                  setNotice(
                    "Deleted. It has moved to the recycle bin and can be restored within 93 days.",
                  );
                })
                .catch((e) => setDirectDeleteError((e as Error).message))
                .finally(() => setDirectDeleteBusy(false));
            }}
          >
            {directDeleteBusy ? "Deleting…" : "Delete"}
          </button>
        </div>
      </div>
    </div>
  );
```

- [ ] **Step 3: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors. If `s.primary` does not exist as a style key in this file's `s` object, use
`s.askBtn` in its place instead (check by running `grep -n "primary:" src/webparts/requests/
components/Requests.tsx` — if it returns nothing, swap `s.primary` for `s.askBtn` above before
proceeding, since `s` is a plain `Record<string, CSSProperties>` and a missing key renders
unstyled with no compile error).

- [ ] **Step 4: Commit**

```bash
git add src/webparts/requests/components/Requests.tsx
git commit -m "feat: direct delete — self-approved CRS Requests row, ETag-guarded, plain confirm dialog"
```

---

### Task 8: Direct share — writer + recipient dialog

**Files:**
- Modify: `src/webparts/requests/components/Requests.tsx`

- [ ] **Step 1: Add the writer function**

Add this immediately after `writeDirectDeletionRequest` (from Task 7):

```typescript

  /**
   * Share a file directly and immediately — no request row at all, matching My Submissions'
   * existing "direct share" path. View-only permission only, no expiry (nothing enforces one on
   * this path today). Runs in the ACTING USER'S own session via SP.Web.ShareObject, so SharePoint
   * itself attributes the invite to them — this is what makes "Approver can share without
   * permission" true: they hold `CRS Share` (Manage Permissions) on the folder already.
   */
  const performDirectShare = async (
    fileRef: string,
    itemName: string,
    recipients: string[],
  ): Promise<string | undefined> => {
    if (recipients.length === 0) return "Add at least one recipient first.";
    const people = recipients.map((e) => ({ Key: e }));
    const res = await post(`${siteUrl}/_api/SP.Web.ShareObject`, {
      url: `${window.location.origin}${fileRef}`,
      peoplePickerInput: JSON.stringify(people),
      roleValue: "role:1073741826", // View — the only option offered on the direct path
      groupId: 0,
      propagateAcl: false,
      sendEmail: true,
      includeAnonymousLinkInEmail: false,
      emailSubject: `A document has been shared with you: ${itemName}`,
      emailBody: "",
      useSimplifiedRoles: true,
    });
    if (!res.ok) {
      if (res.status === 403)
        return "You do not have permission to share that document.";
      return `The document could not be shared (HTTP ${res.status}).`;
    }
    try {
      const body = await res.json();
      const results = (body?.value ?? []) as Array<{
        Status?: boolean;
        Message?: string;
        User?: string;
      }>;
      const failed = results.filter((r) => r && r.Status === false);
      if (failed.length > 0) {
        return failed
          .map((f) => `${f.User ?? "recipient"}: ${f.Message ?? "refused"}`)
          .join("; ");
      }
    } catch {
      /* an unreadable body after a 200 counts as success — the grant is what matters */
    }
    return undefined;
  };
```

- [ ] **Step 2: Add the recipient search effect**

Add this right after the writer function you just added:

```typescript

  /* Debounced tenant people search for the direct-share recipient box. */
  useEffect(() => {
    if (!directShareOpen) return;
    const q = directShareQuery.trim();
    if (q.length < 3) {
      setDirectShareResults([]);
      return;
    }
    let cancelled = false;
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then((people) => {
          if (!cancelled) setDirectShareResults(people.filter((p) => p.email));
        })
        .catch(() => {
          if (!cancelled) setDirectShareResults([]);
        });
    }, 300);
    return () => {
      cancelled = true;
      clearTimeout(t);
    };
  }, [directShareOpen, directShareQuery]);
```

- [ ] **Step 3: Replace the `directShareDialog` placeholder**

Find (still a placeholder from Task 7, Step 2):

```typescript
  const directShareDialog: React.ReactNode = undefined;
```

Replace with:

```typescript
  const directShareDialog: React.ReactNode = directShareOpen && (
    <div
      style={s.modalBg}
      onMouseDown={closeOnBackdrop(() => {
        if (!directShareBusy) setDirectShareOpen(false);
      })}
    >
      <div style={s.modal} onClick={(e) => e.stopPropagation()}>
        {directShareDone ? (
          <>
            <p style={{ fontSize: 16, fontWeight: 600, margin: "0 0 6px" }}>
              Shared
            </p>
            <p style={{ fontSize: 12.5, color: "#605e5c", margin: "0 0 12px" }}>
              The recipient can now view this document.
            </p>
            <button
              style={s.primary}
              onClick={() => {
                setDirectShareOpen(false);
                setDirectShareDone(false);
              }}
            >
              Close
            </button>
          </>
        ) : (
          <>
            <p style={{ fontSize: 16, fontWeight: 600, margin: "0 0 6px" }}>
              Share this file?
            </p>
            <p
              style={{
                fontSize: 12.5,
                color: "#605e5c",
                margin: "0 0 12px",
                lineHeight: 1.5,
              }}
            >
              Shares it now — view-only, no approval needed.
            </p>
            {directShareRecipients.length > 0 && (
              <div
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  gap: 6,
                  marginBottom: 8,
                }}
              >
                {directShareRecipients.map((email) => (
                  <span
                    key={email}
                    style={{
                      ...s.chip,
                      display: "flex",
                      alignItems: "center",
                      gap: 4,
                    }}
                  >
                    {email}
                    {isExternal(email, tenantDomains) && (
                      <strong style={{ color: "#8a4b00" }}>· outside</strong>
                    )}
                    <button
                      type="button"
                      style={{
                        background: "none",
                        border: "none",
                        cursor: "pointer",
                        padding: 0,
                        fontSize: 12,
                      }}
                      onClick={() =>
                        setDirectShareRecipients((r) =>
                          r.filter((e) => e !== email),
                        )
                      }
                    >
                      ✕
                    </button>
                  </span>
                ))}
              </div>
            )}
            <input
              type="text"
              value={directShareQuery}
              onChange={(e) => setDirectShareQuery(e.target.value)}
              placeholder="Search by name or email"
              style={{
                width: "100%",
                boxSizing: "border-box",
                padding: "6px 8px",
                marginBottom: 4,
              }}
            />
            {directShareResults.length > 0 && (
              <div
                style={{
                  border: "1px solid #edebe9",
                  borderRadius: 4,
                  marginBottom: 8,
                  maxHeight: 160,
                  overflowY: "auto",
                }}
              >
                {directShareResults.map((p) => (
                  <button
                    key={p.loginName}
                    type="button"
                    style={{
                      display: "block",
                      width: "100%",
                      textAlign: "left",
                      padding: "6px 8px",
                      border: "none",
                      background: "none",
                      cursor: "pointer",
                      fontSize: 12.5,
                    }}
                    onClick={() => {
                      if (
                        p.email &&
                        directShareRecipients.indexOf(p.email) === -1
                      ) {
                        setDirectShareRecipients((r) => [...r, p.email]);
                      }
                      setDirectShareQuery("");
                      setDirectShareResults([]);
                    }}
                  >
                    {p.displayName} — {p.email}
                  </button>
                ))}
              </div>
            )}
            {directShareError && (
              <p style={{ fontSize: 12.5, color: "#a4262c", margin: "0 0 8px" }}>
                {directShareError}
              </p>
            )}
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button
                style={s.ghost}
                disabled={directShareBusy}
                onClick={() => setDirectShareOpen(false)}
              >
                Cancel
              </button>
              <button
                style={s.primary}
                disabled={directShareBusy || directShareRecipients.length === 0}
                onClick={() => {
                  if (
                    directFileView === undefined ||
                    directFileView.state !== "ready"
                  ) {
                    return;
                  }
                  const view = directFileView;
                  setDirectShareBusy(true);
                  setDirectShareError(undefined);
                  performDirectShare(
                    view.fileRef,
                    view.name,
                    directShareRecipients,
                  )
                    .then((err) => {
                      if (err) {
                        setDirectShareError(err);
                        return;
                      }
                      setDirectShareDone(true);
                    })
                    .catch((e) => setDirectShareError((e as Error).message))
                    .finally(() => setDirectShareBusy(false));
                }}
              >
                {directShareBusy ? "Sharing…" : "Share"}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
```

- [ ] **Step 4: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors. If `s.chip` does not exist in this file's `s` object, check with
`grep -n "chip:" src/webparts/requests/components/Requests.tsx` — this project already uses a
`chip`-style object for the HC tag and status pills elsewhere in this same file (`s.hcTag`,
`s.chip` in the "Shared files" recipients rendering), so it should already exist; if the exact key
differs, use whatever chip-shaped style key this file already defines instead of inventing a new
one.

- [ ] **Step 5: Run the full test suite**

Run: `npx jest`
Expected: PASS, same total as before Task 1 plus the 5 new `directActionsFor` tests, no
regressions.

- [ ] **Step 6: Commit**

```bash
git add src/webparts/requests/components/Requests.tsx
git commit -m "feat: direct share — immediate SP.Web.ShareObject call, view-only, recipient chip entry"
```

---

### Task 9: Final verification pass

**Files:** none new — this task only runs checks.

- [ ] **Step 1: Full typecheck**

Run: `npx tsc --noEmit`
Expected: no errors.

- [ ] **Step 2: Full lint**

Run: `npx eslint src/shared/directFileActions.ts src/shared/pageAccessPolicy.ts src/webparts/requests/components/Requests.tsx`
Expected: only pre-existing warnings on `Requests.tsx` (its documented `max-lines` warning, since
this file was already over the project's 2000-line ceiling before this change) — no NEW warning
categories, no errors.

- [ ] **Step 3: Full test suite**

Run: `npx jest`
Expected: PASS — every existing suite plus the new `directFileActions.test.ts`.

- [ ] **Step 4: Full production build**

Run: `npm run build`
Expected: exits 0. Per this project's own standing lesson, never trust `heft package-solution`
alone — this script runs the full `heft test --clean --production && heft package-solution
--production` pipeline that actually regenerates the hashed bundles.

- [ ] **Step 5: Spot-check the shipped bundle**

```bash
cd sharepoint/solution
unzip -l sd-gatrie.sppkg | grep requests-web-part
```

Note the hash in the filename, then confirm it differs from whatever was shipped before this
change (if unsure, this is a sufficient check on its own: a freshly-built hash for a file this
plan modified is itself evidence the bundle is not stale).

- [ ] **Step 6: Final commit if anything is outstanding**

```bash
git status
```

If clean, nothing further to commit — Tasks 1-8 already covered every change. If a manual
`--fix`-style lint pass or formatting adjustment was needed during verification, commit it:

```bash
git add -A
git commit -m "chore: final lint/format pass for the direct file view feature"
```

---

## What is deliberately NOT in this plan

Per the spec's own Non-goals and open-risk sections:

- **The column formatting JSON is not part of this plan.** It is a manual, per-library, per-site
  SharePoint configuration step (not shippable via the `.sppkg`), and it depends on confirming
  whether `[$UniqueId]` is available as a column-formatting token on this tenant — something that
  can only be checked once this code is deployed and a real `?file=<UniqueId>` link can be tested
  by hand. Deliver it as a follow-up once Task 9 is verified live, per the spec's own instruction.
- **No "Request…" fallback for viewers without direct rights**, and **no extraction of
  `performShare`/`writeApprovedDeletionRequest` into a shared module** — both explicitly declined
  in the spec's Non-goals.

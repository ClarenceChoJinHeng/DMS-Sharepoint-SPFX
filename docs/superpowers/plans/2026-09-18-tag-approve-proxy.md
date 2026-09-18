# Tag/Approve By Proxy — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** PIC and Approver lose direct Edit/Approve rights on the document libraries, so they can
no longer rename or retag files through the native SharePoint library view. Tagging on upload and
the approve/reject decision are instead written as records to two SharePoint lists PIC/Approver
already have plain-list write rights to, and two new Power Automate flows perform the actual
privileged write as `crs@sdguthrie.com`.

**Architecture:** Two new "pending write" lists, one extended (`CRS Submissions`, already exists),
one new (`CRS Pending Decisions`). A browser writes to the list instead of the document item; a
flow reads the pending row and performs the real write. Same shape as the 2026-09-17
deletion-by-proxy work.

**Tech Stack:** TypeScript/React (SPFx), SharePoint REST, Power Automate (built by hand in the
designer — not something this repo can build or test directly).

**Spec:** `docs/superpowers/specs/2026-09-18-tag-approve-proxy-design.md`
**Reversion plan:** `docs/superpowers/specs/2026-09-18-tag-approve-proxy-reversion.md`

**Correction from the design doc, made while writing this plan:** the spec said flow-side
resolution would reuse `resolveStamped`/`findByStamp` from `shared/requests.ts`. Those are pure
TypeScript helpers that operate on results an SPFx component has already fetched — a Power
Automate flow cannot call them, and they exist specifically to handle a multi-day gap between a
deletion request being raised and being decided. Tag/approve records here are written in the same
browser session as the check that produced them, seconds apart, so there is no such gap. Both new
flows resolve the target file with a plain `GetFileById(guid'ItemUniqueId')` — the same simpler
approach the already-shipped deletion flow (`CRS — Execute approved deletion`) actually uses, which
also has no stamp-fallback resolution. This is a simplification from the spec, not a deviation from
its intent.

**Phases:**
- **Phase 1** — schema + pure logic. Fully unit-tested (jest), no SPFx runtime dependency.
- **Phase 2** — SPFx write/read layer for both lists. Not independently jest-testable (imports
  `@microsoft/sp-http`), verified by `tsc --noEmit` plus later live testing.
- **Phase 3** — wire the tagging write path (`Form.tsx`, `BulkUpload.tsx`).
- **Phase 4** — wire the approve/reject write path (`ApprovalDocument.tsx`, `BulkApprovePanel.tsx`).
- **Phase 5** — self-approve sequencing.
- **Phase 6** — the two Power Automate flows and the permission-level cutover. This phase is a
  runbook, not TDD — Power Automate has no automated test runner this repo can invoke, and every
  other flow in this project has been built and verified the same way: exact field-by-field
  instructions, followed by hand, verified with a real test run. Do not invent fake "tests" for it.

---

## Phase 1 — Schema + pure logic

### Task 1: Extend `CRS Submissions`' pure schema with the tag payload fields

**Files:**
- Modify: `src/shared/submissionRecords.ts`
- Test: `src/shared/submissionRecords.test.ts`

- [ ] **Step 1: Write the failing tests**

Add to `src/shared/submissionRecords.test.ts` (find the existing `describe("RECORD_COLUMNS"` block
and the `buildRecordPayload`/`parseRecordRow` describe blocks, and add alongside them):

```typescript
describe("tag payload columns", () => {
  it("RECORD_COLUMNS includes TagPayload, TagStatus, TagError", () => {
    const names = RECORD_COLUMNS.map((c) => c.name);
    expect(names).toContain("TagPayload");
    expect(names).toContain("TagStatus");
    expect(names).toContain("TagError");
  });

  it("TagPayload and TagError are type 3 (Note) — a JSON blob and a free-text reason must not be capped like a Text column", () => {
    const tagPayload = RECORD_COLUMNS.find((c) => c.name === "TagPayload");
    const tagError = RECORD_COLUMNS.find((c) => c.name === "TagError");
    expect(tagPayload?.type).toBe(3);
    expect(tagError?.type).toBe(3);
  });

  it("TagStatus is type 2 (Text) — matches every other status field in this list", () => {
    const tagStatus = RECORD_COLUMNS.find((c) => c.name === "TagStatus");
    expect(tagStatus?.type).toBe(2);
  });
});

describe("buildRecordPayload — tag fields", () => {
  it("includes an untagged row's tagPayload and leaves tagStatus/tagError absent", () => {
    const payload = buildRecordPayload({
      submissionRef: "SUB-20260918-TEST",
      batchRef: "BAT-20260918-TEST",
      fileId: "SFI-20260918-TEST",
      fileName: "test.pdf",
      itemPath: "/sites/CRS/ApprovalDocument/GHO/GF/TAX/test.pdf",
      libraryTitle: "Approval for Document",
      uploadedBy: "pic@example.com",
      metadata: {},
      source: "Form",
      tagPayload: JSON.stringify([{ FieldName: "Year", FieldValue: "2026|guid" }]),
    });
    expect(payload.TagPayload).toBe(
      JSON.stringify([{ FieldName: "Year", FieldValue: "2026|guid" }]),
    );
    expect(payload.TagStatus).toBeUndefined();
  });

  it("omits tagPayload entirely when the caller supplies none — a row must never claim an empty payload was intentional", () => {
    const payload = buildRecordPayload({
      submissionRef: "SUB-20260918-TEST",
      batchRef: "BAT-20260918-TEST",
      fileId: "SFI-20260918-TEST",
      fileName: "test.pdf",
      itemPath: "/sites/CRS/ApprovalDocument/GHO/GF/TAX/test.pdf",
      libraryTitle: "Approval for Document",
      uploadedBy: "pic@example.com",
      metadata: {},
      source: "Form",
    });
    expect(payload.TagPayload).toBeUndefined();
  });
});

describe("parseRecordRow — tag fields", () => {
  it("reads tagStatus and tagError back", () => {
    const record = parseRecordRow({
      Id: 1,
      SubmissionRef: "SUB-20260918-TEST",
      BatchRef: "BAT-20260918-TEST",
      SubmissionFileId: "SFI-20260918-TEST",
      FileName: "test.pdf",
      ItemPath: "/sites/CRS/ApprovalDocument/GHO/GF/TAX/test.pdf",
      LibraryTitle: "Approval for Document",
      UploadedBy: "pic@example.com",
      Source: "Form",
      TagStatus: "Failed",
      TagError: "HasException: DocumentDate is not a valid date",
    });
    expect(record.tagStatus).toBe("Failed");
    expect(record.tagError).toBe(
      "HasException: DocumentDate is not a valid date",
    );
  });

  it("leaves tagStatus undefined, never a guessed value, when the row carries none", () => {
    const record = parseRecordRow({
      Id: 1,
      SubmissionRef: "SUB-20260918-TEST",
      BatchRef: "BAT-20260918-TEST",
      SubmissionFileId: "SFI-20260918-TEST",
      FileName: "test.pdf",
      ItemPath: "/sites/CRS/ApprovalDocument/GHO/GF/TAX/test.pdf",
      LibraryTitle: "Approval for Document",
      UploadedBy: "pic@example.com",
      Source: "Form",
    });
    expect(record.tagStatus).toBeUndefined();
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx heft test --clean 2>&1 | grep -A5 "tag payload columns"`
Expected: FAIL — `TagPayload`/`TagStatus`/`TagError` do not exist on `RECORD_COLUMNS`; `tagPayload`
is not a valid property on the record builder's input type; `tagStatus`/`tagError` do not exist on
the parsed result.

- [ ] **Step 3: Add the columns to `RECORD_COLUMNS`**

In `src/shared/submissionRecords.ts`, find the `RECORD_COLUMNS` array (ends at the `WithdrawnBy`
entry, just before the `RECORD_JOIN_COLUMN` export). Add after `WithdrawnBy`:

```typescript
  /* -- Tag-by-proxy (2026-09-18) --------------------------------------------
     Client: PIC and Approver lose direct Edit rights on the document libraries, so they can no
     longer retag a file by hand in the library view. `TagPayload` is what `CRS - Apply pending
     tags` (Power Automate, running as crs@sdguthrie.com) reads to perform the write the uploader's
     own browser can no longer make directly.

     NOT THE SAME FIELD AS `MetadataSnapshot`, DELIBERATELY. `MetadataSnapshot` stores LABELS
     ONLY, is missing Document Type/Year/Keyword/Business Segment, and exists purely for DISPLAY --
     it is read by My Submissions' detail panel for a deleted/archived/replaced record, and its
     correctness bar is "good enough for a person to read." `TagPayload` is the EXACT write
     payload -- the same `{FieldName, FieldValue}` array `Form.tsx`/`BulkUpload.tsx` used to POST
     to `validateUpdateListItem` directly, taxonomy fields already in `Label|GUID` form -- and its
     correctness bar is "a flow can replay this and get the same result the uploader chose."
     Reusing `MetadataSnapshot` for this would entangle a display concern with a write concern that
     have different tolerances for being slightly wrong. See the 2026-09-18 design doc. */
  { name: "TagPayload", type: 3 },
  /** blank/absent means not yet processed. "Tagged" means applied. "Failed" means see TagError.
   * Same three-state shape as every other status field in this codebase: blank is not a failure. */
  { name: "TagStatus", type: 2 },
  { name: "TagError", type: 3 },
```

- [ ] **Step 4: Extend `SubmissionRecord`**

Find the `SubmissionRecord` interface, add after `withdrawnBy`:

```typescript
  /** The exact `{FieldName, FieldValue}[]` payload for `CRS - Apply pending tags` to replay,
   * JSON-encoded. Absent on a row this flow has not been asked to act on. */
  tagPayload?: string;
  /** blank/undefined means not yet processed by the flow. "Tagged" means applied. "Failed" means
   * see tagError. */
  tagStatus?: string;
  tagError?: string;
```

- [ ] **Step 5: Extend `buildRecordPayload`**

Find `buildRecordPayload` (builds the write body from a `SubmissionRecord`-shaped input). Find
where it sets fields conditionally (matching the existing pattern for `replacedAt`/`replacedBy` —
present-only, never written as an empty string). Add, following the same conditional-inclusion
style already used for the optional fields in that function:

```typescript
  if (input.tagPayload !== undefined) payload.TagPayload = input.tagPayload;
  if (input.tagStatus !== undefined) payload.TagStatus = input.tagStatus;
  if (input.tagError !== undefined) payload.TagError = input.tagError;
```

- [ ] **Step 6: Extend `parseRecordRow`**

Find `parseRecordRow` (reads a raw list-item row back into a `SubmissionRecord`). Add, following
the same optional-field-reading pattern already used for `ReplacedBy`/`WithdrawnBy` a few lines
above (reuse whatever helper that pattern already uses — don't introduce a second way to read an
optional string field):

```typescript
    tagPayload: str("TagPayload") || undefined,
    tagStatus: str("TagStatus") || undefined,
    tagError: str("TagError") || undefined,
```

- [ ] **Step 7: Run the tests to verify they pass**

Run: `npx heft test --clean 2>&1 | tail -40`
Expected: PASS, full suite green, no new failures.

- [ ] **Step 8: Commit**

```bash
git add src/shared/submissionRecords.ts src/shared/submissionRecords.test.ts
git commit -m "feat: CRS Submissions carries the tag-by-proxy payload

TagPayload/TagStatus/TagError, separate from the existing label-only
MetadataSnapshot which stays exactly as it is for display purposes.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 2: New pure module for pending approve/reject decisions

**Files:**
- Create: `src/shared/pendingDecisions.ts`
- Test: `src/shared/pendingDecisions.test.ts`

- [ ] **Step 1: Write the failing tests**

Create `src/shared/pendingDecisions.test.ts`:

```typescript
import {
  PENDING_DECISION_COLUMNS,
  DECISION_JOIN_COLUMN,
  buildDecisionPayload,
  parseDecisionRow,
  PendingDecision,
} from "./pendingDecisions";

describe("PENDING_DECISION_COLUMNS", () => {
  it("names every column CRS Pending Decisions needs", () => {
    const names = PENDING_DECISION_COLUMNS.map((c) => c.name);
    expect(names).toEqual([
      "ItemUniqueId",
      "Decision",
      "DecidedBy",
      "DecidedAt",
      "RejectionComment",
      "IsSelfApprove",
      "Status",
      "StatusError",
    ]);
  });

  it("Decision and Status are TEXT, never Choice", () => {
    const decision = PENDING_DECISION_COLUMNS.find((c) => c.name === "Decision");
    const status = PENDING_DECISION_COLUMNS.find((c) => c.name === "Status");
    expect(decision?.type).toBe(2);
    expect(status?.type).toBe(2);
  });

  it("IsSelfApprove is type 8 (Boolean)", () => {
    const col = PENDING_DECISION_COLUMNS.find((c) => c.name === "IsSelfApprove");
    expect(col?.type).toBe(8);
  });
});

describe("buildDecisionPayload", () => {
  const base: Omit<PendingDecision, "id" | "status" | "statusError"> = {
    itemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
    decision: "Approved",
    decidedBy: "approver@example.com",
    decidedAt: new Date("2026-09-18T02:30:00Z"),
    isSelfApprove: false,
  };

  it("writes the decision and actor, omits RejectionComment for an approval", () => {
    const payload = buildDecisionPayload(base);
    expect(payload.Decision).toBe("Approved");
    expect(payload.DecidedBy).toBe("approver@example.com");
    expect(payload.ItemUniqueId).toBe("f7d2d18c-1234-4a4a-9a9a-abcdefabcdef");
    expect(payload.RejectionComment).toBeUndefined();
  });

  it("includes RejectionComment for a rejection", () => {
    const payload = buildDecisionPayload({
      ...base,
      decision: "Rejected",
      rejectionComment: "Wrong vendor named on the document",
    });
    expect(payload.Decision).toBe("Rejected");
    expect(payload.RejectionComment).toBe("Wrong vendor named on the document");
  });

  it("marks IsSelfApprove true for a self-approve", () => {
    const payload = buildDecisionPayload({ ...base, isSelfApprove: true });
    expect(payload.IsSelfApprove).toBe(true);
  });

  it("never sets Status on the initial write — the row starts Pending by omission, the flow is what marks it Applied/Failed", () => {
    const payload = buildDecisionPayload(base);
    expect(payload.Status).toBeUndefined();
  });
});

describe("parseDecisionRow", () => {
  it("reads an applied decision back", () => {
    const row = parseDecisionRow({
      Id: 42,
      ItemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
      Decision: "Approved",
      DecidedBy: "approver@example.com",
      DecidedAt: "2026-09-18T02:30:00Z",
      IsSelfApprove: false,
      Status: "Applied",
    });
    expect(row.id).toBe(42);
    expect(row.decision).toBe("Approved");
    expect(row.status).toBe("Applied");
    expect(row.isSelfApprove).toBe(false);
  });

  it("leaves status undefined, never a guessed value, on a row with none", () => {
    const row = parseDecisionRow({
      Id: 42,
      ItemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
      Decision: "Approved",
      DecidedBy: "approver@example.com",
      DecidedAt: "2026-09-18T02:30:00Z",
      IsSelfApprove: false,
    });
    expect(row.status).toBeUndefined();
  });

  it("defaults isSelfApprove to false, never undefined, when the column is missing from an older row", () => {
    const row = parseDecisionRow({
      Id: 42,
      ItemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
      Decision: "Approved",
      DecidedBy: "approver@example.com",
      DecidedAt: "2026-09-18T02:30:00Z",
    });
    expect(row.isSelfApprove).toBe(false);
  });
});

describe("DECISION_JOIN_COLUMN", () => {
  it("is ItemUniqueId — the flow addresses the document by this, never by name or path", () => {
    expect(DECISION_JOIN_COLUMN).toBe("ItemUniqueId");
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx heft test --clean 2>&1 | grep -A5 "pendingDecisions"`
Expected: FAIL — `Cannot find module './pendingDecisions'`.

- [ ] **Step 3: Write the module**

Create `src/shared/pendingDecisions.ts`:

```typescript
/**
 * Pending approve/reject decisions — the write, deferred.
 *
 * Spec: docs/superpowers/specs/2026-09-18-tag-approve-proxy-design.md
 *
 * Client: PIC and Approver lose direct Edit/Approve rights on the document libraries, so an
 * approver's own click can no longer flip `OData__ModerationStatus` directly — that write moves to
 * `CRS — Apply pending decisions` (Power Automate, running as crs@sdguthrie.com). This module is
 * the shape of the row that drives it: everything the flow needs to replay the decision, nothing
 * more.
 *
 * A NEW, purpose-built list, not an extension of `CRS Requests`. That list's whole design —
 * `RequestType: "Deletion" | "Share"`, a requester and a SEPARATE decider, the `ViewerScope`
 * visibility model — is built around *someone else decides on your behalf*. An approve/reject
 * decision (including a self-approve, where the same person both uploads and decides) is a
 * different shape: one action, not a request-then-decision pair. Bolting a third `RequestType`
 * onto `CRS Requests`' already heavily-worked machinery risks breaking what currently works, for
 * no benefit over a small new list.
 *
 * Pure and SPFx-free, same split as `submissionRecords.ts` / `spSubmissionRecords.ts`.
 */

export const PENDING_DECISION_COLUMNS: Array<{ name: string; type: number }> = [
  /** The document being decided on. The ONLY identifier — resolved by `GetFileById(guid'...')`,
   * web-scoped, works across every library regardless of normal/HC. Written fresh, moments before
   * this row is written, in the same browser session as the check that produced it — unlike a
   * `CRS Requests` row, there is no multi-day gap here, so no stamp-based fallback is needed. */
  { name: "ItemUniqueId", type: 2 },
  /** TEXT, never Choice — a value absent from a Choice column's `Choices` fails the whole write. */
  { name: "Decision", type: 2 },
  { name: "DecidedBy", type: 2 },
  { name: "DecidedAt", type: 4 },
  /** Only present for a Rejected decision. */
  { name: "RejectionComment", type: 3 },
  /** True when this came from `autoApproveOwnUpload`, never a human clicking Approve on someone
   * else's file. Exists so the applying flow — and any future reporting — can tell the two apart
   * without guessing from `DecidedBy` matching the file's own author. */
  { name: "IsSelfApprove", type: 8 },
  /** blank/absent means not yet processed. "Applied" means the flow performed the write. "Failed"
   * means see StatusError. Same three-state shape as `TagStatus` on `CRS Submissions`. */
  { name: "Status", type: 2 },
  { name: "StatusError", type: 3 },
];

/** The column the flow addresses the document by. Never name, never path — both can change. */
export const DECISION_JOIN_COLUMN = "ItemUniqueId";

export type Decision = "Approved" | "Rejected";

export interface PendingDecision {
  id: number;
  itemUniqueId: string;
  decision: Decision;
  decidedBy: string;
  decidedAt: Date;
  rejectionComment?: string;
  isSelfApprove: boolean;
  status?: string;
  statusError?: string;
}

/** Builds the write body for a NEW decision row. Never sets `Status` — the row starts Pending by
 * omission; only the flow that applies the decision ever writes `Applied`/`Failed`. */
export function buildDecisionPayload(
  input: Omit<PendingDecision, "id" | "status" | "statusError">,
): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    ItemUniqueId: input.itemUniqueId,
    Decision: input.decision,
    DecidedBy: input.decidedBy,
    DecidedAt: input.decidedAt.toISOString(),
    IsSelfApprove: input.isSelfApprove,
  };
  if (input.rejectionComment !== undefined) {
    payload.RejectionComment = input.rejectionComment;
  }
  return payload;
}

/** Reads a raw list-item row back into a `PendingDecision`. `isSelfApprove` defaults to `false`
 * rather than `undefined` when the column is missing (an older row, or a read that omitted it) —
 * an ordinary human approval must never be misread as a self-approve for want of a value. */
export function parseDecisionRow(row: Record<string, unknown>): PendingDecision {
  const str = (k: string): string | undefined => {
    const v = row[k];
    return typeof v === "string" && v.length > 0 ? v : undefined;
  };
  return {
    id: Number(row.Id ?? 0),
    itemUniqueId: str("ItemUniqueId") ?? "",
    decision: (str("Decision") as Decision) ?? "Approved",
    decidedBy: str("DecidedBy") ?? "",
    decidedAt: new Date(String(row.DecidedAt ?? "")),
    rejectionComment: str("RejectionComment"),
    isSelfApprove: row.IsSelfApprove === true,
    status: str("Status"),
    statusError: str("StatusError"),
  };
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx heft test --clean 2>&1 | tail -40`
Expected: PASS, full suite green.

- [ ] **Step 5: Commit**

```bash
git add src/shared/pendingDecisions.ts src/shared/pendingDecisions.test.ts
git commit -m "feat: pure module for pending approve/reject decisions

New CRS Pending Decisions shape - a row an approver's browser writes
instead of flipping moderation status directly. Deliberately separate
from CRS Requests, which is built around a different shape (someone
else decides on your behalf).

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

## Phase 2 — SPFx write/read layer

**Not independently jest-testable** (imports `@microsoft/sp-http`, which the test environment
cannot resolve — same constraint `spSubmissionRecords.ts` and `spAuditLog.ts` already carry).
Verified by `tsc --noEmit` in each task, then by the live end-to-end test in Phase 6.

### Task 3: Extend `spSubmissionRecords.ts` to write the tag payload

**Files:**
- Modify: `src/shared/spSubmissionRecords.ts`

- [ ] **Step 1: Read the existing `writeSubmissionRecord` function in full**

Read `src/shared/spSubmissionRecords.ts` end to end before touching it — confirm the exact shape
of `RECORD_READ_SELECT`/`RECORD_READ_SELECT_NO_WITHDRAWAL`/`RECORD_READ_SELECT_NO_ARCHIVE`/
`RECORD_READ_SELECT_LEGACY` (the four-rung read ladder this list already has), and where
`buildRecordPayload` is called. The new `TagPayload`/`TagStatus`/`TagError` fields need to become a
fifth rung on that same ladder (one unknown field name fails the whole `$select` — gotcha #11 in
this project's own history), not folded into an existing rung.

- [ ] **Step 2: Add the new read rung**

In `src/shared/submissionRecords.ts` (Phase 1, Task 1 already added the columns there), add a new
exported constant alongside the existing `RECORD_READ_SELECT*` exports, following the exact same
"subtract the newest optional columns" pattern already used to build
`RECORD_READ_SELECT_NO_WITHDRAWAL` from `RECORD_READ_SELECT`:

```typescript
/** The read ladder's newest rung, minus TagPayload/TagStatus/TagError — for a site that has not
 * yet had reconciliation asserted since this feature shipped. Derived by SUBTRACTION from the full
 * list, never written out again, so the two cannot drift. */
export const RECORD_READ_SELECT_NO_TAGGING = RECORD_READ_SELECT_NO_WITHDRAWAL.filter(
  (f) => f !== "TagPayload" && f !== "TagStatus" && f !== "TagError",
);
```

(Confirm the exact field name casing matches whatever `RECORD_READ_SELECT` actually contains —
read the file first, per Step 1, rather than guessing the array's exact contents.)

In `spSubmissionRecords.ts`, extend the read function's ladder (wherever it currently retries
`RECORD_READ_SELECT` -> `RECORD_READ_SELECT_NO_WITHDRAWAL` -> `RECORD_READ_SELECT_NO_ARCHIVE` ->
`RECORD_READ_SELECT_LEGACY` on successive 400s) to try `RECORD_READ_SELECT` ->
`RECORD_READ_SELECT_NO_TAGGING` -> `RECORD_READ_SELECT_NO_WITHDRAWAL` ->
`RECORD_READ_SELECT_NO_ARCHIVE` -> `RECORD_READ_SELECT_LEGACY` — the new rung slots in as the
second one tried, since it's the newest optional addition.

- [ ] **Step 3: Confirm `writeSubmissionRecord` already accepts the new fields**

It should, automatically — `writeSubmissionRecord`'s `record` parameter is typed as
`Omit<SubmissionRecord, "itemId">`, and Phase 1 Task 1 already added `tagPayload`/`tagStatus`/
`tagError` as optional fields on `SubmissionRecord`, and `buildRecordPayload` (which
`writeSubmissionRecord` calls internally) already writes them when present. No code change should
be needed here — this step is a verification, not an edit. Read the function body to confirm it
calls `buildRecordPayload(record)` (or equivalent) rather than manually listing fields; if it
manually lists fields anywhere, add `tagPayload`/`tagStatus`/`tagError` to that list too.

- [ ] **Step 4: Type-check**

Run: `npx tsc --noEmit 2>&1 | tail -40`
Expected: no new errors.

- [ ] **Step 5: Commit**

```bash
git add src/shared/spSubmissionRecords.ts src/shared/submissionRecords.ts
git commit -m "feat: SPFx read ladder covers the new tag payload columns

A site that has not yet had reconciliation asserted since this feature
shipped must not lose its whole submissions read to gotcha #11 (one
unknown field name fails the whole \$select) over three new optional
columns.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 4: New SPFx module for `CRS Pending Decisions` — provisioning and write

**Files:**
- Create: `src/shared/spPendingDecisions.ts`

- [ ] **Step 1: Read `spSubmissionRecords.ts` and `naming.ts`'s `LIST_SUFFIX`/`PRIMED_SUFFIXES` in full**

This is the pattern to follow exactly. In particular: `LIST_SUFFIX` in `naming.ts` needs a new
entry (`pendingDecisions: "Pending Decisions"` or similar, matching the existing suffix naming
convention — read the existing entries before choosing the exact string), and it must be added to
`PRIMED_SUFFIXES` in the same file. `LIST_SUFFIX.requests` was left out of `PRIMED_SUFFIXES` for
weeks when that list was first built, and the failure was total and silent — an unprimed suffix
answers the legacy `DMS <suffix>` title forever, which 404s on a renamed site. There is an existing
test pinning that every `LIST_SUFFIX` entry appears in `PRIMED_SUFFIXES`; if this task fails that
test, that is the safety net working, not a bug to work around.

- [ ] **Step 2: Add the `LIST_SUFFIX` entry**

In `src/shared/naming.ts`, add the new suffix to `LIST_SUFFIX` and to `PRIMED_SUFFIXES`. Read the
existing entries first to match the exact naming/casing convention already used for `requests` and
`submissions`.

- [ ] **Step 3: Run the existing suffix-parity test**

Run: `npx heft test --clean 2>&1 | grep -A10 "PRIMED_SUFFIXES"`
Expected: PASS — confirms the new suffix was added to both places correctly.

- [ ] **Step 4: Write `spPendingDecisions.ts`**

Create `src/shared/spPendingDecisions.ts`, mirroring `spSubmissionRecords.ts`'s structure exactly
(the same `WRITE_HEADERS`/`GET_HEADERS` constants — copy them verbatim, they encode a real gotcha
about `odata-version` header injection that cost this project real time twice already — and the
same `listBase()` pattern):

```typescript
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import {
  PendingDecision,
  buildDecisionPayload,
  parseDecisionRow,
  PENDING_DECISION_COLUMNS,
} from "./pendingDecisions";
import { cachedListTitle, LIST_SUFFIX } from "./naming";

/**
 * `CRS Pending Decisions` — the SPFx half. `pendingDecisions.ts` holds everything testable, same
 * split as `submissionRecords.ts` / `spSubmissionRecords.ts` and `auditLog.ts` / `spAuditLog.ts`.
 */

const WRITE_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
  "odata-version": "",
};

const GET_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Cache-Control": "no-cache",
  Pragma: "no-cache",
};

const listBase = (siteUrl: string): string =>
  `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.pendingDecisions))}')`;

/**
 * Creates `CRS Pending Decisions` if it does not already exist. Idempotent — a 400 on the creation
 * call for "a list with this title already exists" is treated as success, matching the pattern
 * `CRS Requests`' own `provision()` function already uses.
 */
export async function provisionPendingDecisionsList(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<void> {
  const existsRes = await sp.get(
    `${listBase(siteUrl)}?$select=Id`,
    SPHttpClient.configurations.v1,
    { headers: GET_HEADERS },
  );
  if (existsRes.ok) return;

  await sp.post(
    `${siteUrl}/_api/web/lists`,
    SPHttpClient.configurations.v1,
    {
      headers: WRITE_HEADERS,
      body: JSON.stringify({
        Title: cachedListTitle(LIST_SUFFIX.pendingDecisions),
        BaseTemplate: 100,
      }),
    },
  );

  for (const col of PENDING_DECISION_COLUMNS) {
    await sp.post(
      `${listBase(siteUrl)}/fields`,
      SPHttpClient.configurations.v1,
      {
        headers: WRITE_HEADERS,
        body: JSON.stringify({
          "@odata.type": "SP.Field",
          Title: col.name,
          FieldTypeKind:
            col.type === 8 ? 8 : col.type === 4 ? 4 : col.type === 3 ? 3 : 2,
        }),
      },
    );
  }
}

/**
 * Write a NEW decision. **Never throws.** The document has already passed every client-side check
 * by the time this runs (destination-folder guard, clash checks, eligibility probe — all reads,
 * all unaffected by this feature) — a failure to write the DECISION record must not be silently
 * swallowed the way a failed audit write is elsewhere in this codebase, because unlike an audit
 * row, this record IS the mechanism. If this write fails, the approve/reject action has not
 * actually happened, and the caller must know that and tell the user, rather than reporting
 * success on a click that did nothing.
 */
export async function writePendingDecision(
  sp: SPHttpClient,
  siteUrl: string,
  decision: Omit<PendingDecision, "id" | "status" | "statusError">,
): Promise<{ ok: boolean; status?: number; body?: string }> {
  const res: SPHttpClientResponse = await sp.post(
    `${listBase(siteUrl)}/items`,
    SPHttpClient.configurations.v1,
    {
      headers: WRITE_HEADERS,
      body: JSON.stringify(buildDecisionPayload(decision)),
    },
  );
  if (res.ok) return { ok: true };
  const body = await res.text().catch(() => "");
  console.error(
    `writePendingDecision failed: HTTP ${res.status} — ${body}`,
  );
  return { ok: false, status: res.status, body };
}

/**
 * Read the current decision row for a document, if any — used by the client to poll whether its
 * own decision has been applied yet (so the UI can refresh once `Status` flips), same shape as the
 * live-refresh polling already used on My Submissions and the Requests page.
 */
export async function readPendingDecision(
  sp: SPHttpClient,
  siteUrl: string,
  itemUniqueId: string,
): Promise<PendingDecision | undefined> {
  const res = await sp.get(
    `${listBase(siteUrl)}/items?$filter=ItemUniqueId eq '${encodeURIComponent(
      itemUniqueId,
    )}'&$orderby=Id desc&$top=1`,
    SPHttpClient.configurations.v1,
    { headers: GET_HEADERS },
  );
  if (!res.ok) return undefined;
  const json = await res.json();
  const rows = json?.value ?? [];
  return rows.length > 0 ? parseDecisionRow(rows[0]) : undefined;
}
```

**Note on the `FieldTypeKind` mapping above** — write it, but confirm the exact numeric constants
against a real, already-working provisioning function in this codebase (e.g. `Requests.tsx`'s own
`provision()`/`ensureColumns` or `optionalColumns.ts`) before trusting the values shown here; they
are written from general SharePoint REST knowledge, not verified against this specific project's
own working reference, and this task's own instructions require reading that reference first.

- [ ] **Step 5: Type-check**

Run: `npx tsc --noEmit 2>&1 | tail -40`
Expected: no new errors.

- [ ] **Step 6: Commit**

```bash
git add src/shared/spPendingDecisions.ts src/shared/naming.ts
git commit -m "feat: SPFx layer for CRS Pending Decisions - provision, write, read

Mirrors spSubmissionRecords.ts's structure. New LIST_SUFFIX entry
added to PRIMED_SUFFIXES in the same commit, per the standing lesson
from CRS Requests being left out of it for weeks.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

## Phase 3 — Wire the tagging write path

### Task 5: `Form.tsx` — stop tagging directly, write `TagPayload` instead

**Files:**
- Modify: `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Read the current tagging sequence in full**

Read `Form.tsx` from where `formValues` is built through the `validateUpdateListItem` POST through
the `writeSubmissionRecord` call. Confirm the exact current order: upload -> tag (direct write) ->
submission-record write (currently a pure "what happened" record, written AFTER tagging succeeds).

- [ ] **Step 2: Move the submission-record write earlier, and stop the direct tag POST**

Replace the `validateUpdateListItem` POST with: build the same `formValues` array exactly as today
(nothing about the array itself changes — same tier values, same `toTaxValue`-encoded taxonomy
fields), JSON-encode it, and pass it as `tagPayload` to `writeSubmissionRecord`. The
submission-record write that currently happens AFTER tagging becomes the ONLY write this function
makes for tagging — remove the separate `validateUpdateListItem` call entirely.

The exact code:

```typescript
      // formValues is built exactly as before this change — see the surrounding code, unchanged.
      // The write below is what used to be `validateUpdateListItem`.
      const tagged = await writeSubmissionRecord(context.spHttpClient, siteUrl, {
        submissionRef: refs.submissionId,
        batchRef: refs.batchId,
        fileId: refs.fileId,
        uniqueId:
          typeof item.UniqueId === "string" ? item.UniqueId : undefined,
        fileName: finalName,
        itemPath: uploadedServerRelativeUrl,
        libraryTitle: libraryTitleForTagging,
        uploadedBy: (context.pageContext.user.email ?? "").toLowerCase(),
        uploadedAt: new Date(),
        metadata: snapshot,
        source: "Form",
        tagPayload: JSON.stringify(formValues),
      });
      if (!tagged) {
        return {
          fileId: sf.id,
          ok: false,
          error:
            "Uploaded, but could not record it for tagging. An administrator will need to tag this file by hand.",
        };
      }
```

**This changes the failure semantics.** Today, a failed `validateUpdateListItem` call reports
"Uploaded, but tagging metadata failed" — the file exists, untagged, immediately, and the uploader
sees the failure right away. After this change, a failed `writeSubmissionRecord` call means the
file will never be tagged (no flow will ever see it, since no row was written at all) — this is a
worse failure than today's, because there is no automatic retry path. The message above reflects
that ("An administrator will need to tag this file by hand") rather than reusing today's wording,
which would understate it.

- [ ] **Step 3: Remove the now-unused `validateUpdateListItem` call and its response handling**

Delete the code block that built the `validateUpdateListItem` request and parsed its response
(`HasException` checking, etc.) — it's fully superseded by Step 2. Check for any other reference to
that response elsewhere in the function (e.g. a variable holding per-field error details used in a
later message) and remove or adapt those too.

- [ ] **Step 4: Confirm the snapshot's own construction — `MetadataSnapshot` is unaffected**

Confirm the `snapshot` object (used for `MetadataSnapshot`) is left completely untouched by this
change — it is a separate, display-only concern (per Phase 1 Task 1's own reasoning) and must keep
being built exactly as it is today, independent of `tagPayload`.

- [ ] **Step 5: Run the full test suite**

Run: `npx heft test --clean 2>&1 | tail -40`
Expected: PASS. This is a component change with no dedicated unit tests of its own (this project
has none for its React components — pure logic is tested, components are verified live), so the
bar here is "nothing else regressed," not new component-level tests.

- [ ] **Step 6: Type-check**

Run: `npx tsc --noEmit 2>&1 | tail -40`
Expected: no new errors — in particular, confirm no leftover reference to the removed
`validateUpdateListItem` response variable.

- [ ] **Step 7: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: Form.tsx writes TagPayload instead of tagging directly

The uploader's own browser no longer POSTs to validateUpdateListItem.
CRS -- Apply pending tags (Power Automate, not yet built) reads the
payload and applies it as crs@sdguthrie.com. Physical upload (Files/Add)
is completely unchanged.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 6: `BulkUpload.tsx` — the identical change

**Files:**
- Modify: `src/webparts/bulkUpload/components/BulkUpload.tsx`

- [ ] **Step 1: Read the equivalent tagging sequence in `BulkUpload.tsx`**

Find its own `validateUpdateListItem` call and its own submission-record write (this file already
stamps `SubmissionId`/`BatchId`/`SubmissionFileId` — confirm it also calls
`writeSubmissionRecord` or an equivalent, and find where).

- [ ] **Step 2: Apply the identical change**

Same shape as Task 5, Steps 2–4: build `formValues` exactly as today, pass it as `tagPayload` to
the submission-record write, remove the direct `validateUpdateListItem` POST, update the failure
message the same way (reflecting that a failed record write now means "never tagged," not "tagged
later").

**`BulkUpload.tsx` writes straight to the approved side (`Documents`/`HC Documents`) with no
approval step** — confirm this task touches ONLY the tagging write, and does not touch anything
approval-related (there is nothing approval-related in this file to touch; noted here so the
person implementing this doesn't go looking for an approve/reject call that doesn't exist).

- [ ] **Step 3: Run the full test suite**

Run: `npx heft test --clean 2>&1 | tail -40`
Expected: PASS.

- [ ] **Step 4: Type-check**

Run: `npx tsc --noEmit 2>&1 | tail -40`
Expected: no new errors.

- [ ] **Step 5: Commit**

```bash
git add src/webparts/bulkUpload/components/BulkUpload.tsx
git commit -m "feat: BulkUpload.tsx writes TagPayload instead of tagging directly

Same change as Form.tsx. No approval-step change here -- this screen
writes straight to the approved side and always has.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 6b: `Form.tsx` — proxy the staging-replace overwrite too

**Not in the original plan.** Found during design review (see
`2026-09-18-tag-approve-proxy-design.md`'s "Replacing a pending draft" section, added the same
day): `Form.tsx` has a THIRD Edit-shaped write the tagging/decision proxying doesn't cover — the
2026-08-28 "yes, replace it" consent on a name clash, which calls `Files/Add(overwrite=true)`
directly against an *existing* item. Overwriting an existing item's content needs `EditListItems`,
not just `AddListItems` — uploading a brand-new file only needs Add. This 403s the moment
`Edit Items` comes off `CRS Upload`, and nothing else in this plan proxies it. `BulkUpload.tsx`
never overwrites (`overwrite` is hardcoded `false` there), so this is `Form.tsx`-only.

**The fix needs no new list or flow.** The 2026-09-17 proxy-deletion work
(`2026-09-17-proxy-deletion-via-power-automate-design.md`) already recycles a file by
`ItemUniqueId` as `crs@sdguthrie.com`, with no regard for authorship — exactly the shape this
needs, since "replace" today is already a unilateral, no-approval action taken by whoever clicked
Yes, whether or not the file is theirs. Re-express the replace as: delete the clashing draft via
that existing proxy path, then let the browser perform an ordinary, Add-only upload of the new
content under the same name once the old item is confirmed gone.

**⚠ THIS DEPENDS ON `CRS — Execute approved deletion` (the Power Automate flow, Task 13 of the
2026-09-17 plan) ACTUALLY EXISTING AND RUNNING.** As of this task, that flow is still not built —
same as every other place in this codebase that already depends on it (`Requests.tsx`,
`MySubmissions.tsx`). Until it exists, pressing "yes, replace it" will write the deletion request
correctly, then time out waiting for a recycle that never happens (see Step 4's timeout handling
below) — this is the same, already-accepted interim state the 2026-09-17 work shipped in, not a
new gap this task introduces.

**Files:**
- Modify: `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Read the current clash-consent-through-upload sequence in full**

Read from where `replaceStaging`/`overwritePending` are computed (search for
`const replaceStaging = replaceStagingIds.has(sf.id);`) through the `Files/Add` POST that reads
`overwrite=${overwritePending}`. Confirm the exact current shape: `overwritePending` is set true
only on a path the uploader has explicitly consented to (`decision.overwrite` from
`decideClash`), and immediately after it's set, an existing block already reads the clashing
item's `SubmissionFileId` into `displacedFileId` (used later, unrelated to this task, to mark the
OLD submission record "Cancelled" once the new upload completes — do not touch that mechanism).

- [ ] **Step 2: Also capture the clashing item's own `UniqueId`, alongside `SubmissionFileId`**

The existing read (search for `/ListItemAllFields?$select=SubmissionFileId`) only selects
`SubmissionFileId`. You need the item's SharePoint `UniqueId` too, to target the deletion-by-proxy
flow (`GetFileById(guid'<uniqueId>')/recycle()` is how it identifies a file — the same call
`Requests.tsx`'s `performDeletion`/`resolveDeletionTarget` already use). Extend the `$select` to
`SubmissionFileId,UniqueId` and capture both:

```typescript
          const priorRes: SPHttpClientResponse = await context.spHttpClient.get(
            `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')` +
              `/ListItemAllFields?$select=SubmissionFileId,UniqueId`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (priorRes.ok) {
            const prior = await priorRes.json();
            const raw =
              typeof prior?.SubmissionFileId === "string"
                ? prior.SubmissionFileId.trim()
                : "";
            if (raw.length > 0) displacedFileId = raw;
            const rawUniqueId =
              typeof prior?.UniqueId === "string" ? prior.UniqueId.trim() : "";
            if (rawUniqueId.length > 0) displacedItemUniqueId = rawUniqueId;
          }
```

Declare `displacedItemUniqueId` alongside the existing `displacedFileId` declaration (same scope,
same `let ... : string | undefined`).

- [ ] **Step 3: Write a proxy-delete-and-wait helper**

Add a new function in `Form.tsx`, near the other upload-time helpers in this same function scope
(it needs `context`, `siteUrl`, `listName` from the enclosing closure, same as the other helpers
here):

```typescript
    /**
     * Recycle the clashing draft via the 2026-09-17 proxy-deletion mechanism, then wait until it
     * is actually gone before returning — `Files/Add` with no `overwrite` flag needs the name to
     * be genuinely free, not merely requested-to-become-free.
     *
     * ⚠ WRITES A SELF-APPROVED `CRS Requests` ROW, THE SAME SHAPE `MySubmissions.tsx`'s
     * `writeApprovedDeletionRequest` ALREADY WRITES — `RequestType: "Deletion"`, `Status:
     * "Approved"` from the moment it's written, no human decision step, because that is already
     * the existing behaviour of "yes, replace it" today: nobody approves this, whoever clicked
     * Yes made the call unilaterally. `CRS — Execute approved deletion` picks up any row matching
     * `RequestType eq 'Deletion' and Status eq 'Approved'` regardless of who wrote it.
     *
     * Returns `true` once `GetFileById` on the old id starts 404ing (confirmed gone), `false` if
     * the deletion request itself could not be written, or if it never resolves within the poll
     * budget (most likely because `CRS — Execute approved deletion` is not yet built/running —
     * see this task's own note on that dependency).
     */
    const deleteClashingDraftByProxy = async (
      clashingUniqueId: string,
    ): Promise<boolean> => {
      const me = (context.pageContext.user.email ?? "").toLowerCase();
      const now = new Date().toISOString();
      try {
        const listTitle = await listName(LIST_SUFFIX.requests);
        const res = await context.spHttpClient.post(
          `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "odata-version": "",
            },
            body: JSON.stringify({
              Title: `Replace on upload — ${finalName}`.slice(0, 255),
              RequestType: "Deletion",
              Status: "Approved",
              ItemUniqueId: clashingUniqueId,
              ItemName: finalName,
              RequestedBy: me,
              RequestedAt: now,
              Reason: "",
              DecidedBy: me,
              DecidedAt: now,
              DecisionNote:
                "No approval needed — the uploader replaced this draft directly.",
            }),
          },
        );
        if (!res.ok) return false;
      } catch {
        return false;
      }
      // Poll budget: ~10 attempts, 3s apart — about 30s. Matches the "roughly a poll cycle"
      // framing already accepted elsewhere in this feature; this one is bounded and short because
      // it blocks the upload button, unlike the tag/decision flows' background ~1-minute poll.
      for (let attempt = 0; attempt < 10; attempt++) {
        await new Promise<void>((resolve) => setTimeout(resolve, 3000));
        try {
          const check = await context.spHttpClient.get(
            `${siteUrl}/_api/web/GetFileById(guid'${clashingUniqueId}')?$select=Exists`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (check.status === 404) return true;
        } catch {
          // Transient — keep polling within the budget rather than giving up on one failed check.
        }
      }
      return false;
    };
```

- [ ] **Step 4: Wire it in before the upload, and stop passing `overwrite=true` to `Files/Add`**

Where `overwritePending` currently feeds `Files/Add(...,overwrite=${overwritePending})`, insert the
proxy-delete-and-wait call BEFORE that POST, gated on `overwritePending` being true and
`displacedItemUniqueId` being available:

```typescript
    if (overwritePending) {
      if (!displacedItemUniqueId) {
        return {
          fileId: sf.id,
          ok: false,
          error:
            `Could not confirm which existing document to replace, so nothing was uploaded. ` +
            `Try again in a moment.`,
        };
      }
      const cleared = await deleteClashingDraftByProxy(displacedItemUniqueId);
      if (!cleared) {
        return {
          fileId: sf.id,
          ok: false,
          error:
            `"${finalName}" could not be cleared for replacement in time. Nothing was uploaded — ` +
            `try again in a moment, or contact an administrator if this keeps happening.`,
        };
      }
    }
```

Then change the `Files/Add` call itself to always pass `overwrite=false` (there is nothing left to
overwrite by the time this runs — the old item is confirmed gone, or this code path already
returned above):

```typescript
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=false)?$select=ServerRelativeUrl`,
```

`overwritePending` keeps its existing name and keeps gating the new block above — it no longer
feeds the `Files/Add` call's `overwrite` argument, which is now always `false`. Do not rename the
variable; its meaning ("the uploader consented to replace") is still accurate, only what it now
triggers has changed.

- [ ] **Step 5: Run the full test suite and type-check**

Run: `npx heft test --clean 2>&1 | tail -40` — expect PASS, no dedicated new tests (component
change, no unit tests for React components in this project, same as Tasks 5/6).
Run: `npx tsc --noEmit 2>&1 | tail -40` — expect no new errors.

- [ ] **Step 6: eslint**

Run: `npx eslint src/webparts/form/components/Form.tsx 2>&1 | tail -60` — confirm no NEW warnings
beyond this file's already-documented pre-existing three (unused `file` var, `max-lines`, missing
return type).

- [ ] **Step 7: Commit**

```bash
git add src/webparts/form/components/Form.tsx
git commit -m "feat: proxy the staging-replace overwrite through deletion-by-proxy

Files/Add(overwrite=true) needed EditListItems, which PIC is also
losing. Re-expressed as: recycle the clashing draft via the existing
2026-09-17 proxy-deletion flow, wait for it to actually be gone, then
an ordinary Add-only upload under the same name. Depends on CRS --
Execute approved deletion actually running, same as every other
consumer of that flow.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

## Phase 4 — Wire the approve/reject write path

### Task 7: `ApprovalDocument.tsx` — write a decision row instead of MERGEing status directly

**Files:**
- Modify: `src/webparts/approvalDocument/components/ApprovalDocument.tsx`

- [ ] **Step 1: Read the current `approve()`/`reject()` implementation in full**

Confirm exactly what today's direct MERGE sends (`OData__ModerationStatus`, `ApprovedBy`, the
moderation comment on reject) and exactly which checks run before it (the destination-folder
guard, `documentsFileClash`, anything else). Every one of those checks is a READ and stays
completely unchanged by this task — only the final write changes.

- [ ] **Step 2: Replace the direct MERGE with a `writePendingDecision` call**

After every existing check has passed (same as today), replace the `OData__ModerationStatus` MERGE
with:

```typescript
      const result = await writePendingDecision(context.spHttpClient, siteUrl, {
        itemUniqueId:
          typeof item.UniqueId === "string" ? item.UniqueId : "",
        decision: approve ? "Approved" : "Rejected",
        decidedBy: (context.pageContext.user.email ?? "").toLowerCase(),
        decidedAt: new Date(),
        rejectionComment: approve ? undefined : comment,
        isSelfApprove: false,
      });
      if (!result.ok) {
        setResult({
          ok: false,
          text: `Could not record your decision (HTTP ${result.status ?? "?"}). Nothing has changed — try again.`,
        });
        return;
      }
```

**Note the failure message: "Nothing has changed."** This must be true — a failed
`writePendingDecision` call means no row was written, so the document is genuinely untouched, same
as it was before the click. Confirm this by reading `writePendingDecision`'s own contract (Phase 2,
Task 4) — it never partially writes.

- [ ] **Step 3: The UI no longer shows an immediate "Approved"/"Rejected" result — it shows "Recorded, applying shortly"**

Since the actual status flip now happens asynchronously (the flow, not this click), the success
path can no longer claim the document IS approved — only that the decision was recorded. Update
the success message accordingly (e.g. "Your decision has been recorded and will be applied
shortly."), and rely on this page's existing live-refresh mechanism (My Submissions and the
Requests page already poll and refresh themselves) to reflect the real status once the flow
catches up. If `ApprovalDocument.tsx` does not already have its own live-refresh polling, check
whether one is needed here too — read `shared/liveRefresh.ts` and consider whether this page needs
the same hook, since without it, an approver who stays on this exact page after deciding would see
a stale "still pending" state until they navigate away and back.

- [ ] **Step 4: Run the full test suite**

Run: `npx heft test --clean 2>&1 | tail -40`
Expected: PASS.

- [ ] **Step 5: Type-check**

Run: `npx tsc --noEmit 2>&1 | tail -40`
Expected: no new errors.

- [ ] **Step 6: Commit**

```bash
git add src/webparts/approvalDocument/components/ApprovalDocument.tsx
git commit -m "feat: ApprovalDocument.tsx records a decision instead of deciding directly

Every existing check (destination-folder guard, clash checks) is
unchanged -- they are reads. Only the final status-flip write moves to
CRS -- Apply pending decisions (Power Automate, not yet built).

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 8: `BulkApprovePanel.tsx` — the identical change, per file

**Files:**
- Modify: `src/extensions/bulkApprove/components/BulkApprovePanel.tsx`

- [ ] **Step 1: Read the current bulk-approve MERGE loop**

Confirm the per-file sequence: the existing checks (shared with `ApprovalDocument.tsx` via
`shared/approvalGuards.ts`, per this project's own established "one implementation, never
re-inlined" rule) run first, then the direct MERGE per file.

- [ ] **Step 2: Replace the per-file MERGE with a `writePendingDecision` call**

Same shape as Task 7, Step 2, called once per file in the existing loop. Preserve the existing
per-file error reporting shape (this panel already names each skipped/failed file individually) —
a failed `writePendingDecision` for one file in a batch of fifty must be reported against THAT
file, not abort the whole run, matching how a failed clash check is already handled today.

- [ ] **Step 3: Update the panel's own success/progress messaging**

Same reasoning as Task 7 Step 3 — the panel currently reports "Approved" per file as it happens;
after this change it should report "Recorded" (or similar), since the actual approval is now
asynchronous.

- [ ] **Step 4: Run the full test suite**

Run: `npx heft test --clean 2>&1 | tail -40`
Expected: PASS.

- [ ] **Step 5: Type-check**

Run: `npx tsc --noEmit 2>&1 | tail -40`
Expected: no new errors.

- [ ] **Step 6: Commit**

```bash
git add src/extensions/bulkApprove/components/BulkApprovePanel.tsx
git commit -m "feat: BulkApprovePanel.tsx records decisions instead of deciding directly

Same change as ApprovalDocument.tsx, per file in the existing loop.
Shared checks in shared/approvalGuards.ts are unchanged.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

## Phase 5 — Self-approve sequencing

### Task 9: `Form.tsx` — defer the self-approve decision until tagging is confirmed done

**Files:**
- Modify: `src/webparts/form/components/Form.tsx`

- [ ] **Step 1: Read the current `autoApproveOwnUpload` block in full**

Find where this fires today — after tagging succeeds, checking `probeFolderApproveAccess`
(`EffectiveBasePermissions`, a read), then MERGEing the status directly if eligible. Confirm the
eligibility probe's exact call shape, since Step 3 below keeps it completely unchanged.

- [ ] **Step 2: After Task 5's change, this block now runs after `writeSubmissionRecord` (the tag
      request) rather than after a direct tag write — poll for `TagStatus = "Tagged"` before
      proceeding**

```typescript
      if (autoApproveOwnUpload && tagged) {
        const eligible = await probeFolderApproveAccess(/* unchanged args */);
        if (eligible === "granted") {
          // Poll for the tag flow to actually finish before deciding — writing the decision
          // before tagging completes risks the document routing with blank metadata, the exact
          // defect this project already hit once (2026-08-26, SDG).
          const tagStatus = await pollForTagStatus(
            context.spHttpClient,
            siteUrl,
            refs.fileId,
          );
          if (tagStatus === "Tagged") {
            await writePendingDecision(context.spHttpClient, siteUrl, {
              itemUniqueId:
                typeof item.UniqueId === "string" ? item.UniqueId : "",
              decision: "Approved",
              decidedBy: (context.pageContext.user.email ?? "").toLowerCase(),
              decidedAt: new Date(),
              isSelfApprove: true,
            });
          }
          // tagStatus === "Failed" or the poll timed out: no decision row is written. The
          // document stays Pending; the failed/incomplete tag is visible via the existing
          // CRS Submissions row, same as any other tagging failure.
        }
      }
```

- [ ] **Step 3: Write the `pollForTagStatus` helper**

Add to `Form.tsx` (local helper, not exported — this polling shape is specific to the self-approve
sequencing need, unlike the general-purpose `liveRefresh.ts` hook, which is UI-refresh polling on a
mounted page, not a one-shot wait during an in-flight upload):

```typescript
  /**
   * Polls the submission record for TagStatus, up to a bounded number of attempts, so self-approve
   * never fires before tagging is confirmed done. Returns "Tagged", "Failed", or "timeout" — never
   * throws, and a timeout is treated the same as "do not self-approve" by the caller.
   */
  async function pollForTagStatus(
    sp: SPHttpClient,
    siteUrl: string,
    fileId: string,
  ): Promise<"Tagged" | "Failed" | "timeout"> {
    const MAX_ATTEMPTS = 6;
    const DELAY_MS = 15000;
    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
      const record = await readSubmissionRecordByFileId(sp, siteUrl, fileId);
      if (record?.tagStatus === "Tagged") return "Tagged";
      if (record?.tagStatus === "Failed") return "Failed";
      await new Promise((resolve) => setTimeout(resolve, DELAY_MS));
    }
    return "timeout";
  }
```

**This assumes a `readSubmissionRecordByFileId` function exists or needs to be added to
`spSubmissionRecords.ts`.** Check whether an equivalent read-by-`SubmissionFileId` function already
exists there (this project's existing pattern for reading records back — the same list is already
read elsewhere, e.g. My Submissions' own record-merge logic). If it does not exist, add it to
`spSubmissionRecords.ts` in this same task, following the exact same `listBase()`/`GET_HEADERS`
pattern as `readPendingDecision` in Phase 2 Task 4, filtering on `SubmissionFileId eq '<fileId>'`.

**6 attempts x 15 seconds = 90 seconds max wait, matching the design doc's "up to two poll cycles"
(~2 minutes) estimate for self-approve, with margin.** If this number needs tuning once the actual
flow's real-world timing is observed live, that's a one-line change here, not a redesign.

- [ ] **Step 4: Run the full test suite**

Run: `npx heft test --clean 2>&1 | tail -40`
Expected: PASS.

- [ ] **Step 5: Type-check**

Run: `npx tsc --noEmit 2>&1 | tail -40`
Expected: no new errors.

- [ ] **Step 6: Commit**

```bash
git add src/webparts/form/components/Form.tsx src/shared/spSubmissionRecords.ts
git commit -m "feat: self-approve waits for tagging to be confirmed done

Prevents a self-approved document from routing with blank metadata --
the exact defect already hit once on 2026-08-26 (SDG). Eligibility
probe (a read) is unchanged; only the decision write is deferred.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

## Phase 6 — Power Automate flows and the permission cutover (runbook, not TDD)

**This phase is not unit-testable.** Power Automate flows are built by hand in the designer and
verified with real test runs — every other flow in this project (the deletion proxy, the audit
flows, Auto-route) was built this way. Do not skip this phase's live-testing steps by treating a
successful build as equivalent to a passing test; a flow that saves cleanly can still be wired
wrong, exactly as happened repeatedly during the 2026-09-17 deletion-proxy work in this same
project (a wrong library reference, a stray CRLF in an expression, a duplicate flow left active).

### Task 10: Build `CRS — Apply pending tags`

- [ ] **Step 1: Sign in to Power Automate as `crs@sdguthrie.com`** before creating this flow — its
      connection is what makes every action run as that account. Building it signed in as anyone
      else means it silently stops working the day that person's password changes.

- [ ] **Step 2: Trigger** — *When an item is created or modified*, on `CRS Submissions`.

- [ ] **Step 3: Trigger condition** (Settings -> Trigger Conditions — type it, don't paste, to
      avoid an invisible trailing line break, per this project's own repeated experience with
      exactly that class of bug):
      ```
      @and(not(equals(triggerOutputs()?['body/TagPayload'], null)), equals(triggerOutputs()?['body/TagStatus'], ''))
      ```
      This fires only on a row that has a tag payload and has NOT yet been marked
      `Tagged`/`Failed`.

- [ ] **Step 4: Action — `GetFile`** (Send an HTTP request to SharePoint):
      - Method: `GET`
      - URI: `_api/web/GetFileById(guid'@{triggerOutputs()?['body/ItemUniqueId']}')?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields,ListItemAllFields/ParentList`
      - Headers: `Accept: application/json;odata=nometadata`

      This resolves which LIST/library the file actually lives in (needed because
      `validateUpdateListItem` is called against a specific list title, not just a file GUID) and
      its own item id within that list.

- [ ] **Step 5: Action — `ApplyTags`** (Send an HTTP request to SharePoint), `runAfter GetFile:
      ['Succeeded']`:
      - Method: `POST`
      - URI:
        ```
        _api/web/lists/getbytitle('@{body('GetFile')?['ListItemAllFields']?['ParentList']?['Title']}')/items(@{body('GetFile')?['ListItemAllFields']?['Id']})/validateUpdateListItem
        ```
      - Headers: `Accept: application/json;odata=nometadata`, `Content-Type: application/json`
      - Body:
        ```json
        {
          "formValues": @{json(triggerOutputs()?['body/TagPayload'])},
          "bNewDocumentUpdate": false
        }
        ```
        (`TagPayload` is stored as a JSON string on the row — `json(...)` parses it back into an
        array for the request body.)

- [ ] **Step 6: Action — `CheckResult`** (Compose), `runAfter ApplyTags: ['Succeeded', 'Failed']`:
      ```
      @and(equals(outputs('ApplyTags')['statusCode'], 200), not(contains(string(body('ApplyTags')), 'HasException\":true')))
      ```
      `validateUpdateListItem` returns HTTP 200 even on a per-field error (gotcha #4 in this
      project's own established notes) — a real failure is inside the response body, not the
      status code, so this must check both.

- [ ] **Step 7: Condition — `TagSucceeded`**, `runAfter CheckResult: ['Succeeded']`, condition =
      `outputs('CheckResult')` is equal to `true`:
      - **True branch — `MarkTagged`** (Update item, `CRS Submissions`, `id =
        triggerOutputs()?['body/ID']`): `item/TagStatus` = `Tagged`.
      - **False branch — `MarkFailed`** (Update item, same list, same id): `item/TagStatus` =
        `Failed`, `item/TagError` = `string(body('ApplyTags'))` (the raw response — never
        swallow it; this project's own history is full of cases where a generic failure message
        cost hours a real error string would have saved).

- [ ] **Step 8: Save.** Confirm in the flow's own connection panel (same check performed for the
      deletion-proxy flow) that it reads **"Connected to crs@sdguthrie.com."** on both HTTP
      actions.

- [ ] **Step 9: Turn the flow On.**

---

### Task 11: Build `CRS — Apply pending decisions`

- [ ] **Step 1: Sign in as `crs@sdguthrie.com`**, same as Task 10.

- [ ] **Step 2: Trigger** — *When an item is created or modified*, on `CRS Pending Decisions`.

- [ ] **Step 3: Trigger condition:**
      ```
      @equals(triggerOutputs()?['body/Status'], '')
      ```
      Fires only on a row not yet marked `Applied`/`Failed`.

- [ ] **Step 4: Action — `GetFile`** (identical shape to Task 10 Step 4, resolving the library and
      item id from `triggerOutputs()?['body/ItemUniqueId']`).

- [ ] **Step 5: Condition — `IsRejection`**, `runAfter GetFile: ['Succeeded']`, condition =
      `triggerOutputs()?['body/Decision']` equals `Rejected`.

- [ ] **Step 6: True branch (Rejected) — `RejectItem`** (Send an HTTP request to SharePoint),
      `POST` to the item's `/items(<id>)`, body:
      ```json
      {
        "OData__ModerationStatus": 1,
        "OData__ModerationComments": "@{triggerOutputs()?['body/RejectionComment']}"
      }
      ```
      Headers: `Accept: application/json;odata=nometadata`, `Content-Type: application/json`,
      `X-HTTP-Method: MERGE`, `IF-MATCH: *`.

      **Do NOT combine `OData__ModerationStatus` with any other field in this same MERGE** — this
      project has already hit, and documented, that SharePoint's REST API rejects a MERGE that
      sets moderation status alongside an unrelated field with *"You cannot change moderation
      status and set other item properties at that same time."* `OData__ModerationComments` is a
      moderation field itself, so it's fine alongside `OData__ModerationStatus`; nothing else may
      join this call.

- [ ] **Step 7: False branch (Approved) — `ApprovedByFirst`** (Update item, same item, SEPARATE
      call from the status flip, for the same reason as Step 6's warning): `item/ApprovedBy` =
      `triggerOutputs()?['body/DecidedBy']`.

- [ ] **Step 8: `ApproveItem`** (Send an HTTP request to SharePoint), `runAfter ApprovedByFirst:
      ['Succeeded', 'Failed']` (an audit-style stamp failing must never block the actual approval —
      same rule this project applies everywhere else a secondary write sits beside a primary one):
      `POST` to the item, body `{"OData__ModerationStatus": 0}`, same headers/MERGE shape as
      Step 6.

- [ ] **Step 9: `CheckApplied`** (Compose), `runAfter` both `RejectItem` and `ApproveItem`
      (whichever branch ran) with `['Succeeded', 'Failed']`:
      ```
      @or(equals(outputs('RejectItem')?['statusCode'], 204), equals(outputs('ApproveItem')?['statusCode'], 204))
      ```
      (A MERGE typically returns 204 No Content on success — confirm this against a real test run
      rather than trusting the number written here; if it differs, use whatever status code the
      actual successful call returns, visible in the run's own history.)

- [ ] **Step 10: `DecisionSucceeded`** Condition, same True/False -> `MarkApplied`/`MarkFailed`
      shape as Task 10 Step 7, writing back onto the `CRS Pending Decisions` row's own `Status`/
      `StatusError`.

- [ ] **Step 11: Save, confirm the connection reads `crs@sdguthrie.com` on every HTTP action,
      turn the flow On.**

---

### Task 12: End-to-end verification — BEFORE the permission cutover

**Do this with the old permission levels still in place.** The code no longer writes directly
(Phases 3–5 already removed those calls), so this test genuinely exercises the new proxy path
regardless of what rights the test accounts still hold — nothing here is invalidated by testing
before the cutover.

- [ ] Upload a normal document as a PIC test account. Confirm: physical file lands correctly
      (unchanged), `CRS Submissions` row shows `TagStatus` blank momentarily then `Tagged` within
      about a minute, and the document's actual metadata columns match what was chosen in the form.

- [ ] Upload an HC document. Confirm the same flow (no separate HC clone) handles it correctly.

- [ ] Approve a document via the Preview page as a human approver. Confirm: `CRS Pending Decisions`
      row shows `Status` blank then `Applied`, the document's `OData__ModerationStatus` flips, and
      Auto-route routes it exactly as it did before this change.

- [ ] Reject a document via the Preview page, with a comment. Confirm the comment lands on
      `OData__ModerationComments`.

- [ ] Approve several documents via the bulk sidebar in one run. Confirm each gets its own decision
      row and each applies correctly, and a deliberately-clashing one in the batch is still
      reported individually (per the panel's existing per-file error handling).

- [ ] Upload as a Head of Unit eligible for self-approve. Confirm: it does NOT self-approve
      immediately, waits for `TagStatus = Tagged`, then writes and applies the decision — confirm
      the whole thing completes with no extra click from the uploader, just a longer "Pending"
      window than an ordinary tag-only upload.

- [ ] Deliberately break a tag payload (e.g. a malformed date) for one test upload. Confirm
      `TagStatus = Failed` with a real, readable `TagError`, not a generic message.

### Task 13: The permission-level cutover

**This is the single highest-risk step in the whole rollout.** Do not do this on the same day as
any other major deploy. Confirm Task 12 passed completely before starting.

- [ ] In Site Settings -> Permission Levels on the target site, open `CRS Upload`. Untick
      `Edit Items`. Confirm `Add Items` remains ticked. Save.

- [ ] Open `CRS Approve`. Untick `Edit Items`. Confirm `Approve Items` remains ticked. Save.

- [ ] Immediately after saving, sign in as a PIC test account and upload one document. Confirm it
      still works exactly as it did in Task 12 (upload succeeds, tags apply within a minute).

- [ ] Sign in as an approver test account and approve one document. Confirm it still works.

- [ ] Sign in as either test account and attempt to rename a file, or change one of its metadata
      columns, directly in the native SharePoint library view. **Confirm this is now refused with
      an access-denied error.** This is the actual point of the whole feature — confirm it, don't
      assume it from the permission-level change alone.

- [ ] Repeat this entire task on every other site this package is deployed to (do not assume one
      site's cutover applies anywhere else — this is a per-site SharePoint setting, not something
      the app package carries).

---

## Self-review

**Spec coverage:**
- PIC's tagging write, both normal and HC, both `Form.tsx` and `BulkUpload.tsx` — Tasks 5, 6.
- Approver's decision write, both screens — Tasks 7, 8.
- Self-approve sequencing — Task 9.
- Schema for both lists — Tasks 1–4.
- The two flows — Tasks 10, 11.
- The permission cutover, sequenced last, with its own testing checklist — Tasks 12, 13.
- Non-goals (Share, deletion, Auto-route) — none of the tasks above touch any of them; confirmed by
  absence.

**Gap, called out rather than silently left:** the design doc's own "Open questions" section (where
does a `Failed` state surface to a human, beyond the raw list row) is not resolved by this plan —
it was left open in the spec deliberately, and this plan does not invent an answer for it. If that
matters before shipping, it needs its own small follow-up task (likely a filtered view or a banner
on an admin screen reading `TagStatus`/`Status` = `Failed` across both lists) — not included here
since the spec explicitly flagged it as undecided rather than as a requirement.

**Placeholder scan:** no "TBD"/"add validation"/"handle edge cases" language — every step shows
real code or a real runbook value. Two steps are explicitly flagged as needing verification against
the live codebase before trusting the exact value shown (Task 4's `FieldTypeKind` mapping, Task 11
Step 9's status-code assumption) — these are honest uncertainty flags, not placeholders, and each
names exactly what to check and how.

**Type consistency:** `SubmissionRecord.tagPayload`/`tagStatus`/`tagError` (Task 1) are read the
same way in `spSubmissionRecords.ts` (Task 3) and written the same way in `Form.tsx`/
`BulkUpload.tsx` (Tasks 5, 6). `PendingDecision`'s shape (Task 2) matches `writePendingDecision`'s
parameter type (Task 4) and `ApprovalDocument.tsx`/`BulkApprovePanel.tsx`'s calls to it (Tasks 7,
8) exactly — `itemUniqueId`, `decision`, `decidedBy`, `decidedAt`, `rejectionComment`,
`isSelfApprove` used consistently throughout.

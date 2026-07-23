# Share Guard — Retirement

**Date:** 2026-07-23
**Status:** approved, not yet implemented
**Retires:** `2026-07-20-share-guard-design.md`, `2026-07-20-share-guard-runbook.md`

## Why

Share Guard was built on the premise, stated in its own design, that:

> A web part/extension cannot intercept SharePoint's native Share button or hold a sharing link
> for approval. Native sharing grants access the instant a link is created; there is no supported
> "pending" hook.

**That premise is wrong for this site's configuration.** With *"Only site owners can share files,
folders, and the site"* enabled, SharePoint does not grant on share — it converts the member's
share attempt into an **owner-approved request**.

Observed live on 2026-07-21: a member attempted to share the `Group Head Office` folder in the
Documents library. No access was granted. The site owner received a SharePoint Online email —
*"Kaisya wants to share 'Group Head Office' with Elyn Chai"* — with **Accept** / **Decline**
actions and the requested resource URL.

That is precisely the gate Share Guard was written to provide. The native flow already delivers
it, so the custom implementation is redundant.

Confirmation that the lockdown is active: the same email carries the line *"Tired of approving
requests? Allow members to share with others without requiring owner approval."* That prompt only
renders when owner approval is currently **required**.

## Coverage check

Every component of the Share Guard design has a native equivalent:

| Share Guard component | Native equivalent |
|---|---|
| Guarded "Request to Share" command set button | Native Share button — auto-converts to a request under the lockdown |
| `DMS Share Requests` list (request store + audit trail) | Site Settings → **Access requests and invitations** |
| Approval Queue web part (admin-only) | The same page, plus the Accept/Decline email |
| Requester chooses Read / Edit | Owner chooses the permission level on Accept |
| Reason field | Message box in the native share dialog |
| Internal recipients only | Governed by tenant + site external sharing settings |

No capability is lost. The decision is to delete custom code that duplicates platform behaviour.

## Decision

Retire Share Guard entirely. Rely on native share approval, which depends on **one site setting
staying enabled** (see Operational risks).

Secondary effect: `RequestShare.tsx` is the last consumer of `msGraphClientFactory` in the
solution. Removing it takes `webApiPermissionRequests` to **zero Graph scopes** — see
`2026-07-23-native-sharepoint-groups-design.md`. This is a consequence of the retirement, **not
the reason for it**. Share Guard would be redundant regardless of the permission question.

## Scope of removal

**Delete:**

| Path | Kind |
|---|---|
| `src/webparts/requestShare/` | web part — request form |
| `src/webparts/shareApproval/` | web part — approval queue |
| `src/extensions/shareRequest/` | ListView Command Set — the toolbar button |
| `src/shared/shareGuard.ts` | `parseShareTarget`, `buildRequestPayload`, `AccessLevel` |
| `src/shared/shareGuard.test.ts` (if present) | its unit tests |

**Edit:**

| File | Change |
|---|---|
| `config/config.json:61-84` | drop `share-approval-web-part`, `request-share-web-part`, `share-request-extension` bundles |
| `config/config.json:93-95` | drop `ShareApprovalWebPartStrings`, `RequestShareWebPartStrings`. **Keep `ControlStrings`** only if another web part still uses `@pnp/spfx-controls-react` — verify before removing |
| `config/package-solution.json` | remove component ids `443fdf05-34bd-42e4-b469-244197c769ac`, `f15ba20c-f89d-4eff-b081-83a3479de1ac`, `d239b54b-47e5-430b-be20-7deecf0bd652`; bump solution version |
| `package.json` | drop `@pnp/spfx-controls-react` **only if** no other web part imports it |

**Tenant cleanup (manual, after deploy):**

- Remove the Approval Queue web part from any page it is placed on.
- The `DMS Share Requests` list: **do not delete immediately.** It holds the audit trail of
  requests already processed. Set it to hidden/read-only, retain per the client's record-keeping
  policy, delete later.

## Operational risks

These are handover-documentation items, not code changes. Both must be written into the DMS admin
handover.

1. **The gate is one click from being disabled.** Every approval email contains *"Allow members to
   share with others without requiring owner approval."* Clicking it disables owner approval site
   wide and silently removes the protection this retirement depends on. **Never click this link.**
   With Share Guard, the enforcement lived partly in code; now it lives entirely in one setting.

2. **Accepting a request punches a hole in the unit ACL model.** The observed request was for a
   Documents-library folder whose permissions the DMS unit-group model manages. Accepting grants
   a per-item exception outside that model. This was equally true under Share Guard — but the
   native flow makes it a one-click Accept inside an email, which invites less scrutiny than a
   dedicated admin queue did. Owners must treat Accept as a real access decision.

3. **The setting must be verified after any site reconfiguration.** Add to the reconciliation /
   handover checklist: confirm Site Permissions → Sharing settings is still *"Only site owners can
   share files, folders, and the site."*

## Verification

After removal:

- `npm run build` succeeds with no unresolved imports from the deleted modules.
- The solution's `webApiPermissionRequests` is empty and no code calls `msGraphClientFactory`
  (`grep -r msGraphClientFactory src/` returns nothing).
- The DMS libraries no longer show the custom "Request to Share" toolbar button.
- A member attempting to share still produces the native owner-approval email — re-test the
  2026-07-21 scenario end to end, since the entire retirement rests on this behaviour.
- Existing `DMS Share Requests` rows remain readable for audit.

## Out of scope

- Changing tenant or site external-sharing settings.
- Any change to the unit-folder ACL model, `DMS Group Map`, or the upload/approval pipeline.
- Deleting the `DMS Share Requests` list (retention decision belongs to the client).

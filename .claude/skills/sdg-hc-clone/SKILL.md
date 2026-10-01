---
name: sdg-hc-clone
description: Use when a Power Automate flow is copied (Save As) to make its Highly Confidential (HC) twin, or when a fix made to a normal flow must also be applied to its HC twin.
---
# Clone a flow to its HC twin

Every HC clone in this project has shipped with a library reference nobody swapped. Check every item.

1. **Swap every library reference**, not just the trigger:
   - Trigger list → `Approval for Highly Confidential Document` (approval side) or `Highly Confidential Document` (approved side)
   - Every `Get item` / `Get items` / `Get file` list name
   - Every `getbytitle('...')` inside HTTP actions
   - Every list **GUID** inside HTTP URIs (item ids are per list, so a wrong list edits a different document)
   - URL segments: `ApprovalDocument` → `HCApprovalDocument`, `Shared Documents` → `HCDocuments`
2. **Roles**: `Role eq 'APR'` → `Role eq 'APRHC'`.
3. **Links**: add `&lib=hc` to any ApprovalDocument link.
4. **Do NOT swap shared lists**: `CRS Submissions`, `CRS Requests`, `CRS Audit Log`, `CRS Group Map` are shared by both.
5. **Audit text**: "Moved to Highly Confidential Document", not "Documents".
6. Verify with `/sdg-flow-check` on a fresh export. The LIBRARIES line must list only HC libraries and shared lists.
7. Test with one real HC document end to end.

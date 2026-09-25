# Permissions model

- Nobody holds a delete grant. Deletes = a self-approved `CRS Requests` row → the flow `CRS — Execute approved deletion` recycles it as `gdc@sdguthrie.com`.
- Tagging = `TagPayload` on `CRS Submissions` → the flow `CRS — Apply pending tags` (as gdc).
- `gdc@sdguthrie.com` is the service account. Never remove it from Owners. `crs@sdguthrie.com` is a normal account.
- `PERSONAS` in `shared/groupMapModel.ts` is the only role list. A persona change needs the old groups deleted and re-created.

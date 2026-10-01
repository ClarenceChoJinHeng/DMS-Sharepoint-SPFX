---
name: sdg-logic-change
description: Use before changing any behaviour or rule in the code (not just styling or wording) - permissions, routing, filters, what a button does, what gets written.
---
# Safe logic change

1. **Read the component note** in the vault (`Code Map.md` → `Components/<name>.md`). Is this logic listed under "Firm logic"?
   - If yes, and the user didn't clearly ask to change it: **stop and ask**.
2. **Find every caller.** Grep the function or value name across `src/`. Shared rules live in `src/shared/*.ts`, and some are used by several web parts (for example `approvalGuards.ts` is used by both the approval page and bulk approve).
3. **Check the other side:**
   - Does a Power Automate flow read or write the same list/column? (See the vault note `Power Automate Flows.md`.)
   - Does another screen show the same data?
   - Is there an HC twin (normal + HC)?
4. **Tell the user the side effects** in bullets before editing.
5. Change the pure logic in `src/shared/x.ts` first, and update `x.test.ts`.
6. Verify: `npx tsc --noEmit`, `npx heft test --clean`, lint on the changed files (no new warnings).
7. Update the vault with `/sdg-obsidian-update`.

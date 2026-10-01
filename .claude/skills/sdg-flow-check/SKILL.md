---
name: sdg-flow-check
description: Use when reading, reviewing or advising on a Power Automate flow for this project, or when a flow misbehaves (no email, wrong actor, double rows, failed run).
---
# Check a Power Automate flow from its export

Flows are not in source control. The only truth is the export in `PowerAutomateFlowsSDG/`.

1. Find the newest export: `ls PowerAutomateFlowsSDG | grep -i "<part of flow name>"`
   - If it's older than the last edit the user made, ask them to export it again first.
2. Run the checker:
   `node .claude/skills/sdg-flow-check/check-flow.js "PowerAutomateFlowsSDG/<file>.zip"`
3. Read the output:
   - **ACTIONS**: the tree, with `after:` showing each run-after. An audit write or email should usually run after both `Succeeded` and `Failed` of a lookup.
   - **LIBRARIES**: every `getbytitle(...)`. For an HC flow, all of them must be HC libraries.
   - **ACCOUNTS**: hardcoded emails (gdc is the service account).
   - **WARNINGS**: known traps. Check each one, but first read the flow's "Firm logic" in the vault's `Power Automate Flows.md`. Some warnings are deliberate (e.g. NotifyApprovers' owner/admin lookups).
4. To see one action in full:
   `unzip -p "<zip>" "$(unzip -Z1 "<zip>" | grep definition.json)" | node -e "let s='';process.stdin.on('data',d=>s+=d).on('end',()=>console.log(JSON.stringify(JSON.parse(s).properties.definition,null,1)))" | grep -n -A30 '"<ActionName>"'`
5. When giving edit instructions, follow `rules/` and the vault note `Gotchas - Power Automate.md`:
   - Type expressions in the fx box, never pasted with a line break.
   - A switch case value is a typed literal.
   - Condition values like `0` go through fx.
   - Header keys have no colon.
6. After the user saves, ask for a fresh export and run the checker again. A screenshot doesn't prove the edit landed.

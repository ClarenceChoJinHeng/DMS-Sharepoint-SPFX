---
name: sdg-handoff
description: Use when the context window is nearly full (the user says ~80%, or Claude Code warns), when the session is very long, or when the user asks to pause, hand off, or continue in a new session. Also use at the START of a session when the user says "continue from <handoff>".
---
# Handoff to a new session

## Writing the handoff (end of a session)
1. Stop at a safe point: no half-edited file, and no command left running.
2. Copy `handoff-template.md` from this folder to the vault:
   `Specs/YYYY-MM-DD-handoff-<short-topic>.md`
   (vault: `C:\Users\clare\OneDrive - Trinergy\Documents\Obsidian\Personal-Obsidian\Trinergy-Notes\SDG-Claude-Config`)
3. Fill it in: short bullets, exact file paths and function names, no story.
4. Add one line to the vault's `Open Items.md`: `- [ ] In progress: <topic>, see [[YYYY-MM-DD-handoff-<short-topic>]]`
5. If there are code changes, list what is committed and what is not.
6. Tell the user, in bullets:
   - Handoff saved: `<file name>`
   - Start a **new** session (`/clear` or the new-conversation button). **Don't** use `/compact` or resume.
   - In the new session, type: `continue from <file name>`

## Picking up (start of a new session)
1. Read ONLY the named handoff note. Don't search the old conversation.
2. Read the files it lists under "Read these first". Nothing else until needed.
3. Check its "Verify first" items (branch, uncommitted changes).
4. Carry on from "Next steps". When it's finished, tick the Open Items line.

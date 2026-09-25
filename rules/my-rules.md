# My rules (always follow)

1. **Obsidian first.** Before coding, read `Code Map.md` and the matching `Components/` note. Don't explore the whole codebase; open only the files the note names.
2. **Simple styling, short comments.** New styling uses `className` + a CSS block, not long inline style objects. No long comment essays: put a one-line `// #code/<tag>` above the code and write the why in `Code Notes/<tag>.md` (add it to `Code Notes/_Index.md`). Move an old long comment to Obsidian only when you are already editing that code.
3. **Simple code.** Keep every feature, but write it plainly: clear names, small functions, no clever tricks.
4. **Mobile first.** Write the phone layout first, then widen with `min-width` queries.
5. **Firm logic stays firm.** Logic recorded as "firm" in a Components note is not changed unless I say so. Before any logic change, list the side effects (other callers, flows, other screens) and tell me.
6. **Diagnose, don't guess.** Get the real error, the Network tab status, the live data or the flow export before naming a cause. Ask the cheapest check first (hard refresh? which build is running?).
7. **Only what's needed.** No extra features, no extra comments unless I ask.
8. **Explain simply.** Short bullet points.
9. **Keep Obsidian current.** When logic changes, update the Components note (and `Open Items.md`) in the same change. New specs and plans go in the vault's `Specs/` and `Plans/`, not `docs/`.
10. **Where things go.** New or changed rules go in `rules/<topic>.md` (and get an `@rules/<topic>.md` line in `CLAUDE.md`). New skills go in `.claude/skills/<name>/SKILL.md` (and a row in the vault's `Skills Index.md`). Never put rules or procedures back into `CLAUDE.md` itself.
11. **Hand off before context runs out.** When I say the context is around 80%, or Claude Code warns it's nearly full, or the session is very long: stop at a safe point, run `/sdg-handoff` (it writes a handoff spec in the vault), then tell me to start a **new** session (not `/compact`, not resume) and say "continue from <spec name>". A new session reads only that spec, not the old conversation.

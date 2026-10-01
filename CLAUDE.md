# SDG CRS: Claude Code Project Rules

SPFx 1.23 React web parts + Power Automate flows for SD Guthrie's document system (CRS, formerly DMS).
Client: SD Guthrie · Agency: Trinergy Digital · Dev: Clarence.

## Rules (in `rules/`, one file per topic)
@rules/my-rules.md
@rules/obsidian.md
@rules/build-and-deploy.md
@rules/diagnosing.md
@rules/sharepoint-rest.md
@rules/react-spfx.md
@rules/permissions.md

## Commands
```
nvm use 22
npm run start      # Heft dev server (never gulp)
npm run build      # the ONLY way to package: heft test --clean --production && package-solution
npx tsc --noEmit   # type check
npx heft test --clean
```

## Git
- Work on a branch, never `main`. Commit only when asked.

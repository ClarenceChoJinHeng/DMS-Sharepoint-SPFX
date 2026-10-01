/**
 * Every bundled component must also be REGISTERED in the solution feature.
 *
 * ⚠ THIS IS THE SECOND TIME A COMPONENT SHIPPED BUNDLED BUT UNREGISTERED, and both were silent.
 *   - The `+ New Folder` application customizer sat correct, deployed and inert for months because
 *     no `UserCustomAction` pointed at it (2026-08-16).
 *   - `CRS Requests` (`7a4f1e93-…`) was absent from `componentIds` from the day it was written until
 *     2026-08-20, so it never appeared in the web-part toolbox on any site. The whole deletion and
 *     share workflow was undeployable, and the only symptom was an admin unable to find a web part
 *     that the repo, the specs and CLAUDE.md all said existed.
 *
 * BUNDLING MAKES A COMPONENT AVAILABLE; THE FEATURE IS WHAT REGISTERS IT. Nothing fails, nothing
 * logs, the package deploys cleanly and every other web part works — which is exactly why this has
 * to be a test rather than a habit.
 *
 * Reads the two config files directly. Regex rather than JSON.parse for the MANIFESTS: several carry
 * formatting a strict parser rejects, and this test must not start failing over a stray tab in a
 * file it is only inspecting for an id.
 */

/* SPFx's tsconfig carries no Node types, so `fs`, `path` and `__dirname` are all unavailable — the
   same constraint that rules out `Promise.allSettled` and spreading a Set (CLAUDE.md #3). Declared
   narrowly here rather than adding `@types/node` to the build: this is the only test that touches the
   filesystem, and widening the type environment for every file to serve one test invites Node APIs
   into browser code, where they fail at RUNTIME rather than at compile time.

   Forward slashes throughout — Node accepts them on Windows, so no `path.join` is needed. */
interface NodeFs {
  readFileSync(file: string, encoding: string): string;
}
/* `require` is already in scope from the test environment's own types — re-declaring it is a
   duplicate-identifier error, so only the SHAPE this test uses is declared above. The lint rule
   against require() is disabled for this ONE line: `import` of a Node module does not type-check
   without @types/node, and adding those globally would invite Node APIs into browser code where
   they fail at runtime rather than at compile time. */
// eslint-disable-next-line @typescript-eslint/no-require-imports
const fs = require("fs") as NodeFs;

/* Paths are RELATIVE, resolved by Node against the working directory — which avoids `process`
   too, since SPFx types it as the browser shim with no `cwd`. Jest runs with the project directory
   as its cwd; if that ever stops being true this test fails LOUDLY on a missing file, never by
   silently checking nothing. */

interface Bundled {
  manifestPath: string;
  id: string;
  label: string;
}

function readBundledComponents(): Bundled[] {
  const cfg = JSON.parse(fs.readFileSync("config/config.json", "utf8")) as {
    bundles: Record<string, { components: Array<{ manifest: string }> }>;
  };
  const out: Bundled[] = [];
  for (const name of Object.keys(cfg.bundles ?? {})) {
    for (const c of cfg.bundles[name].components ?? []) {
      const rel = c.manifest.replace(/^\.\//, "");
      const text = fs.readFileSync(rel, "utf8");
      const id = /"id"\s*:\s*"([^"]+)"/.exec(text);
      const title = /"title"\s*:\s*\{\s*"default"\s*:\s*"([^"]+)"/.exec(text);
      const alias = /"alias"\s*:\s*"([^"]+)"/.exec(text);
      if (!id) throw new Error(`${rel} has no "id"`);
      out.push({
        manifestPath: rel,
        id: id[1].toLowerCase(),
        label: (title && title[1]) || (alias && alias[1]) || rel,
      });
    }
  }
  return out;
}

function readComponentIds(): string[] {
  const sol = JSON.parse(fs.readFileSync("config/package-solution.json", "utf8")) as {
    solution: { features: Array<{ componentIds: string[] }> };
  };
  const ids: string[] = [];
  for (const f of sol.solution.features ?? []) {
    for (const id of f.componentIds ?? []) ids.push(id.toLowerCase());
  }
  return ids;
}

describe("solution packaging — bundled components are registered in the feature", () => {
  it("registers EVERY bundled component, naming any that are missing", () => {
    const registered = readComponentIds();
    const missing = readBundledComponents().filter((c) => registered.indexOf(c.id) === -1);
    // NAMED, not counted. "1 component is unregistered" sends someone reading eighteen GUIDs; the
    // label is the thing they can act on.
    expect(missing.map((m) => `${m.label} (${m.id}) — ${m.manifestPath}`)).toEqual([]);
  });

  it("has no componentId that no bundled manifest claims", () => {
    // The other direction, and it fails differently: a leftover id is a component that was deleted or
    // never written, and it makes the count look right while the thing itself is absent.
    const bundled = readBundledComponents().map((c) => c.id);
    const orphans = readComponentIds().filter((id) => bundled.indexOf(id) === -1);
    expect(orphans).toEqual([]);
  });

  it("pins CRS Requests specifically, because its absence was invisible for weeks", () => {
    // A regression test for the exact id, so a future edit that rewrites componentIds wholesale
    // cannot quietly drop the one that has already been dropped once.
    expect(readComponentIds()).toContain("7a4f1e93-2c58-4d07-b6a1-9e3c85f2d410");
  });

  it("keeps skipFeatureDeployment false, or nothing in elements.xml ever provisions", () => {
    // Not packaging in the same sense, but the same failure shape and the same file: with this true
    // the solution's feature is not activated on the site, so its element manifests never run — and
    // every web part still works, which is what makes it so hard to spot.
    const sol = JSON.parse(fs.readFileSync("config/package-solution.json", "utf8")) as {
      solution: { skipFeatureDeployment?: boolean };
    };
    expect(sol.solution.skipFeatureDeployment).toBe(false);
  });
});

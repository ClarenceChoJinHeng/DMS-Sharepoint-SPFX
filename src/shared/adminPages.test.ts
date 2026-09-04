import {
  CARDS,
  SitePage,
  allLinks,
  columnCards,
  resolveAll,
  resolveLink,
  suggestedPageName,
  tabFromHash,
} from "./adminPages";

function page(fileName: string, title = ""): SitePage {
  return {
    fileName,
    title: title || fileName.replace(/\.aspx$/i, "").replace(/-/g, " "),
    serverRelativeUrl: `/sites/Example/SitePages/${fileName}`,
  };
}

/** A plausible site: every admin page, named the way a deploy would name them. */
const SITE: SitePage[] = [
  page("Home.aspx"),
  page("Folder-Administration.aspx"),
  page("Folder-Access.aspx"),
  page("Site-Access.aspx"),
  page("Approval-Library-Access.aspx"),
  page("Page-Access.aspx"),
  page("Group-Management.aspx"),
  page("CRS-Configuration.aspx"),
  page("CRS-Audit-Log.aspx"),
  page("Bulk-Upload.aspx"),
  page("Upload-Form.aspx"),
  page("ApprovalDocument.aspx"),
  page("My-Submissions.aspx"),
];

function linkFor(key: string): { key: string; label: string; match: RegExp; tab?: string } {
  const hit = allLinks().filter((l) => l.key === key)[0];
  if (!hit) throw new Error(`no link named ${key}`);
  return hit;
}

describe("the page as designed", () => {
  it("has the three cards the client asked for, and drops CRS Mapping", () => {
    const titles = CARDS.map((c) => c.title);
    expect(titles).toContain("Folder Management");
    expect(titles).toContain("User Access Management");
    /* Retitled 2026-08-30 at the client's request. The CARD is named for what it does; the page it
       opens is still called "CRS Configuration", which is why `self.label` and `match` were left
       alone — see the comment on that entry. */
    expect(titles).toContain("File Type Management");
    expect(titles).toContain("Audit Log");
    // Dropped deliberately: term abbreviations moved under Folder Management.
    expect(titles).not.toContain("CRS Mapping");
  });

  it("makes Folder Management ONE link, because its screens became flow steps", () => {
    // Term Abbreviations / Folder Structure / Reconciliation stopped being destinations on 2026-08-14.
    // Three rows into a flow picker that ignores which was clicked would recreate the confusion the
    // flows removed, one level up. The blurb keeps them discoverable.
    const folders = CARDS.filter((c) => c.key === "folders")[0];
    expect(folders.links).toEqual([]);
    expect(folders.self).toBeDefined();
    expect(folders.blurb.toLowerCase()).toContain("reconciliation");
  });

  it("lists Group Management under User Access, FIRST", () => {
    // Not on the mockup, and required: it is the page that creates the groups, writes their
    // mappings and adds the people, so every other access screen presupposes it.
    const access = CARDS.filter((c) => c.key === "access")[0];
    expect(access.links[0].label).toBe("Group Management");
  });

  // Removed 2026-08-23: its mapping half moved into group creation and its membership half into
  // Group Management's group list, so the card would point at a screen with nothing left to do.
  it("no longer offers Folder Access anywhere on the landing page", () => {
    for (const c of CARDS) {
      for (const l of c.links) expect(l.key).not.toBe("folderAccess");
      if (c.self) expect(c.self.key).not.toBe("folderAccess");
    }
  });

  it("splits the grid into two columns, both populated", () => {
    expect(columnCards(1).length).toBeGreaterThan(0);
    expect(columnCards(2).length).toBeGreaterThan(0);
    expect(columnCards(1).concat(columnCards(2)).length).toBe(CARDS.length);
  });

  it("gives every single-link card a `self` and no rows, and vice versa", () => {
    for (const c of CARDS) {
      if (c.self) expect(c.links.length).toBe(0);
      else expect(c.links.length).toBeGreaterThan(0);
    }
  });

  it("has a unique key per link — they key the property-pane overrides", () => {
    const keys = allLinks().map((l) => l.key);
    const distinct: Record<string, true> = {};
    for (const k of keys) distinct[k] = true;
    expect(Object.keys(distinct).length).toBe(keys.length);
  });
});

describe("resolveLink — the collisions that would send an admin to the wrong page", () => {
  it("sends Folder Management to Folder Administration, not Folder Access", () => {
    // The bug this pattern exists to prevent: a bare /folder/i claims Folder-Access.aspx, so the
    // folder card would open a permissions screen.
    const t = resolveLink(linkFor("folders"), SITE);
    expect(t.state).toBe("resolved");
    if (t.state !== "missing") {
      expect(t.url).toContain("Folder-Administration.aspx");
      expect(t.url).not.toContain("Folder-Access.aspx");
    }
  });

  it("sends Page Access to Page Access, not to another access screen", () => {
    const t = resolveLink(linkFor("pageAccess"), SITE);
    expect(t.state).toBe("resolved");
    if (t.state !== "missing") expect(t.url).toContain("Page-Access.aspx");
  });

  // "libraryAccess" and "siteAccess" are GONE (2026-09-02) — Approval Library Access and Site
  // Access retired as signposts, and the "access" card's links no longer name them. See
  // `ApprovalLibraryAccessPage.tsx` / `SiteAccessPage.tsx` for the retirement itself.

  it("does not confuse Page Access with Group Management", () => {
    const p = resolveLink(linkFor("pageAccess"), SITE);
    const g = resolveLink(linkFor("groups"), SITE);
    if (p.state !== "missing") expect(p.url).toContain("Page-Access.aspx");
    if (g.state !== "missing") expect(g.url).toContain("Group-Management.aspx");
  });

  it("resolves every link on a normally-named site — none left missing", () => {
    const all = resolveAll(SITE);
    const missing = Object.keys(all).filter((k) => all[k].state === "missing");
    expect(missing).toEqual([]);
  });

  it("gives every link a page of its own — no two rows lead to the same place", () => {
    // True again since Folder Management collapsed to one link. Two rows sharing a destination is only
    // acceptable when each names a distinct tab, and no card does that any more.
    const all = resolveAll(SITE);
    const distinct: Record<string, true> = {};
    for (const k of Object.keys(all)) {
      const t = all[k];
      if (t.state === "missing") continue;
      distinct[t.url.split("#")[0]] = true;
    }
    expect(Object.keys(distinct).length).toBe(allLinks().length);
  });
});

describe("resolveLink — missing, ambiguous, overridden", () => {
  it("reports MISSING rather than guessing when no page matches", () => {
    // A dead arrow on a navigation page is the worst failure here: it teaches the client the tool is
    // broken. The row renders disabled instead.
    expect(resolveLink(linkFor("audit"), [page("Home.aspx")]).state).toBe("missing");
    expect(resolveLink(linkFor("audit"), []).state).toBe("missing");
  });

  it("survives a missing page list rather than throwing", () => {
    expect(resolveLink(linkFor("audit"), undefined as unknown as SitePage[]).state).toBe("missing");
  });

  it("picks the SHORTEST file name on several matches, and says it was ambiguous", () => {
    const dupes = [page("CRS-Audit-Log-Copy.aspx"), page("CRS-Audit-Log.aspx")];
    const t = resolveLink(linkFor("audit"), dupes);
    expect(t.state).toBe("ambiguous");
    if (t.state === "ambiguous") {
      expect(t.url).toContain("CRS-Audit-Log.aspx");
      expect(t.others).toEqual(["CRS-Audit-Log-Copy.aspx"]);
    }
  });

  it("still NAVIGATES when ambiguous — a best guess beats a dead row", () => {
    const t = resolveLink(linkFor("audit"), [page("Audit-A.aspx"), page("Audit-BB.aspx")]);
    expect(t.state).toBe("ambiguous");
    if (t.state === "ambiguous") expect(t.url.length).toBeGreaterThan(0);
  });

  it("is deterministic when two names are the same length", () => {
    const a = resolveLink(linkFor("audit"), [page("Audit-B.aspx"), page("Audit-A.aspx")]);
    const b = resolveLink(linkFor("audit"), [page("Audit-A.aspx"), page("Audit-B.aspx")]);
    expect(a).toEqual(b);
  });

  it("matches on the TITLE too, since either half can be renamed", () => {
    const renamed: SitePage[] = [{
      fileName: "SitePage1.aspx",
      title: "CRS Audit Log",
      serverRelativeUrl: "/sites/Example/SitePages/SitePage1.aspx",
    }];
    expect(resolveLink(linkFor("audit"), renamed).state).toBe("resolved");
  });

  it("lets an OVERRIDE win outright, without pattern-checking it", () => {
    // The only fix available without a redeploy on a site whose pages this cannot guess.
    const t = resolveLink(linkFor("audit"), SITE, "/sites/Other/SitePages/Anything.aspx");
    expect(t.state).toBe("resolved");
    if (t.state !== "missing") expect(t.url).toBe("/sites/Other/SitePages/Anything.aspx");
  });

  it("an override rescues a link that would otherwise be missing", () => {
    const t = resolveLink(linkFor("audit"), [], "/sites/Example/SitePages/Log.aspx");
    expect(t.state).toBe("resolved");
  });

  it("ignores a blank or whitespace override rather than navigating nowhere", () => {
    expect(resolveLink(linkFor("audit"), [], "   ").state).toBe("missing");
    expect(resolveLink(linkFor("audit"), [], "").state).toBe("missing");
  });
});

/**
 * The `#tab=` contract is asymmetric on purpose, and that asymmetry is worth pinning.
 *
 * No CARD writes a tab any more — Folder Management is one link since the flows landed. But the READER
 * (`tabFromHash`, used by FolderManager and FolderAdmin) must stay, because links generated by the
 * previous version are in people's bookmarks and on published pages. Keeping the writer here too is what
 * keeps both halves of the format in one module; a parser for a format nothing can produce is the worse
 * of the two shapes.
 */
describe("the tab hash", () => {
  const withTab = { key: "x", label: "X", match: /folder.?admin/i, tab: "reconciliation" };

  it("is appended for a link that names a tab", () => {
    const t = resolveLink(withTab, SITE);
    if (t.state !== "missing") expect(t.url).toContain("#tab=reconciliation");
  });

  it("is absent for a link that does not", () => {
    const t = resolveLink(linkFor("pageAccess"), SITE);
    if (t.state !== "missing") expect(t.url).not.toContain("#tab=");
  });

  it("survives on an override too — the tab belongs to the link, not the page", () => {
    const t = resolveLink(withTab, SITE, "/sites/X/SitePages/FA.aspx");
    if (t.state !== "missing") expect(t.url).toBe("/sites/X/SitePages/FA.aspx#tab=reconciliation");
  });

  it("no card writes a tab now, so no two rows can collide on one page", () => {
    expect(allLinks().filter((l) => l.tab !== undefined)).toEqual([]);
  });
});

describe("tabFromHash", () => {
  it("reads the slug", () => {
    expect(tabFromHash("#tab=reconciliation")).toBe("reconciliation");
  });

  it("reads it from among other hash parameters", () => {
    expect(tabFromHash("#foo=1&tab=structure")).toBe("structure");
  });

  it("lowercases and trims, so a hand-typed hash still works", () => {
    expect(tabFromHash("#tab=Structure")).toBe("structure");
    expect(tabFromHash("#tab=%20structure%20")).toBe("structure");
  });

  it("returns BLANK for anything unrecognised — the caller opens its default tab", () => {
    // A stray hash must never leave a tabbed page showing nothing.
    expect(tabFromHash("")).toBe("");
    expect(tabFromHash("#")).toBe("");
    expect(tabFromHash("#other=1")).toBe("");
    expect(tabFromHash(undefined as unknown as string)).toBe("");
  });

  it("returns blank rather than throwing on a malformed escape", () => {
    expect(tabFromHash("#tab=%E0%A4%A")).toBe("");
  });

  it("round-trips every tab a link declares", () => {
    for (const l of allLinks()) {
      if (!l.tab) continue;
      const t = resolveLink(l, SITE);
      if (t.state === "missing") continue;
      expect(tabFromHash(`#${t.url.split("#")[1]}`)).toBe(l.tab);
    }
  });
});

describe("suggestedPageName", () => {
  it("suggests a name that matches the row's own label", () => {
    expect(suggestedPageName(linkFor("audit"))).toBe("CRS-Audit-Log.aspx");
    expect(suggestedPageName(linkFor("groups"))).toBe("Group-Management.aspx");
  });

  it("names the PAGE, not the row label, when they differ", () => {
    // Folder Management lives on a page called Folder Administration. Suggesting
    // "Folder-Management.aspx" would have an admin create a page the pattern can never match, leaving
    // the row dead after they did exactly as they were told.
    expect(suggestedPageName(linkFor("folders"))).toBe("Folder-Administration.aspx");
  });

  it("suggests a name the pattern will actually match — otherwise the advice is a dead end", () => {
    for (const l of allLinks()) {
      const suggested = suggestedPageName(l);
      expect(resolveLink(l, [page(suggested)]).state).not.toBe("missing");
    }
  });
});

/* ⚠ REMOVED 2026-08-22, and pinned so it cannot drift back in unnoticed. Bulk Upload was listed here
   as an ADMIN tool; the client handed it to uploaders, and this page is admin-only — so a card here
   would point administrators at a screen that is no longer theirs while staying invisible to everyone
   who now uses it. */
describe("Bulk Upload is not an admin card", () => {
  it("has no bulk card, and no bulk link to resolve", () => {
    expect(CARDS.filter((c) => c.key === "bulk")).toEqual([]);
    expect(allLinks().filter((l) => l.key === "bulk")).toEqual([]);
  });
});

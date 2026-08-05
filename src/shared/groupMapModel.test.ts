import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  roleFromGroupName,
  suggestGroupName,
  suffixForRole,
  ROLE_SUFFIXES,
  LIBRARY_ENTRY_ROLE,
  STAGING_FACING_ROLES,
  GroupMapWriteRow,
  GroupMapDraft,
  PERSONAS,
  PERSONA_FAMILIES,
  SELECTABLE_ROLES,
  personaByKey,
  GroupMapRole,
  normalizeScope,
  isForbiddenPageTarget,
} from "./groupMapModel";

describe("personas", () => {
  it("never invents a bundle group — every persona is a set of atomic roles", () => {
    const atomic = ["MEMBER", "UPL", "APR", "DEL", "DELS", "SEGVIEW", "GLOBAL", "HC"];
    for (const p of PERSONAS) {
      expect(p.roles.length).toBeGreaterThan(0);
      for (const r of p.roles) expect(atomic).toContain(r);
    }
  });

  it("gives every persona a distinct key", () => {
    const keys = PERSONAS.map((p) => p.key);
    expect(new Set(keys).size).toBe(keys.length);
  });

  it("gives each family a distinct role set — else its numbered groups collapse", () => {
    // Within a family, two identical role sets would be indistinguishable once
    // provisioned: the 2x2 of Head-of bundles collapses if approve+upload equals
    // approve+upload+delete. ACROSS families they are MEANT to repeat — Head of
    // Department and Head of Unit differ only by scope.
    for (const fam of PERSONA_FAMILIES) {
      const sets = PERSONAS.filter((p) => p.family === fam)
        .map((p) => p.roles.slice().sort().join("+"));
      expect(new Set(sets).size).toBe(sets.length);
    }
  });

  it("gives every Head-of persona the approve role — that is what makes it a Head-of", () => {
    for (const p of PERSONAS) {
      if (p.family.indexOf("Head of") === 0) expect(p.roles).toContain("APR");
    }
  });

  it("matches Head of Department and Head of Unit role-for-role, differing only in scope", () => {
    // The client's document lists them as eight groups with identical wording. If
    // these ever diverge, one family has silently gained or lost a capability.
    const hod = PERSONAS.filter((p) => p.family === "Head of Department");
    const hou = PERSONAS.filter((p) => p.family === "Head of Unit");
    expect(hod.length).toBe(4);
    expect(hou.length).toBe(4);
    for (let i = 0; i < hod.length; i++) {
      expect(hou[i].roles).toEqual(hod[i].roles);
      expect(hod[i].scope).toBe("department");
      expect(hou[i].scope).toBe("unit");
    }
  });

  it("scopes every persona below C-Level and Head of Department to the unit", () => {
    // A PIC or an employee fanned across a department would be a silent
    // cross-unit grant — the exact leak the isolation model exists to prevent.
    // C-Level is scoped ABOVE the department by design (one segment, or all of
    // them) and is the only family allowed to be.
    for (const p of PERSONAS) {
      if (p.family === "Head of Department" || p.family === "C-Level") continue;
      expect(p.scope).toBe("unit");
    }
  });

  it("offers GLOBAL as the only C-Level role, and keeps retired SEGVIEW out", () => {
    // SEGVIEW was retired on 2026-08-04, the day after it was added: the client
    // settled on a single global C-level rather than one per segment. It stays in
    // the type so roleFromGroupName still recognises the suffix (see below), but it
    // must never become offerable again.
    expect(SELECTABLE_ROLES).toContain("GLOBAL");
    expect(SELECTABLE_ROLES).not.toContain("SEGVIEW");
    expect(personaByKey("clevel_global")?.unavailable).toBeUndefined();
    expect(personaByKey("clevel_segment")).toBeUndefined();
    expect(PERSONAS.filter((p) => p.family === "C-Level").length).toBe(1);
  });

  it("still parses a retired _SEGVIEW group rather than letting it fall through to MEMBER", () => {
    // The reason SEGVIEW was not deleted outright. A group created during the one
    // day the role existed would otherwise parse as MEMBER and be granted Read at
    // whatever tier its row sits on. Recognised-but-inert fails safe; mis-parsed
    // grants access nobody asked for.
    expect(roleFromGroupName("DMS_MHO_SEGVIEW")).toBe("SEGVIEW");
  });

  it("gives the C-Level personas no power beyond reading", () => {
    for (const p of PERSONAS.filter((x) => x.family === "C-Level")) {
      expect(p.roles).not.toContain("UPL");
      expect(p.roles).not.toContain("APR");
      expect(p.roles).not.toContain("DEL");
      expect(p.roles).not.toContain("DELS");
    }
  });

  it("keeps PIC off Documents by omitting the base group", () => {
    // A PIC uploads and sees the unit's Staging files, which is what UPL grants. Reading
    // the approved archive is a SEPARATE decision, made by adding them to the unit's base
    // group — the SDG Employee role. Bundling MEMBER in here made every PIC a Documents
    // reader by default: the wrong default for a permission, and not what the client's
    // "can see the files in the unit" line meant.
    expect(personaByKey("pic1")?.roles).toEqual(["UPL"]);
  });

  it("offers exactly one PIC upload persona, plus the unavailable HC one", () => {
    // PIC 1 and PIC 3 collapsed once "own files only" was dropped as unachievable and
    // MEMBER came out of PIC 1: both are UPL alone. Two identical entries in a picker are
    // a trap, not a choice.
    const pics = PERSONAS.filter((p) => p.family === "PIC");
    const selectable = pics.filter((p) => !p.unavailable);
    expect(selectable.length).toBe(1);
    expect(selectable[0].roles).toEqual(["UPL"]);
    expect(personaByKey("pic3")).toBeUndefined();
  });

  it("lists PIC 2 but marks it unavailable rather than hiding it", () => {
    // Silently dropping one of the client's own numbered groups reads as an
    // oversight, and an admin would go hunting for it.
    const pic2 = personaByKey("pic2");
    expect(pic2).toBeDefined();
    expect(pic2?.unavailable).toBeTruthy();
    expect(pic2?.roles).toEqual(["HC"]);
  });

  it("keeps HC out of every provisionable persona and out of the role picker", () => {
    for (const p of PERSONAS) {
      if (p.unavailable) continue; // pic2 IS HC by definition, and is blocked
      expect(p.roles).not.toContain("HC");
    }
    expect(SELECTABLE_ROLES).not.toContain("HC");
  });

  it("covers all four families", () => {
    for (const fam of PERSONA_FAMILIES) {
      expect(PERSONAS.filter((p) => p.family === fam).length).toBeGreaterThan(0);
    }
  });

  it("returns undefined for an unknown persona key rather than throwing", () => {
    expect(personaByKey("nope")).toBeUndefined();
  });
});

describe("roleFromGroupName — DEL and HC suffixes", () => {
  it("derives DEL from a _DEL suffix", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_DEL")).toBe("DEL");
  });

  it("derives HC from an _HC suffix, so a Phase 2 group is not misread as MEMBER", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_HC")).toBe("HC");
  });

  it("does not mistake a unit whose name merely ends in those letters", () => {
    // No underscore before the suffix — this is a unit called "MODEL", not a DEL group.
    expect(roleFromGroupName("DMS_GHO_GF_MODEL")).toBe("MEMBER");
  });
});

describe("suggestGroupName — new role suffixes", () => {
  it("suffixes DEL", () => {
    expect(suggestGroupName("GHO", ["Group Finance"], "DEL")).toBe(
      "GHO_Group Finance_DELETER_DOCUMENTS",
    );
  });

  it("leaves MEMBER unsuffixed — the base group is identified by having no suffix", () => {
    expect(suggestGroupName("GHO", ["Group Finance"], "MEMBER")).toBe("GHO_Group Finance");
  });

  it("round-trips: a suggested name parses back to the role it was built for", () => {
    for (const role of SELECTABLE_ROLES) {
      if (role === "GLOBAL") continue; // fixed name, no tier
      expect(roleFromGroupName(suggestGroupName("GHO", ["GF", "CORU"], role))).toBe(role);
    }
  });
});

describe("roleFromGroupName", () => {
  it("derives APR from an _APR suffix", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_APR")).toBe("APR");
  });

  it("derives UPL from an _UPL suffix", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_UPL")).toBe("UPL");
  });

  it("defaults a base (no-suffix) group to MEMBER", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU")).toBe("MEMBER");
  });

  it("matches the suffix case-insensitively and trims", () => {
    expect(roleFromGroupName("  dms_gho_gf_coru_apr  ")).toBe("APR");
    expect(roleFromGroupName("DMS_GHO_GF_CORU_Upl")).toBe("UPL");
  });

  it("never auto-derives GLOBAL and is safe on empty/undefined", () => {
    expect(roleFromGroupName("DMS_GLOBAL_UPLOADERS")).toBe("MEMBER");
    expect(roleFromGroupName("")).toBe("MEMBER");
    expect(roleFromGroupName(undefined as unknown as string)).toBe("MEMBER");
  });

  it("derives DEL from a _DEL suffix", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_DEL")).toBe("DEL");
  });

  // The regression this case exists for: "_DEL" is a prefix of "_DELS", so a
  // naive suffix order parses every Staging-delete group as a DOCUMENTS-delete
  // role. The library rule would then drop it from Staging and grant it on
  // Documents — a head of unit silently handed delete over approved documents
  // while losing it on the pending files they were provisioned for.
  it("derives DELS from a _DELS suffix, and does NOT fall through to DEL", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_DELS")).toBe("DELS");
    expect(roleFromGroupName("  dms_gho_gf_coru_dels  ")).toBe("DELS");
  });
});

// Long-form suffixes, the client's convention from 2026-08-04.
//
// Every case here is a prefix collision: "_UPL" is a prefix of "_UPLOADER", "_APR" of
// "_APPROVER", "_DEL"/"_DELS" of both "_DELETER_*" forms. A name that fails to match any
// suffix reads as MEMBER, which LIBRARY_ROLES puts on Documents — so the wrong answer is
// not "unknown role", it is a silent grant on the wrong library.
describe("roleFromGroupName — long-form suffixes (prefix-less convention)", () => {
  it("derives UPL from _UPLOADER, not via the shorter _UPL", () => {
    expect(roleFromGroupName("GHO_GF_CORU_UPLOADER")).toBe("UPL");
  });

  it("derives APR from _APPROVER", () => {
    expect(roleFromGroupName("GHO_GF_CORU_APPROVER")).toBe("APR");
  });

  it("keeps the two delete roles apart in long form, per library", () => {
    expect(roleFromGroupName("GHO_GF_CORU_DELETER_STAGING")).toBe("DELS");
    expect(roleFromGroupName("GHO_GF_CORU_DELETER_DOCUMENTS")).toBe("DEL");
  });

  it("still accepts the short forms, so groups already on the site keep working", () => {
    expect(roleFromGroupName("GHO_GF_CORU_UPL")).toBe("UPL");
    expect(roleFromGroupName("DMS_GHO_GF_CORU_UPL")).toBe("UPL");
    expect(roleFromGroupName("GHO_GF_CORU_APR")).toBe("APR");
    expect(roleFromGroupName("GHO_GF_CORU_DELS")).toBe("DELS");
    expect(roleFromGroupName("GHO_GF_CORU_DEL")).toBe("DEL");
  });

  it("is case-insensitive and trims, in long form too", () => {
    expect(roleFromGroupName("  gho_gf_coru_uploader  ")).toBe("UPL");
    expect(roleFromGroupName("GHO_GF_CORU_Approver")).toBe("APR");
  });

  it("a prefix-less base group is still MEMBER", () => {
    expect(roleFromGroupName("GHO_GF_CORU")).toBe("MEMBER");
  });

  // Guards the property rather than the declaration order: whatever order the table is
  // written in, the LONGEST matching suffix must win. Sorting by length is what makes this
  // true, so this test fails if someone replaces the sort with a hand-ordered list.
  it("always resolves to the longest matching suffix", () => {
    for (const { suffix, role } of ROLE_SUFFIXES) {
      const longer = ROLE_SUFFIXES.filter(
        (o) => o.suffix.length > suffix.length && o.suffix.endsWith(suffix),
      );
      for (const l of longer) {
        expect(roleFromGroupName(`GHO_GF_CORU${l.suffix}`)).toBe(l.role);
      }
      expect(roleFromGroupName(`GHO_GF_CORU${suffix}`)).toBe(role);
    }
  });
});

// ENTRY = "may open this library". Its safety rests on two properties, both asserted here
// because both are invisible at the call site.
describe("ENTRY — the library-entry role", () => {
  it("is never offered by the folder-scope picker", () => {
    // A folder-scope ENTRY row would be meaningless; worse, an admin picking it from the role
    // list would believe they had granted something.
    expect(SELECTABLE_ROLES.indexOf(LIBRARY_ENTRY_ROLE)).toBe(-1);
  });

  it("is not derived from any group name suffix", () => {
    // Nothing should turn into ENTRY by being named a certain way — the tab writes it
    // explicitly or it does not exist.
    for (const { role } of ROLE_SUFFIXES) expect(role).not.toBe(LIBRARY_ENTRY_ROLE);
    expect(roleFromGroupName("GHO_GF_CORU_ENTRY")).toBe("MEMBER");
  });

  it("builds a library row with no term, keeping the target", () => {
    const row = buildGroupMapRow({
      groupId: "27",
      groupName: "GHO_GF_CORU_UPLOADER",
      role: LIBRARY_ENTRY_ROLE,
      scope: "Library",
      target: "Staging",
    });
    expect(row.Scope).toBe("Library");
    expect(row.Target).toBe("Staging");
    // A library grant applies to the whole list; a term here would imply a folder it does
    // not touch, and reconciliation would read the row as folder-scoped.
    expect(row.UnitTermGuid).toBe("");
    expect(row.Segment).toBe("");
    expect(row.Role).toBe("ENTRY");
  });
});

// Scope is a Choice column, and single- vs multi-select is invisible in the SharePoint form —
// both render the value as a pill. Multi-select returns an array, so this is what stops a
// checkbox in the column settings throwing "raw.trim is not a function" at runtime.
describe("normalizeScope — Choice column shapes", () => {
  it("reads a single-select string", () => {
    expect(normalizeScope("Library")).toBe("Library");
    expect(normalizeScope("site")).toBe("Site");
    expect(normalizeScope("  PAGE  ")).toBe("Page");
  });

  it("reads a multi-select array by taking its first entry", () => {
    expect(normalizeScope(["Library"])).toBe("Library");
    expect(normalizeScope(["Site"])).toBe("Site");
  });

  it("falls back to Folder on every empty shape, never throwing", () => {
    expect(normalizeScope(undefined)).toBe("Folder");
    expect(normalizeScope("")).toBe("Folder");
    expect(normalizeScope([])).toBe("Folder");
    // An emptied multi-value Choice returns null, not [] — verified live 2026-07-30.
    expect(normalizeScope(null as unknown as string)).toBe("Folder");
  });

  it("treats an unrecognised value as Folder rather than guessing", () => {
    // A typo'd choice must not become a Library grant by accident.
    expect(normalizeScope("Librray")).toBe("Folder");
  });
});

describe("STAGING_FACING_ROLES", () => {
  it("lists exactly the roles that act on Staging", () => {
    expect(STAGING_FACING_ROLES).toEqual(["UPL", "APR", "DELS"]);
  });

  // The isolation rule: a viewer group reaching Staging reads other people's unapproved
  // drafts. These two are the ones an admin is most likely to think belong here.
  it("excludes the view-only roles", () => {
    expect(STAGING_FACING_ROLES.indexOf("MEMBER")).toBe(-1);
    expect(STAGING_FACING_ROLES.indexOf("GLOBAL")).toBe(-1);
  });
});

describe("suffixForRole", () => {
  it("writes the long form for new names", () => {
    expect(suffixForRole("UPL")).toBe("_UPLOADER");
    expect(suffixForRole("APR")).toBe("_APPROVER");
    expect(suffixForRole("DELS")).toBe("_DELETER_STAGING");
    expect(suffixForRole("DEL")).toBe("_DELETER_DOCUMENTS");
  });

  it("returns nothing for the roles that are not named by suffix", () => {
    expect(suffixForRole("MEMBER")).toBe("");
    expect(suffixForRole("GLOBAL")).toBe("");
    expect(suffixForRole("")).toBe("");
  });

  it("round-trips through roleFromGroupName for every selectable role", () => {
    for (const role of SELECTABLE_ROLES) {
      if (role === "MEMBER" || role === "GLOBAL") continue; // no suffix by design
      expect(roleFromGroupName(`GHO_GF_CORU${suffixForRole(role)}`)).toBe(role);
    }
  });
});

describe("PERSONAS — Staging delete belongs to the upload groups only", () => {
  // The client's document is explicit that Head-of groups 2 and 3 cannot delete
  // on Staging, and equally explicit that a PIC needs the head of unit's approval
  // first. Both facts are carried by which personas hold DELS, so a wrong entry
  // here is a silent permission grant rather than a visible bug.
  const rolesOf = (key: string): string[] => personaByKey(key)?.roles ?? [];

  it.each(["hod1", "hod4", "hou1", "hou4"])("gives %s Staging delete", (key) => {
    expect(rolesOf(key)).toContain("DELS");
    expect(rolesOf(key)).toContain("UPL");
  });

  it.each(["hod2", "hod3", "hou2", "hou3"])("withholds Staging delete from %s", (key) => {
    expect(rolesOf(key)).not.toContain("DELS");
    expect(rolesOf(key)).not.toContain("UPL");
  });

  it("gives the PIC upload without Staging delete", () => {
    // The client's rule: "if any PIC wants to delete they need approval from the head of
    // unit first". Enforced by withholding DELS, not by trusting the process.
    expect(rolesOf("pic1")).toContain("UPL");
    expect(rolesOf("pic1")).not.toContain("DELS");
  });

  it("keeps Documents delete and Staging delete distinct", () => {
    // Group 3 deletes in Documents only; group 4 in both. If these ever coincide
    // by accident, one role is standing in for the other and the library rule
    // stops meaning anything.
    expect(rolesOf("hod3")).toContain("DEL");
    expect(rolesOf("hod3")).not.toContain("DELS");
    expect(rolesOf("hod4")).toEqual(expect.arrayContaining(["DEL", "DELS"]));
  });
});

describe("buildGroupMapRow", () => {
  it("builds a trimmed row for a unit-tier UPL mapping", () => {
    const draft: GroupMapDraft = {
      groupId: "  00000000-aaaa-bbbb-cccc-000000000001 ",
      groupName: " DMS_LRC_GCO_UPL ",
      role: "UPL",
      segmentGuid: " efa87c6a-9536-4f7c-910f-011bf7413b80 ",
      tierGuid: " 479a74d8-4126-40f0-9b1f-3374de74152b ",
    };
    expect(buildGroupMapRow(draft)).toEqual({
      GroupId: "00000000-aaaa-bbbb-cccc-000000000001",
      GroupName: "DMS_LRC_GCO_UPL",
      Segment: "efa87c6a-9536-4f7c-910f-011bf7413b80",
      UnitTermGuid: "479a74d8-4126-40f0-9b1f-3374de74152b",
      Role: "UPL",
      // A draft that names no scope is a folder row — the historical meaning, and the
      // only shape this list held before 2026-08-04.
      Scope: "Folder",
      Target: "",
    });
  });

  it("leaves Segment and UnitTermGuid empty for a GLOBAL row", () => {
    const draft: GroupMapDraft = {
      groupId: "g-1",
      groupName: "DMS_GLOBAL_UPLOADERS",
      role: "GLOBAL",
      segmentGuid: "should-be-ignored",
      tierGuid: "should-be-ignored",
    };
    expect(buildGroupMapRow(draft)).toEqual({
      GroupId: "g-1",
      GroupName: "DMS_GLOBAL_UPLOADERS",
      Segment: "",
      UnitTermGuid: "",
      Role: "GLOBAL",
      Scope: "Folder",
      Target: "",
    });
  });

  it("supports a segment-tier row where the tier GUID equals the segment GUID", () => {
    const row = buildGroupMapRow({
      groupId: "g-2",
      groupName: "DMS_GHO",
      role: "MEMBER",
      segmentGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
      tierGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
    });
    expect(row.Segment).toBe(row.UnitTermGuid);
  });
});

describe("isDuplicateRow", () => {
  const existing: GroupMapWriteRow[] = [
    { GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-gco", Role: "UPL", Scope: "Folder", Target: "" },
  ];

  it("matches on GroupId + UnitTermGuid + Role, case/space-insensitive", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: " G-1 ", GroupName: "A", Segment: "s", UnitTermGuid: "T-GCO", Role: "UPL", Scope: "Folder", Target: "",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(true);
  });

  it("allows the same group with a different role", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-gco", Role: "APR", Scope: "Folder", Target: "",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(false);
  });

  it("allows the same group mapping a different tier", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-risk", Role: "UPL", Scope: "Folder", Target: "",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(false);
  });

  it("treats two Library rows for different libraries as distinct", () => {
    // Both are termless, so without Scope+Target in the key they collide and the second
    // is rejected as a duplicate — telling the admin the mapping already exists while the
    // library they actually wanted stays unmapped.
    const docs: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "", UnitTermGuid: "", Role: "MEMBER", Scope: "Library", Target: "Documents",
    };
    expect(isDuplicateRow([docs], { ...docs, Target: "Staging" })).toBe(false);
    expect(isDuplicateRow([docs], { ...docs, Target: " documents " })).toBe(true);
  });

  it("treats a Site row and a Folder row for one group as distinct", () => {
    const site: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "", UnitTermGuid: "", Role: "MEMBER", Scope: "Site", Target: "",
    };
    expect(isDuplicateRow(existing, site)).toBe(false);
  });
});

describe("scope", () => {
  it("reads a blank, missing or unrecognised Scope as Folder", () => {
    // The migration guarantee: every row that exists today predates the column and all of
    // them are folder rows. Any other default would silently stop applying an existing
    // site's folder permissions the moment the column is added, while still reporting a
    // clean run.
    expect(normalizeScope(undefined)).toBe("Folder");
    expect(normalizeScope("")).toBe("Folder");
    expect(normalizeScope("   ")).toBe("Folder");
    expect(normalizeScope("nonsense")).toBe("Folder");
  });

  it("accepts the four scopes case-insensitively", () => {
    expect(normalizeScope("site")).toBe("Site");
    expect(normalizeScope(" LIBRARY ")).toBe("Library");
    expect(normalizeScope("Page")).toBe("Page");
    expect(normalizeScope("folder")).toBe("Folder");
  });

  it("strips terms and target from the scopes they are meaningless for", () => {
    // A leftover segment on a Site row would look like a folder row to reconciliation AND
    // to the orphan-repair pass, which could then re-point it.
    const site = buildGroupMapRow({
      groupId: "g", groupName: "n", role: "MEMBER", scope: "Site",
      segmentGuid: "seg", tierGuid: "unit", target: "Documents",
    });
    expect(site.Segment).toBe("");
    expect(site.UnitTermGuid).toBe("");
    expect(site.Target).toBe("");
  });

  it("keeps the trimmed target on Library rows", () => {
    const lib = buildGroupMapRow({ groupId: "g", groupName: "n", role: "MEMBER", scope: "Library", target: " Documents " });
    expect(lib.Target).toBe("Documents");
    expect(lib.Segment).toBe("");
  });

  it("requires a target for Library and Page, and terms only for Folder", () => {
    const base = { groupId: "g", groupName: "n", role: "MEMBER" as GroupMapRole };
    expect(validateDraft({ ...base, scope: "Site" })).toEqual([]);
    expect(validateDraft({ ...base, scope: "Library" }).length).toBe(1);
    expect(validateDraft({ ...base, scope: "Library", target: "Documents" })).toEqual([]);
    expect(validateDraft({ ...base, scope: "Folder" }).length).toBe(2);
  });

  it("refuses a Page row aimed at the home page", () => {
    // Granting site entry and then blocking the only page a user can navigate to makes
    // the site-entry layer worthless. Caught at write time, because nothing in this
    // system revokes a grant once it exists.
    const base = { groupId: "g", groupName: "n", role: "MEMBER" as GroupMapRole, scope: "Page" as const };
    expect(validateDraft({ ...base, target: "Home.aspx" }).length).toBe(1);
    expect(validateDraft({ ...base, target: "SitePages/home.aspx" }).length).toBe(1);
    expect(validateDraft({ ...base, target: "Upload.aspx" })).toEqual([]);
    expect(isForbiddenPageTarget("/home.aspx")).toBe(true);
    expect(isForbiddenPageTarget("Upload.aspx")).toBe(false);
  });
});

describe("validateDraft", () => {
  it("returns [] for a complete non-global draft", () => {
    expect(
      validateDraft({ groupId: "g", groupName: "n", role: "MEMBER", segmentGuid: "s", tierGuid: "t" }),
    ).toEqual([]);
  });

  it("requires a group and a role", () => {
    const errs = validateDraft({ groupId: "  ", groupName: "", role: undefined as unknown as "MEMBER" });
    expect(errs).toContain("Select a group.");
    expect(errs).toContain("Select a role.");
  });

  it("requires segment + tier unless GLOBAL", () => {
    expect(validateDraft({ groupId: "g", groupName: "n", role: "UPL" })).toEqual([
      "Select a segment.",
      "Select a tier.",
    ]);
  });

  it("does not require segment/tier for GLOBAL", () => {
    expect(validateDraft({ groupId: "g", groupName: "n", role: "GLOBAL" })).toEqual([]);
  });
});

// Convention as of 2026-08-04 (client request): NO "DMS_" prefix, long-form role suffix.
// e.g. GHO_GF_CORU_UPLOADER. See the 2026-08-04 amendment in
// docs/superpowers/specs/2026-07-22-role-from-group-name-suffix.md.
describe("suggestGroupName", () => {
  it("joins segment + tier labels + long-form role suffix, with NO prefix", () => {
    expect(
      suggestGroupName("Group Head Office", ["Group Finance", "Corporate Reporting"], "UPL"),
    ).toBe("Group Head Office_Group Finance_Corporate Reporting_UPLOADER");
  });

  it("builds the exact name the client asked for", () => {
    expect(suggestGroupName("GHO", ["GF", "CORU"], "UPL")).toBe("GHO_GF_CORU_UPLOADER");
  });

  it("APR suffix works the same way", () => {
    expect(suggestGroupName("Minamas Head Office", ["Finance"], "APR")).toBe(
      "Minamas Head Office_Finance_APPROVER",
    );
  });

  it("MEMBER gets no suffix (base/viewer group)", () => {
    expect(suggestGroupName("Group Head Office", ["Group Finance"], "MEMBER")).toBe(
      "Group Head Office_Group Finance",
    );
  });

  it("GLOBAL is a fixed name regardless of labels", () => {
    expect(suggestGroupName("anything", ["x", "y"], "GLOBAL")).toBe("GLOBAL");
  });

  it("skips empty labels and trims the rest", () => {
    expect(suggestGroupName(" GHO ", ["", "  Unit A "], "APR")).toBe("GHO_Unit A_APPROVER");
  });

  it("empty role and empty labels degrade gracefully", () => {
    expect(suggestGroupName("GHO", [], "")).toBe("GHO");
    expect(suggestGroupName("", [], "")).toBe("");
  });

  it("never emits a trailing underscore while the admin is still picking a role", () => {
    expect(suggestGroupName("GHO", ["GF"], "")).toBe("GHO_GF");
  });
});

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
  personaTouchesStaging,
  GroupMapRole,
  normalizeScope,
  isForbiddenPageTarget,
  normalizeRoleValue,
  isSiteEntryGroupTitle,
  findSiteEntryGroup,
  siteEntryGroupTitle,
  validateGroupName,
} from "./groupMapModel";
import { setSiteEntryName } from "./naming";

describe("personas", () => {
  it("never invents a bundle group — every persona is a set of atomic roles", () => {
    // 2026-08-17: DEPTVIEW, DELSHC and MEMBERHC added; APRHC removed (retired).
    const atomic = ["MEMBER", "UPL", "APR", "DEL", "DELS", "SEGVIEW", "DEPTVIEW", "GLOBAL", "SHARE",
                    "UPLHC", "DELSHC", "MEMBERHC"];
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

  it("gives approve to Head of Unit and to nobody else", () => {
    // Reversed on 2026-08-07: approval used to belong to every Head-of family. The client
    // moved it to Head of Unit alone, so an APR appearing on any other persona means the
    // department tier has quietly regained the power to approve its units' documents.
    const approvers = PERSONAS.filter((p) => p.roles.indexOf("APR") !== -1);
    expect(approvers.map((p) => p.key)).toEqual(["hou"]);
  });

  it("keeps Head of Department to delete, department-scoped", () => {
    const hod = PERSONAS.filter((p) => p.family === "Head of Department");
    expect(hod.length).toBe(1);
    expect(hod[0].scope).toBe("department");
    // MEMBER dropped 2026-08-09: DEL maps to "CRS Delete" = Read + Delete Items, so the read
    // was always there. The second membership only made it look as though it were not.
    // SHARE added 2026-08-15: a HoD shares without approval.
    // 2026-08-17: DEL and SHARE both dropped — "HOD no need deletion power, he only view" and
    // "Head of Department do not have share functionality that is HOU". DEPTVIEW, not MEMBER: MEMBER
    // is the SDG Employee role and is absent from HC Documents, so reusing it would either lose this
    // persona's HC read or hand every SDG Employee HC read.
    expect(hod[0].roles).toEqual(["DEPTVIEW"]);
    expect(hod[0].roles).not.toContain("DEL");
    expect(hod[0].roles).not.toContain("SHARE");
    // DEL is the DOCUMENTS delete. DELS would put a department head on Staging, which
    // LIBRARY_ROLES routes there — and HoD has no Staging access at all.
    expect(hod[0].roles).not.toContain("DELS");
    expect(hod[0].roles).not.toContain("UPL");
  });

  it("keeps Head of Unit to approve + delete pending, unit-scoped", () => {
    // TWO since 2026-08-15: the ordinary Head of Unit, and the Highly Confidential one. Separate
    // personas rather than one with a flag, because the client's whole reason for an HC approver
    // group was that "not all approver can see HC files" — a variant shipped alongside the plain one
    // would defeat that on the first careless click.
    const hou = PERSONAS.filter((p) => p.family === "Head of Unit");
    // ONE again since 2026-08-17: hou_hc retired. The client reversed their own rule — "any approver
    // which is HOU can see Highly Confidential files as well" — and with APR on both HC rows of
    // LIBRARY_ROLES the HC variant became an exact duplicate.
    expect(hou.map((p) => p.key)).toEqual(["hou"]);
    for (const p of hou) expect(p.scope).toBe("unit");
    // 2026-08-09: MEMBER dropped (APR now carries Documents read by itself) and DELS added —
    // the client's "they should have the power to delete" pending or rejected files.
    // 2026-08-15: UPL, DEL and SHARE added. DEL is no longer the department head's alone —
    // a Head of Unit needs it to CARRY OUT an approved deletion request, because an approver can
    // only approve what they can perform. The two deletes are still kept apart by LIBRARY_ROLES,
    // which is where that separation actually lives; the model no longer asserts it here.
    // 2026-08-17: UPLHC REPLACES UPL ("any type of HOU can upload to HC"). A superset, so carrying
    // both would be two rows granting the same thing on the same library.
    // DELSHC too, and NOT `DEL` on the HC approval library — see the DEL-holders test below for why
    // that distinction is load-bearing.
    expect(hou[0].roles).toEqual(["APR", "DELS", "DEL", "SHARE", "UPLHC", "DELSHC"]);
    // The Head of Unit now REACHES the HC libraries — through APR and DEL for reading, approving and
    // deleting, and through UPLHC for filing. HC clearance is a dedicated group only for the PIC and
    // the SDG Employee.
    expect(hou[0].roles).toContain("UPLHC");
    expect(hou[0].roles).not.toContain("UPL");
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

  it("offers both C-Level shapes — all segments, and one segment", () => {
    // SEGVIEW was retired on 2026-08-04 and UN-retired on 2026-08-07: the client's third
    // restatement asks for both "view the entire business segments" and "view its own
    // business segment only". The two differ only in reach — GLOBAL rows are termless and
    // reach everything, a SEGVIEW row carries a segment term and reaches that segment.
    expect(SELECTABLE_ROLES).toContain("GLOBAL");
    // DEL and SHARE joined both on 2026-08-15 — C-Level acts without approval.
    expect(personaByKey("clevel_global")?.roles).toEqual(["GLOBAL", "DEL", "SHARE"]);
    expect(personaByKey("clevel_segment")?.roles).toEqual(["SEGVIEW", "DEL", "SHARE"]);
    expect(personaByKey("clevel_global")?.unavailable).toBeUndefined();
    expect(personaByKey("clevel_segment")?.unavailable).toBeUndefined();
    expect(PERSONAS.filter((p) => p.family === "C-Level").length).toBe(2);
  });

  it("parses a _SEGVIEW group rather than letting it fall through to MEMBER", () => {
    // Why SEGVIEW was never deleted outright, and why that turned out to matter: groups
    // created during the one day the role first existed still parse correctly now that it
    // is live again. Mis-parsing one as MEMBER would grant Read at whatever tier its row
    // sits on — for a segment-tier row, an entire business segment.
    expect(roleFromGroupName("DMS_MHO_SEGVIEW")).toBe("SEGVIEW");
  });

  it("gives the C-Level personas delete and share, but never upload or approve", () => {
    // Rewritten 2026-08-15. C-Level is no longer view-only: the client confirmed they delete and
    // share without approval. What must NOT change is the approval library — UPL and APR would put
    // them among unapproved drafts, and DELS would let them delete other people's pending work.
    for (const p of PERSONAS.filter((x) => x.family === "C-Level")) {
      expect(p.roles).toContain("DEL");
      expect(p.roles).toContain("SHARE");
      expect(p.roles).not.toContain("UPL");
      expect(p.roles).not.toContain("APR");
      expect(p.roles).not.toContain("DELS");
    }
  });

  it("keeps PIC off Documents by omitting the base group", () => {
    // A PIC uploads and sees the unit's Staging files, which is what UPL grants. Reading
    // the approved archive is a SEPARATE decision, made by adding them to the unit's base
    // group — the SDG Employee role. Bundling MEMBER in here made every PIC a Documents
    // reader by default: the wrong default for a permission, and not what the client's
    // "can see the files in the unit" line meant.
    expect(personaByKey("pic")?.roles).toEqual(["UPL", "DELS"]);
  });

  it("offers two PIC personas — ordinary, and Highly Confidential cleared", () => {
    // REVERSED 2026-08-15, and worth stating precisely, because the old rule was right for the
    // reason it gave. "Confidentiality is metadata, not a permission" still holds for every level
    // EXCEPT the one routing to the HC library pair: that level is a different DESTINATION, and a
    // destination is a permission. So there is exactly one clearance variant, not one per level.
    //
    // The client's three-PIC document is still wrong for the reason it always was — pic2 and pic3
    // collapse into pic — so those keys stay absent.
    const pics = PERSONAS.filter((p) => p.family === "PIC");
    expect(pics.map((p) => p.key)).toEqual(["pic", "pic_hc"]);
    expect(pics[0].roles).toEqual(["UPL", "DELS"]);
    // DELS -> DELSHC 2026-08-17, and this closed a live leak rather than renaming anything: the
    // PLAIN PIC also holds DELS, so while DELS was on the HC approval library every plain PIC held
    // CRS Delete there.
    expect(pics[1].roles).toEqual(["UPLHC", "DELSHC"]);
    for (const p of pics) expect(p.unavailable).toBeUndefined();
    expect(personaByKey("pic2")).toBeUndefined();
    expect(personaByKey("pic3")).toBeUndefined();
  });

  it("never puts a plain uploader and an HC uploader on the same persona", () => {
    // UPLHC is a SUPERSET — it covers the normal approval library too — so carrying both is not
    // wider access, it is two group memberships to keep in step through a term rename. One person,
    // one group, as everywhere else since 2026-08-09.
    // APRHC retired 2026-08-17, so only the upload pair remains to check. UPLHC is a superset of UPL,
    // so a persona carrying both would hold two rows granting the same thing on the same library —
    // redundant, and two things to keep in step through a term rename.
    for (const p of PERSONAS) {
      if (p.roles.indexOf("UPLHC") !== -1) expect(p.roles).not.toContain("UPL");
    }
  });

  it("has retired the bare HC role in favour of UPLHC and APRHC", () => {
    // The old `HC` belonged to the superseded 2026-07-16 design, where HC was a secured subfolder
    // and a group was simply "cleared". With two libraries and two sides that word is ambiguous —
    // cleared to upload, or to approve? — so the role now says which. Asserted because a stray "HC"
    // row would match no library and grant nothing, silently.
    for (const p of PERSONAS) expect(p.roles).not.toContain("HC");
    expect(SELECTABLE_ROLES).not.toContain("HC");
    expect(SELECTABLE_ROLES).toContain("UPLHC");
    expect(SELECTABLE_ROLES).toContain("DELSHC");
    expect(SELECTABLE_ROLES).toContain("MEMBERHC");
    // APRHC retired 2026-08-17 — offering it would let an admin author a row granting nothing.
    expect(SELECTABLE_ROLES).not.toContain("APRHC");
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

  it("derives the HC roles from either spelling the client might type", () => {
    // Both accepted: the long form is what the client wrote, the short one is what an administrator
    // types when the name is already long at four tiers.
    expect(roleFromGroupName("GHO_GF_CORU_UPL_HIGHLY_CONFIDENTIAL")).toBe("UPLHC");
    // Both _APR_HC spellings now resolve to plain APR (2026-08-17). The suffixes stay MAPPED rather
    // than deleted, so a group somebody already created keeps working and means what it now means;
    // dropping them would make such a name fall through to MEMBER — a view-only base group.
    expect(roleFromGroupName("GHO_GF_CORU_APR_HIGHLY_CONFIDENTIAL")).toBe("APR");
    expect(roleFromGroupName("GHO_GF_CORU_UPL_HC")).toBe("UPLHC");
    expect(roleFromGroupName("GHO_GF_CORU_APR_HC")).toBe("APR");
    expect(roleFromGroupName("GHO_GF_CORU_DELS_HC")).toBe("DELSHC");
    expect(roleFromGroupName("GHO_GF_CORU_VIEWER_HC")).toBe("MEMBERHC");
    expect(roleFromGroupName("GHO_GF_DEPARTMENT_VIEWER")).toBe("DEPTVIEW");
  });

  it("tests the HC suffix BEFORE the plain one — the length sort is load-bearing", () => {
    // "..._UPL_HC" ends with "_HC" AND with "_UPL". Matching the shorter first would parse an HC
    // uploader as a plain one: HC clearance granted to nobody, and the list still reading correctly.
    expect(roleFromGroupName("GHO_GF_CORU_UPL_HC")).not.toBe("UPL");
    // _DELS_HC (8) must beat _DELS (5) and _HC (3); _VIEWER_HC (10) must beat _HC.
    expect(roleFromGroupName("GHO_GF_CORU_DELS_HC")).not.toBe("DELS");
    expect(roleFromGroupName("GHO_GF_CORU_VIEWER_HC")).not.toBe("MEMBER");
    // _APR_HC deliberately DOES resolve to APR now — see the previous case.
  });

  it("reads a legacy bare _HC group as an HC UPLOADER", () => {
    // Under the superseded design such a group was the unit's HC-cleared members — its uploaders.
    // Falling through to MEMBER instead would grant plain Read at whatever tier its row sits on.
    expect(roleFromGroupName("DMS_GHO_GF_CORU_HC")).toBe("UPLHC");
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
    // UPLHC and APRHC joined 2026-08-15: both face the normal approval library as well as the HC
    // one, so a persona carrying either belongs on the approval-library side of the toggle.
    // DELSHC replaced APRHC 2026-08-17. DEPTVIEW is absent for the same reason MEMBER and GLOBAL
    // are: a view role on an approval library reads other people's unapproved drafts.
    expect(STAGING_FACING_ROLES).toEqual(["UPL", "APR", "DELS", "UPLHC", "DELSHC"]);
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

describe("normalizeRoleValue — the Role COLUMN, not the group name", () => {
  // Live 2026-08-07: the only row on the test site had Role = "UPLOADER", because the GROUP
  // NAMES use long suffixes now and an admin filling the column in by hand matched them. It was
  // skipped silently, and the log said "no group-map groups for this unit" — which reads as a
  // missing row, sending the admin to look for one that is sitting right there.
  it("maps the long forms an admin will actually type", () => {
    expect(normalizeRoleValue("UPLOADER")).toBe("UPL");
    expect(normalizeRoleValue("APPROVER")).toBe("APR");
    expect(normalizeRoleValue("DELETER_DOCUMENTS")).toBe("DEL");
    expect(normalizeRoleValue("DELETER_STAGING")).toBe("DELS");
    expect(normalizeRoleValue("VIEWER")).toBe("MEMBER");
  });

  it("passes short codes through untouched", () => {
    for (const code of ["MEMBER", "UPL", "APR", "DEL", "DELS", "GLOBAL", "SEGVIEW", "ENTRY"]) {
      expect(normalizeRoleValue(code)).toBe(code);
    }
  });

  it("tolerates the casing, spacing and hyphens of a hand-typed cell", () => {
    expect(normalizeRoleValue("  uploader ")).toBe("UPL");
    expect(normalizeRoleValue("Deleter Documents")).toBe("DEL");
    expect(normalizeRoleValue("deleter-staging")).toBe("DELS");
  });

  it("never invents a role from an unrecognised value", () => {
    // Passing the value through unchanged means it still fails accepts() and is skipped. The
    // dangerous alternative is defaulting to something — MEMBER would grant Read on a folder
    // to a group the admin meant to give nothing.
    expect(normalizeRoleValue("SUPERVISOR")).toBe("SUPERVISOR");
    expect(normalizeRoleValue("")).toBe("");
    expect(normalizeRoleValue("   ")).toBe("");
    expect(normalizeRoleValue(undefined as unknown as string)).toBe("");
  });

  it("cannot re-map an already-valid short code via an alias", () => {
    // DEL must never become DELS or vice versa: they map to the same permission LEVEL and are
    // told apart only by LIBRARY_ROLES, so a swap crosses Staging and Documents.
    expect(normalizeRoleValue("DEL")).toBe("DEL");
    expect(normalizeRoleValue("DELS")).toBe("DELS");
  });
});

describe("PERSONAS — delete belongs to the Head of Department, in Documents only", () => {
  // Rewritten 2026-08-07. Delete used to ride with upload across four HoD and four HoU
  // bundles; the client's third restatement gives delete to the Head of Department and
  // says nothing about Staging delete at all. A wrong entry here is a silent permission
  // grant rather than a visible bug, which is why each half is asserted separately.
  const rolesOf = (key: string): string[] => personaByKey(key)?.roles ?? [];

  it("no longer gives the Head of Department any delete — view only since 2026-08-17", () => {
    // The client removed it: "HOD no need deletion power, he only view". Deletion authority is now
    // entirely the Head of Unit's, which is where the request workflow already put the performing
    // half of it. Asserted as the EXACT set: a leftover DEL row keeps granting delete, silently.
    expect(rolesOf("hod")).toEqual(["DEPTVIEW"]);
  });

  it("keeps DEL to the Head of Unit and C-Level — which is WHY it is not on an approval library", () => {
    /* This assertion exists because getting it wrong on 2026-08-17 nearly shipped a leak. DEL was
       briefly added to StagingHC to give a Head of Unit delete on pending HC work — on the belief
       that HoD losing DEL had made it HoU-exclusive. It had not: C-Level carries DEL too, so a
       C-Level would have gained read and delete on UNAPPROVED HC drafts, breaking the rule that
       keeps every view role off both approval libraries.

       DELSHC is the role that IS exclusive to the two unit personas, so that is what carries HC
       pending delete. Assert the exact holders, because the safety of an HC row depends on it. */
    const withDocDelete = PERSONAS.filter((p) => p.roles.indexOf("DEL") !== -1).map((p) => p.key);
    expect(withDocDelete.sort()).toEqual(["clevel_global", "clevel_segment", "hou"]);
    const withHcPendingDelete = PERSONAS.filter((p) => p.roles.indexOf("DELSHC") !== -1).map((p) => p.key);
    expect(withHcPendingDelete.sort()).toEqual(["hou", "pic_hc"]);
  });

  it("gives Staging delete to the Head of Unit, and to nobody else", () => {
    // Reversed 2026-08-09. DELS used to belong to NO persona — kept alive only so a
    // hand-authored row would still work. The client assigned it to the Head of Unit, whose
    // job it already is to act on pending work.
    //
    // "And to nobody else" is the half that still matters: DELS is the power to delete other
    // people's PENDING documents. Anyone else acquiring it is a silent grant, which is why
    // this asserts the exact set rather than just that HoU has it.
    // 2026-08-15: the PIC joins, correcting the model — they delete their OWN pending and rejected
    // files. No "own files only" rule is needed or possible: Draft Item Security already hides a
    // peer's pending work, so what they can delete is exactly what they can see.
    //
    // The exact set still matters. DELS is the power to delete pending work, so anyone ELSE
    // acquiring it is a silent grant.
    // Later the same day: the HC variants join, for the same reasons as their plain counterparts.
    // The exact set is still what matters.
    const withStagingDelete = PERSONAS.filter((p) => p.roles.indexOf("DELS") !== -1);
    // 2026-08-17: hou_hc retired, and pic_hc moved to DELSHC — so DELS is now the plain pair only.
    expect(withStagingDelete.map((p) => p.key)).toEqual(["hou", "pic"]);
    // DELSHC is its own role precisely so the plain PIC cannot reach the HC approval library.
    const withHcStagingDelete = PERSONAS.filter((p) => p.roles.indexOf("DELSHC") !== -1);
    // The two UNIT personas and nobody else — which is exactly what makes DELSHC safe to place on
    // the HC approval library where DEL is not.
    expect(withHcStagingDelete.map((p) => p.key)).toEqual(["hou", "pic_hc"]);
  });

  it("keeps the PIC to upload alone", () => {
    // Still UPL only, even though a PIC now reads Documents too. That read comes from
    // LIBRARY_ROLES listing UPL under Documents — NOT from adding MEMBER here. If this ever
    // grows a second role, the "one group per person" property has been lost.
    expect(rolesOf("pic")).toEqual(["UPL", "DELS"]);
  });

  it("leaves SDG Employee as the only persona that is purely MEMBER", () => {
    // Since 2026-08-09 every other persona carries its own Documents read, so the base group
    // means exactly what its name says: someone who views and does nothing else. A second
    // MEMBER-only persona would be an alias for it, and two names for one thing is how an
    // admin ends up assigning the wrong one.
    const memberOnly = PERSONAS.filter(
      (p) => p.roles.length === 1 && p.roles[0] === "MEMBER",
    );
    expect(memberOnly.map((p) => p.key)).toEqual(["employee"]);
  });

  // The Folder Access page groups personas behind a library toggle (2026-08-09). These assert
  // the grouping it derives, because the client's first sketch of that toggle had the two lists
  // swapped — and a UI that files C-Level under the approval library is suggesting the widest
  // accidental grant this system has.
  it("puts exactly Head of Unit and PIC on the approval-library side of the toggle", () => {
    const staging = PERSONAS.filter(personaTouchesStaging).map((p) => p.key);
    // hou_hc retired 2026-08-17; the HC-cleared PIC still belongs here, via UPLHC and DELSHC.
    expect(staging.sort()).toEqual(["hou", "pic", "pic_hc"]);
  });

  it("keeps every C-Level and viewer persona off the approval-library side", () => {
    // GLOBAL/SEGVIEW on Staging would read every unapproved draft in a segment; MEMBER would
    // read a unit's. Asserted through the same helper the UI calls, so the screen cannot
    // disagree with the model.
    for (const key of ["clevel_global", "clevel_segment", "hod", "employee"]) {
      const p = personaByKey(key);
      expect(p).toBeDefined();
      expect(personaTouchesStaging(p!)).toBe(false);
    }
  });

  it("gives BOTH deletes only to the Head of Unit", () => {
    // Was "never puts DEL and DELS on the same persona" until 2026-08-15. That rule is genuinely
    // superseded: a Head of Unit deletes pending work (DELS, approval library) AND carries out
    // approved deletion requests (DEL, Documents). An approver can only approve what they can
    // perform, so the combination is the requirement, not an accident.
    //
    // It is still asserted as an EXACT set, because the original worry stands for everyone else:
    // the two roles map to the same permission level and are told apart only by LIBRARY_ROLES, so
    // a persona quietly acquiring both reaches approved documents and other people's pending ones.
    const both = PERSONAS.filter(
      (p) => p.roles.indexOf("DEL") !== -1 && p.roles.indexOf("DELS") !== -1,
    );
    // hou_hc retired 2026-08-17.
    expect(both.map((p) => p.key)).toEqual(["hou"]);
  });

  it("keeps C-Level and Head of Department off every Staging role", () => {
    // The families that see widest must never reach unapproved drafts: a C-Level on
    // Staging reads a whole segment's, an HoD a whole department's.
    for (const key of ["clevel_global", "clevel_segment", "hod"]) {
      for (const stagingRole of ["UPL", "APR", "DELS"]) {
        expect(rolesOf(key)).not.toContain(stagingRole);
      }
    }
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

// The site-entry matchers behind src/shared/siteEntryGroup.ts. A member outside the entry group
// holds a folder grant they cannot navigate to, so a missed match here presents as a permissions
// bug that is not one.
describe("isSiteEntryGroupTitle", () => {
  it("matches the resolved name", () => {
    expect(isSiteEntryGroupTitle(siteEntryGroupTitle())).toBe(true);
  });

  it("ignores case and surrounding whitespace", () => {
    // The title reaches this function from a search result, a Group Map row and a name box, and
    // only one of those carries the stored casing.
    expect(isSiteEntryGroupTitle(`  ${siteEntryGroupTitle().toUpperCase()}  `)).toBe(true);
    expect(isSiteEntryGroupTitle(siteEntryGroupTitle().toLowerCase())).toBe(true);
  });

  it("does not match another group", () => {
    expect(isSiteEntryGroupTitle("GHO_GF_CORU_UPLOADER")).toBe(false);
  });

  it("treats blank and absent as not the entry group, rather than throwing", () => {
    expect(isSiteEntryGroupTitle("")).toBe(false);
    expect(isSiteEntryGroupTitle(undefined)).toBe(false);
  });

  // setSiteEntryName mutates module state, so this restores the default afterwards. Without that,
  // a later test comparing against the resolved title would pass or fail on test ORDER.
  it("follows the name resolved for THIS site", () => {
    const before = siteEntryGroupTitle();
    try {
      // A half-renamed site is the live case: CRS lists alongside DMS_SITE_MEMBERS (2026-08-05).
      setSiteEntryName(["CRS_SITE_MEMBERS"]);
      expect(siteEntryGroupTitle()).toBe("CRS_SITE_MEMBERS");
      expect(isSiteEntryGroupTitle("crs_site_members")).toBe(true);
      expect(isSiteEntryGroupTitle("DMS_SITE_MEMBERS")).toBe(false);
    } finally {
      setSiteEntryName([before]);
    }
  });
});

describe("findSiteEntryGroup", () => {
  // Titles are read INSIDE each test, never captured at describe-time: describe bodies all run
  // before any test does, so a captured title would be the pre-mutation one.
  const other = { id: 3, title: "GHO_GF_CORU_UPLOADER" };

  it("finds it among other groups", () => {
    const entry = { id: 12, title: siteEntryGroupTitle() };
    expect(findSiteEntryGroup([other, entry])).toBe(entry);
  });

  it("returns undefined when it is ABSENT — which is not the same as unreadable", () => {
    // The caller must have read the list successfully first; siteEntryGroup.ts keeps those two
    // states apart, because only one of them is a setup step the admin can act on.
    expect(findSiteEntryGroup([other])).toBeUndefined();
    expect(findSiteEntryGroup([])).toBeUndefined();
    expect(findSiteEntryGroup(undefined)).toBeUndefined();
  });

  it("matches case-insensitively, like the title check", () => {
    const odd = { id: 9, title: siteEntryGroupTitle().toLowerCase() };
    expect(findSiteEntryGroup([odd])).toBe(odd);
  });
});

describe("validateGroupName", () => {
  it("accepts the convention's own output", () => {
    expect(validateGroupName("GHO_GF_CORU_UPLOADER")).toEqual([]);
    expect(validateGroupName(suggestGroupName("GHO", ["GF", "CORU"], "APR"))).toEqual([]);
  });

  it("rejects blank and whitespace-only, and says nothing else about them", () => {
    // One message, not four: an empty box does not also need to be told about illegal
    // characters and length.
    expect(validateGroupName("")).toEqual(["Enter a group name."]);
    expect(validateGroupName("   ")).toEqual(["Enter a group name."]);
  });

  it("names the characters SharePoint refuses", () => {
    const errs = validateGroupName("GHO/GF");
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("/");
  });

  it("lists every illegal character it found, not just the first", () => {
    const errs = validateGroupName('a#b%c"');
    expect(errs[0]).toContain("#");
    expect(errs[0]).toContain("%");
    expect(errs[0]).toContain('"');
  });

  it("catches a duplicate and NAMES the existing group", () => {
    const errs = validateGroupName("GHO_GF_CORU_UPLOADER", ["GHO_GF_CORU_UPLOADER"]);
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("GHO_GF_CORU_UPLOADER");
  });

  it("treats a duplicate case-insensitively, because SharePoint does", () => {
    // Reporting these as available produces a server failure the admin cannot explain from
    // anything on screen.
    expect(validateGroupName("gho_gf_coru_uploader", ["GHO_GF_CORU_UPLOADER"]).length).toBe(1);
    expect(validateGroupName("  GHO_GF_CORU_UPLOADER  ", ["gho_gf_coru_uploader"]).length).toBe(1);
  });

  it("does not flag a different group", () => {
    expect(validateGroupName("GHO_GF_CORU_APPROVER", ["GHO_GF_CORU_UPLOADER"])).toEqual([]);
  });

  it("has no opinion when it was given no list to compare against", () => {
    expect(validateGroupName("Anything At All")).toEqual([]);
    expect(validateGroupName("Anything At All", [])).toEqual([]);
  });

  it("rejects a name past SharePoint's limit", () => {
    expect(validateGroupName("x".repeat(256)).length).toBe(1);
    expect(validateGroupName("x".repeat(255))).toEqual([]);
  });
});

/**
 * The 2026-08-15 correction — spec `2026-08-15-deletion-and-share-requests-design.md`.
 *
 * Pinned because the compositions ARE the security model, and until now nothing tested them: the
 * whole suite passed both before and after the roles moved. Each case below is a rule that is silent
 * when broken.
 */
describe("corrected role model (2026-08-15)", () => {
  const rolesOf = (key: string): string[] => PERSONAS.filter((p) => p.key === key)[0].roles as string[];

  it("a PIC deletes in the approval library but NOT in Documents", () => {
    // DELS is Staging-only; DEL is the Documents one. A PIC holding DEL would make the whole
    // deletion-request workflow pointless — they would simply delete.
    expect(rolesOf("pic")).toContain("DELS");
    expect(rolesOf("pic")).not.toContain("DEL");
  });

  it("a PIC cannot share — that is the request workflow", () => {
    expect(rolesOf("pic")).not.toContain("SHARE");
  });

  it("a Head of Unit can upload — through UPLHC since 2026-08-17", () => {
    // The client extended it: "any type of HOU can upload to HC". UPLHC REPLACES UPL rather than
    // joining it, because it is a superset — LIBRARY_ROLES lists it on the normal approval library
    // too — so both would be two rows granting the same thing on the same folder.
    //
    // This only works because collectMembership counts UPLHC as an uploader role, fixed the same
    // day. Before that it recognised UPL alone, so pic_hc had no uploadable paths at all and HC
    // upload had never worked for anyone.
    expect(rolesOf("hou")).toContain("UPLHC");
    expect(rolesOf("hou")).not.toContain("UPL");
  });

  it("a Head of Unit can perform BOTH things they approve", () => {
    // An approver can only approve what they can carry out. Without DEL an approved deletion fails
    // at the last step; without SHARE an approved share does.
    expect(rolesOf("hou")).toContain("DEL");
    expect(rolesOf("hou")).toContain("SHARE");
  });

  it("C-Level acts without approval — both delete and share", () => {
    for (const key of ["clevel_global", "clevel_segment"]) {
      expect(rolesOf(key)).toContain("DEL");
      expect(rolesOf(key)).toContain("SHARE");
    }
  });

  it("but the Head of Department no longer does either — view only since 2026-08-17", () => {
    // "HOD no need deletion power, he only view" / "do not have share functionality that is HOU".
    // Asserted separately from C-Level because they were one rule until today, and a test that still
    // grouped them would pass the moment DEL crept back onto this persona.
    expect(rolesOf("hod")).toEqual(["DEPTVIEW"]);
  });

  it("SHARE NEVER reaches the approval library", () => {
    // Sharing an unapproved draft hands someone a document nobody has approved yet.
    expect(STAGING_FACING_ROLES).not.toContain("SHARE");
  });

  it("C-Level still never touches the approval library, despite its new powers", () => {
    // The oldest invariant here: a segment-wide viewer on Staging reads every unapproved draft in
    // that segment. Adding DEL and SHARE must not have dragged them in.
    for (const key of ["clevel_global", "clevel_segment"]) {
      const persona = PERSONAS.filter((p) => p.key === key)[0];
      expect(personaTouchesStaging(persona)).toBe(false);
    }
  });

  it("Head of Department still never touches the approval library", () => {
    expect(personaTouchesStaging(PERSONAS.filter((p) => p.key === "hod")[0])).toBe(false);
  });

  it("PIC and Head of Unit both touch the approval library", () => {
    for (const key of ["pic", "hou"]) {
      expect(personaTouchesStaging(PERSONAS.filter((p) => p.key === key)[0])).toBe(true);
    }
  });
});

// ── A persona's SUMMARY must not contradict its ROLES ────────────────────────
//
// Both `hou` and `pic` gained roles on 2026-08-15 and kept summaries written for the previous
// model, so until 2026-08-17 the persona picker told administrators that a Head of Unit
// "Cannot upload" and a PIC "Cannot approve or delete" — the opposite of what those personas do.
// Nothing caught it, because nothing in this suite asserted on prose, and prose is the only part
// of a persona an administrator actually reads before choosing one.
//
// This is deliberately a NEGATIVE check. Asserting that a summary describes every role would
// force a brittle word-for-word match; asserting it never DENIES a role it holds catches the
// drift that actually happened, and stays quiet about wording.
describe("persona summaries do not deny a capability the persona has", () => {
  // Literal patterns, never built from a string at runtime: @rushstack/security/no-unsafe-regexp
  // rejects the dynamic form, and the warning baseline is a hard 19.
  const CLAIMS: Array<{ verb: string; deny: RegExp; roles: GroupMapRole[] }> = [
    { verb: "upload",  deny: /(cannot|can't|no)\s+upload/i,  roles: ["UPL", "UPLHC"] },
    { verb: "approve", deny: /(cannot|can't|no)\s+approve/i, roles: ["APR"] },
    { verb: "delete",  deny: /(cannot|can't|no)\s+delete/i,  roles: ["DEL", "DELS", "DELSHC"] },
    // "shar" so one pattern covers both "share" and "sharing".
    { verb: "share",   deny: /(cannot|can't|no)\s+shar/i,    roles: ["SHARE"] },
  ];

  for (const persona of PERSONAS) {
    for (const claim of CLAIMS) {
      const holds = claim.roles.some((r) => persona.roles.indexOf(r) !== -1);
      if (!holds) continue;
      it(`${persona.key} holds ${claim.roles.join("/")} so its summary must not deny "${claim.verb}"`, () => {
        expect(claim.deny.test(persona.summary)).toBe(false);
      });
    }
  }
});

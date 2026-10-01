import {
  AccessRow,
  UserGroupRef,
  GroupMapRenameRow,
  personDisplay,
  personaForRoles,
  plannedGroupRenames,
  summarizeUserAccess,
  describeGroupAccess,
} from "./userAccess";
import { PERSONAS, siteEntryGroupTitle } from "./groupMapModel";

const resolve = {
  segmentLabel: (guid: string): string =>
    guid === "seg-gho" ? "Group Head Office" : guid === "seg-mho" ? "Minamas Head Office" : "",
  tierChain: (guid: string): string[] =>
    guid === "term-tax" ? ["Group Finance", "Tax"] : guid === "term-gf" ? ["Group Finance"] : [],
};

const row = (over: Partial<AccessRow> & { groupId: number }): AccessRow => ({
  segment: "seg-gho",
  unitTermGuid: "term-tax",
  role: "UPL",
  scope: "Folder",
  target: "Staging",
  ...over,
});

const g = (id: number, title: string): UserGroupRef => ({ id, title });

describe("personaForRoles", () => {
  it("matches a persona on its exact role set, whatever the order or case", () => {
    // pic lost DELS a THIRD time on 2026-09-17 — direct delete moved to a Power Automate proxy
    // (client: "they cannot directly delete anymore but someone do it on their behalf via power
    // automate"), so its exact set is back down to ["UPL"] alone.
    // See 2026-09-17-proxy-deletion-via-power-automate-design.md.
    expect(personaForRoles(["UPL"])?.key).toBe("pic");
    expect(personaForRoles(["upl"])?.key).toBe("pic");
    // hod's set lost DEL/DELHC the same day — SHARE/SHAREHC stay, deletion moved to the proxy.
    expect(personaForRoles(["SHARE", "DEPTVIEW", "SHAREHC"])?.key).toBe("hod");
    // hou_hc's set lost DELS/DEL/DELSHC/DELHC the same day — APRHC/SHARE/UPLHC/SHAREHC remain.
    expect(personaForRoles(["APRHC", "SHARE", "UPLHC", "SHAREHC"])?.key).toBe("hou_hc");
    // The plain hou set: no HC role reaches it at all — a role held by two personas cannot grant to
    // one and withhold from the other, so any single HC role here would match nothing.
    expect(personaForRoles(["APR", "SHARE", "UPL"])?.key).toBe("hou");
  });

  it("ignores duplicates, which a row set legitimately contains", () => {
    expect(personaForRoles(["UPL", "UPL"])?.key).toBe("pic");
  });

  // The whole reason this is an exact match. `employee` is ["MEMBER"] and `employee_hc` is
  // ["MEMBER","MEMBERHC"], so a subset rule would let a CLEARED viewer group answer as the plain
  // one — presenting HC clearance as absent, which is the dangerous direction.
  it("does NOT match a subset or a superset", () => {
    expect(personaForRoles(["MEMBER"])?.key).toBe("employee");
    expect(personaForRoles(["MEMBER", "MEMBERHC"])?.key).toBe("employee_hc");
    expect(personaForRoles(["MEMBER", "MEMBERHC", "UPL"])).toBeUndefined();
    expect(personaForRoles(["APR"])).toBeUndefined();
  });

  it("answers undefined for nothing and for an unknown role", () => {
    expect(personaForRoles([])).toBeUndefined();
    expect(personaForRoles([""])).toBeUndefined();
    expect(personaForRoles(["NOT_A_ROLE"])).toBeUndefined();
  });

  it("can name every persona in the model, so none is unreachable", () => {
    for (const p of PERSONAS) {
      expect(personaForRoles(p.roles as unknown as string[])?.key).toBe(p.key);
    }
  });
});

describe("summarizeUserAccess", () => {
  it("reports the persona, the roles and the place for a mapped group", () => {
    // pic's mapping row is back to UPL alone since 2026-09-17 (DELS moved to a Power Automate
    // proxy). See 2026-09-17-proxy-deletion-via-power-automate-design.md.
    const s = summarizeUserAccess(
      [g(42, "GHO_GF_TAX_UPLOADER")],
      [row({ groupId: 42, role: "UPL" })],
      resolve,
    );
    expect(s.none).toBe(false);
    expect(s.groups).toHaveLength(1);
    expect(s.groups[0].persona?.key).toBe("pic");
    expect(s.groups[0].roles).toEqual(["UPL"]);
    expect(s.groups[0].places).toEqual([
      {
        segmentLabel: "Group Head Office",
        tierChain: ["Group Finance", "Tax"],
        scope: "Folder",
        target: "Staging",
      },
    ]);
  });

  // An approver group carries six role rows at ONE term. Listing that term six times says nothing
  // the first line did not.
  it("collapses many rows at one place into a single place", () => {
    // Plain hou's role set as of 2026-09-17 — DEL and DELS left the persona entirely (deletion moved
    // to a Power Automate proxy), leaving APR/SHARE/UPL.
    const roles = ["APR", "SHARE", "UPL"];
    const s = summarizeUserAccess(
      [g(7, "GHO_GF_TAX_APPROVER")],
      roles.map((r) => row({ groupId: 7, role: r })),
      resolve,
    );
    expect(s.groups[0].places).toHaveLength(1);
    expect(s.groups[0].roles).toEqual(roles);
    expect(s.groups[0].persona?.key).toBe("hou");
  });

  it("keeps genuinely different places apart", () => {
    const s = summarizeUserAccess(
      [g(7, "X")],
      [
        row({ groupId: 7, unitTermGuid: "term-tax" }),
        row({ groupId: 7, unitTermGuid: "term-gf" }),
        row({ groupId: 7, segment: "seg-mho", unitTermGuid: "term-tax" }),
      ],
      resolve,
    );
    expect(s.groups[0].places).toHaveLength(3);
  });

  // The state that exists because creating a group stopped implying a mapping (2026-08-14), and the
  // most useful thing this page can tell someone asking why a person cannot get in.
  it("flags a group with no rows as granting nothing", () => {
    const s = summarizeUserAccess([g(9, "GHO_GF_TAX_EMPLOYEE")], [], resolve);
    expect(s.groups[0].unmapped).toBe(true);
    expect(s.groups[0].places).toEqual([]);
    expect(s.unmappedCount).toBe(1);
    expect(describeGroupAccess(s.groups[0])).toContain("Grants nothing yet");
  });

  it("flags the site-entry group, which is not a document grant", () => {
    const s = summarizeUserAccess([g(1, siteEntryGroupTitle())], [], resolve);
    expect(s.groups[0].siteEntry).toBe(true);
    expect(describeGroupAccess(s.groups[0])).toContain("Opens the site");
  });

  // The owners group holds Full Control on the WEB, so it has no mapping rows and never will. It
  // read "Grants nothing yet", in red, about the widest access on the site (client, 2026-09-07).
  it("flags the owners group instead of calling it unmapped", () => {
    const s = summarizeUserAccess([g(3, "Guthrie Central Repository System Owners")], [], resolve, 3);
    expect(s.groups[0].owners).toBe(true);
    expect(describeGroupAccess(s.groups[0])).toContain("Full control");
    expect(describeGroupAccess(s.groups[0])).not.toContain("Grants nothing");
  });

  // Matched on the ID, never the title — this client renames everything at import, and a title match
  // would silently stop recognising the group and put the red message back.
  it("does not flag owners when the id was not resolved", () => {
    const s = summarizeUserAccess([g(3, "Guthrie Central Repository System Owners")], [], resolve);
    expect(s.groups[0].owners).toBe(false);
    expect(describeGroupAccess(s.groups[0])).toContain("Grants nothing yet");
  });

  // The count answers "why can this person not get in", and neither of these two is ever the reason.
  // Counting them produced "3 groups on this site · 2 of them grant nothing" about somebody who
  // could open the site AND administer it.
  it("excludes the site-entry and owners groups from the grants-nothing count", () => {
    const s = summarizeUserAccess(
      [g(1, siteEntryGroupTitle()), g(3, "Owners"), g(9, "GHO_GF_TAX_EMPLOYEE")],
      [],
      resolve,
      3,
    );
    expect(s.groups).toHaveLength(3);
    expect(s.unmappedCount).toBe(1);
  });

  // Joining on the NAME would break on a rename (the id survives, the stored GroupName goes stale)
  // and on two groups renamed alike.
  it("joins rows on the group id, never the title", () => {
    const s = summarizeUserAccess([g(42, "RENAMED_SINCE")], [row({ groupId: 42 })], resolve);
    expect(s.groups[0].roles).toEqual(["UPL"]);
    const other = summarizeUserAccess([g(43, "RENAMED_SINCE")], [row({ groupId: 42 })], resolve);
    expect(other.groups[0].unmapped).toBe(true);
  });

  // Reporting an unresolved chain as a bare leaf presents a UNIT as a DEPARTMENT — the difference
  // between granting one unit and granting all of them.
  it("leaves an unresolvable chain EMPTY rather than guessing a leaf", () => {
    const s = summarizeUserAccess(
      [g(42, "X")],
      [row({ groupId: 42, unitTermGuid: "term-unknown", segment: "seg-unknown" })],
      resolve,
    );
    expect(s.groups[0].places[0].tierChain).toEqual([]);
    expect(s.groups[0].places[0].segmentLabel).toBe("");
  });

  it("defaults a blank scope to Folder and tolerates missing fields", () => {
    const s = summarizeUserAccess([g(42, "X")], [{ groupId: 42, role: "UPL" }], resolve);
    expect(s.groups[0].places[0].scope).toBe("Folder");
    expect(s.groups[0].places[0].target).toBe("");
  });

  it("says so when the person is in no groups", () => {
    const s = summarizeUserAccess([], [row({ groupId: 42 })], resolve);
    expect(s.none).toBe(true);
    expect(s.groups).toEqual([]);
  });

  it("ignores rows belonging to groups the person is not in", () => {
    const s = summarizeUserAccess([g(42, "X")], [row({ groupId: 99 })], resolve);
    expect(s.groups[0].unmapped).toBe(true);
  });
});

describe("describeGroupAccess", () => {
  const base = {
    groupId: 1,
    groupTitle: "X",
    places: [],
    unmapped: false,
    siteEntry: false,
    owners: false,
  };

  it("prefers the persona label", () => {
    expect(
      describeGroupAccess({
        ...base,
        roles: ["UPL"],
        roleLabels: ["Upload"],
        persona: PERSONAS.filter((p) => p.key === "pic")[0],
      }),
    ).toBe("Upload only");
  });

  // An unrecognised role set must still say something true — a hand-made group with one odd role is
  // exactly the case an admin is investigating.
  it("falls back to the role labels, never to blank", () => {
    expect(describeGroupAccess({ ...base, roles: ["APR"], roleLabels: ["Approve"] })).toBe("Approve");
  });

  it("says something even with rows but no roles", () => {
    const d = describeGroupAccess({ ...base, roles: [], roleLabels: [] });
    expect(d.length).toBeGreaterThan(0);
    expect(d).toContain("No roles recorded");
  });
});

describe("personDisplay", () => {
  // The reported symptom, twice over: a guest whose display name is its own local part, and a
  // free-text address that resolves to itself.
  it("drops a display name that is only the address again", () => {
    expect(personDisplay("clarencechojinheng", "clarencechojinheng@gmail.com"))
      .toBe("clarencechojinheng@gmail.com");
    expect(personDisplay("test@gmail.com", "test@gmail.com")).toBe("test@gmail.com");
    expect(personDisplay("CLARENCECHOJINHENG", "clarencechojinheng@gmail.com"))
      .toBe("clarencechojinheng@gmail.com");
  });

  // A real name is the USEFUL half and must survive.
  it("keeps a real name alongside the address", () => {
    expect(personDisplay("Wyanet Falasifa", "wyanet@trinergydigital.com"))
      .toBe("Wyanet Falasifa · wyanet@trinergydigital.com");
  });

  it("never renders blank, whichever half is missing", () => {
    expect(personDisplay("Crystal Kong", "")).toBe("Crystal Kong");
    expect(personDisplay("", "crystal@trinergydigital.com")).toBe("crystal@trinergydigital.com");
    expect(personDisplay(undefined, undefined)).toBe("");
    expect(personDisplay("  ", " a@b.com ")).toBe("a@b.com");
  });

  // A prefix match is not enough: "clarence" must not swallow a colleague's address.
  it("matches the local part exactly, not as a prefix", () => {
    expect(personDisplay("clarence", "clarencechojinheng@gmail.com"))
      .toBe("clarence · clarencechojinheng@gmail.com");
  });
});

describe("plannedGroupRenames", () => {
  const chains: Record<string, string[] | undefined> = {
    "term-tax": ["GF", "TAX"],
    "term-tax-new": ["GF", "TAXX"],
  };
  const chainFor = (guid: string): string[] | undefined => chains[guid.toLowerCase()];

  // ONE row, matching `pic`'s EXACT role set since DELS moved to a Power Automate proxy on
  // 2026-09-17 (see 2026-09-17-proxy-deletion-via-power-automate-design.md) — pic is back down to
  // `["UPL"]` alone, so a lone `UPL` row is now the correctly-shaped fixture rather than a trap.
  const picRows = (overUpl?: Partial<GroupMapRenameRow>): GroupMapRenameRow[] => [
    { GroupId: "101", GroupName: "GHO_GF_TAX_UPLOADER", Role: "UPL", UnitTermGuid: "term-tax", ...overUpl },
  ];

  it("offers a rename for a group mapped at a just-renamed term, when the name is now stale", () => {
    // "TAX" -> "TAXX": the chain the caller supplies already reflects the new code.
    const out = plannedGroupRenames(
      ["term-tax"],
      false,
      "GHO",
      picRows(),
      (guid) => (guid.toLowerCase() === "term-tax" ? ["GF", "TAXX"] : chainFor(guid)),
    );
    expect(out).toEqual([{ groupId: 101, from: "GHO_GF_TAX_UPLOADER", to: "GHO_GF_TAXX_UPLOADER" }]);
  });

  it("offers nothing for a term that was not renamed", () => {
    const out = plannedGroupRenames(["term-other"], false, "GHO", picRows(), chainFor);
    expect(out).toEqual([]);
  });

  it("offers nothing when the computed name already matches the stored one", () => {
    // Nothing actually changed for this term, so the group's current name is already correct.
    const out = plannedGroupRenames(["term-tax"], false, "GHO", picRows(), chainFor);
    expect(out).toEqual([]);
  });

  it("collapses a multi-role group (hou) to ONE entry, not one per role row", () => {
    // hou's role set as of 2026-09-17 — DEL and DELS left the persona entirely (deletion moved to a
    // Power Automate proxy), leaving APR/SHARE/UPL.
    const rows: GroupMapRenameRow[] = [
      { GroupId: "202", GroupName: "GHO_GF_TAX_APPROVER", Role: "APR", UnitTermGuid: "term-tax" },
      { GroupId: "202", GroupName: "GHO_GF_TAX_APPROVER", Role: "SHARE", UnitTermGuid: "term-tax" },
      { GroupId: "202", GroupName: "GHO_GF_TAX_APPROVER", Role: "UPL", UnitTermGuid: "term-tax" },
    ];
    const out = plannedGroupRenames(
      ["term-tax"],
      false,
      "GHO",
      rows,
      (guid) => (guid.toLowerCase() === "term-tax" ? ["GF", "TAXX"] : chainFor(guid)),
    );
    expect(out).toEqual([{ groupId: 202, from: "GHO_GF_TAX_APPROVER", to: "GHO_GF_TAXX_APPROVER" }]);
  });

  // This blank-guid shape is what `buildGroupMapRow` actually writes for GLOBAL ("C-Level — all
  // segments") — its `UnitTermGuid` is forced empty because GLOBAL carries no term at all. A SEGVIEW
  // ("C-Level — one segment") row is NOT blank in production — see the test right below this one.
  it("offers a segment-tier row (blank UnitTermGuid, the GLOBAL shape) ONLY when the segment itself was recoded", () => {
    const segRow: GroupMapRenameRow = {
      GroupId: "303",
      GroupName: "PCAR_C_LEVEL",
      Role: "SEGVIEW",
      UnitTermGuid: "",
    };
    expect(plannedGroupRenames(["term-tax"], false, "PCT", [segRow], chainFor)).toEqual([]);
    expect(plannedGroupRenames([], true, "PCT", [segRow], chainFor)).toEqual([
      { groupId: 303, from: "PCAR_C_LEVEL", to: "PCT_C_LEVEL" },
    ]);
  });

  // ⚠⚠ THE BUG FOUND LIVE ON GHO, 2026-09-16 (a second one, reported right after the unit/department
  // fix above shipped): "I notice c Level is still not renamed." A SEGVIEW row's `UnitTermGuid` is
  // NOT blank in production — `bulkGroups.ts`'s `push()` sets it to the SEGMENT's own term-SET guid
  // (`tierGuid: segment.termSetGuid`), which `codeChainFor` has never heard of (it only knows terms
  // walked from INSIDE that set) — so the row failed the "chain could not be resolved" guard and was
  // silently dropped, exactly like an ordinary unit row with a broken parent link. Fixed by threading
  // the segment's own guid through as a 6th, optional parameter.
  it("offers a segment-tier (SEGVIEW) row whose UnitTermGuid equals the segment's OWN term set guid", () => {
    const segRow: GroupMapRenameRow = {
      GroupId: "606",
      GroupName: "GHO_C_LEVEL",
      Role: "SEGVIEW",
      UnitTermGuid: "seg-guid-gho", // the segment's own term SET id — not a term chainFor() can resolve
    };
    // Without the segment's own guid supplied, this is the bug: the row is silently skipped.
    expect(plannedGroupRenames([], true, "GHOS", [segRow], chainFor)).toEqual([]);
    // With it supplied, the row is recognised as segment-tier (no chain to resolve) and renamed.
    expect(
      plannedGroupRenames([], true, "GHOS", [segRow], chainFor, "seg-guid-gho"),
    ).toEqual([{ groupId: 606, from: "GHO_C_LEVEL", to: "GHOS_C_LEVEL" }]);
    // Case-insensitive, matching every other guid comparison in this function.
    expect(
      plannedGroupRenames([], true, "GHOS", [segRow], chainFor, "SEG-GUID-GHO"),
    ).toEqual([{ groupId: 606, from: "GHO_C_LEVEL", to: "GHOS_C_LEVEL" }]);
  });

  // ⚠⚠ THE BUG FOUND LIVE ON GHO, 2026-09-16: a UNIT-tier group (the overwhelming majority of every
  // segment's groups) was never offered a rename by a pure segment recode — `renamedTermGuids` is
  // empty in that case (no term was renamed, only the segment's own code), so the old gate asked
  // "was THIS term renamed" and always answered no. Recoding GHO to GHOS therefore left every
  // `GHO_<dept>_<unit>_<role>` group stale, and re-running bulk provisioning built a parallel
  // `GHOS_*` set instead of renaming the real one.
  it("offers a unit-tier row too when the SEGMENT was recoded, with no term guid renamed", () => {
    // `renamedTermGuids` is deliberately empty — this is a pure segment recode, exactly GHO's case.
    const out = plannedGroupRenames(
      [],
      true,
      "GHOS",
      picRows(),
      (guid) => (guid.toLowerCase() === "term-tax" ? ["GF", "TAX"] : chainFor(guid)),
    );
    expect(out).toEqual([{ groupId: 101, from: "GHO_GF_TAX_UPLOADER", to: "GHOS_GF_TAX_UPLOADER" }]);
  });

  // The safety this fix must not lose: a segment recode alone must never GUESS a name for a unit
  // whose code chain could not be resolved — it still has to be skipped, exactly as for a term rename.
  it("still skips a unit row on a segment recode when its code chain could not be resolved", () => {
    const out = plannedGroupRenames([], true, "GHOS", picRows(), () => undefined);
    expect(out).toEqual([]);
  });

  // ⚠ THE THIRD TIER, checked explicitly rather than inferred from the unit-tier test above — this is
  // a DEPARTMENT-scope group (`hod`'s five-role set: DEPTVIEW/DEL/SHARE/DELHC/SHAREHC), mapped at the
  // DEPARTMENT's own term guid with a one-element chain, not a unit's two-element one. Same code path
  // as the unit-tier fix, but worth pinning on its own rather than trusted by similarity.
  it("offers a department-tier (HOD) row too on a segment recode", () => {
    // hod's role set as of 2026-09-17 — DEL and DELHC left the persona entirely (deletion moved to a
    // Power Automate proxy), leaving DEPTVIEW/SHARE/SHAREHC.
    const rows: GroupMapRenameRow[] = [
      { GroupId: "505", GroupName: "GHO_GF_HOD", Role: "DEPTVIEW", UnitTermGuid: "term-gf" },
      { GroupId: "505", GroupName: "GHO_GF_HOD", Role: "SHARE", UnitTermGuid: "term-gf" },
      { GroupId: "505", GroupName: "GHO_GF_HOD", Role: "SHAREHC", UnitTermGuid: "term-gf" },
    ];
    const out = plannedGroupRenames(
      [],
      true,
      "GHOS",
      rows,
      (guid) => (guid.toLowerCase() === "term-gf" ? ["GF"] : undefined),
    );
    expect(out).toEqual([{ groupId: 505, from: "GHO_GF_HOD", to: "GHOS_GF_HOD" }]);
  });

  it("skips a term whose code chain could not be resolved, rather than guessing a name", () => {
    const out = plannedGroupRenames(["term-tax"], false, "GHO", picRows(), () => undefined);
    expect(out).toEqual([]);
  });

  it("skips a group whose role set matches no persona exactly", () => {
    const rows: GroupMapRenameRow[] = [
      { GroupId: "404", GroupName: "GHO_GF_TAX_ODD", Role: "ENTRY", UnitTermGuid: "term-tax" },
    ];
    const out = plannedGroupRenames(
      ["term-tax"],
      false,
      "GHO",
      rows,
      (guid) => (guid.toLowerCase() === "term-tax" ? ["GF", "TAXX"] : chainFor(guid)),
    );
    expect(out).toEqual([]);
  });

  it("skips a row with no GroupId, and matches case-insensitively on the term guid", () => {
    const rows: GroupMapRenameRow[] = [
      { GroupId: "", GroupName: "orphan", Role: "UPL", UnitTermGuid: "term-tax" },
      ...picRows({ UnitTermGuid: "TERM-TAX" }),
    ];
    const out = plannedGroupRenames(
      ["Term-Tax"],
      false,
      "GHO",
      rows,
      (guid) => (guid.toLowerCase() === "term-tax" ? ["GF", "TAXX"] : chainFor(guid)),
    );
    expect(out).toEqual([{ groupId: 101, from: "GHO_GF_TAX_UPLOADER", to: "GHO_GF_TAXX_UPLOADER" }]);
  });

  it("sorts results by the OLD name", () => {
    // `employee` ("MEMBER" alone) is an exact one-role persona, so one row per group is enough here.
    const rows: GroupMapRenameRow[] = [
      { GroupId: "2", GroupName: "GHO_ZZ_VIEWER", Role: "MEMBER", UnitTermGuid: "term-zz" },
      { GroupId: "1", GroupName: "GHO_AA_VIEWER", Role: "MEMBER", UnitTermGuid: "term-aa" },
    ];
    const out = plannedGroupRenames(
      ["term-zz", "term-aa"],
      false,
      "GHO",
      rows,
      (guid) =>
        guid.toLowerCase() === "term-zz" ? ["ZZZ"] : guid.toLowerCase() === "term-aa" ? ["AAA"] : undefined,
    );
    expect(out.map((r) => r.from)).toEqual(["GHO_AA_VIEWER", "GHO_ZZ_VIEWER"]);
  });
});

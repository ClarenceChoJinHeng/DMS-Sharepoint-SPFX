import { SpGroupMember } from "./spGroupsFilter";
import {
  MemberMap,
  buildRemoval,
  loadOf,
  memberCountLabel,
  memberLabel,
  membersOf,
  membershipIndex,
  otherAllowedGroups,
  removalVerdict,
  sortMembers,
  unreadableGroups,
} from "./accessMembers";

function user(id: number, title: string, email = ""): SpGroupMember {
  return { id, title, email, loginName: `i:0#.f|membership|${email || title}` };
}

const ALICE = user(11, "Alice Tan", "alice@example.com");
const BAYA = user(12, "Baya Sen", "baya@example.com");
const SITI = user(13, "Siti Rahman", "siti@example.com");

// Group 101 = ..._UPL, 102 = ..._APR, 103 = a group that grants nothing at this scope.
const TITLES: Record<number, string> = {
  101: "GHO_GF_CORU_UPL",
  102: "GHO_GF_CORU_APR",
  103: "GHO_GF_CORU_MEMBER",
};
const titleOf = (id: number): string => TITLES[id] ?? "";

function loaded(members: SpGroupMember[]): { state: "loaded"; members: SpGroupMember[] } {
  return { state: "loaded", members };
}

describe("loadOf / membersOf — three states, never two", () => {
  it("treats a group never fetched as LOADING, not empty", () => {
    // The distinction the whole module rests on: an admin told a group is empty stops looking for
    // the person they came to remove.
    expect(loadOf({}, 101).state).toBe("loading");
  });

  it("returns [] for an unreadable group, but does not call it loaded", () => {
    const map: MemberMap = { 101: { state: "error", message: "HTTP 500" } };
    expect(membersOf(map, 101)).toEqual([]);
    expect(loadOf(map, 101).state).toBe("error");
  });

  it("survives a missing map rather than throwing", () => {
    expect(loadOf(undefined as unknown as MemberMap, 101).state).toBe("loading");
  });
});

describe("membershipIndex", () => {
  it("records every allowed group a person is in", () => {
    const map: MemberMap = { 101: loaded([ALICE, SITI]), 102: loaded([SITI]) };
    const idx = membershipIndex([101, 102], map);
    expect(idx.get(ALICE.id)).toEqual([101]);
    expect(idx.get(SITI.id)).toEqual([101, 102]);
  });

  it("IGNORES groups that are not allowed at this scope", () => {
    // Membership of a group that grants nothing here is not a reason to warn that access survives.
    // Counting it would warn on nearly every row, and the one case that matters would be lost.
    const map: MemberMap = { 101: loaded([ALICE]), 103: loaded([ALICE]) };
    const idx = membershipIndex([101], map);
    expect(idx.get(ALICE.id)).toEqual([101]);
  });

  it("keys on user id, not email — a principal may have no email", () => {
    const noEmail = user(14, "Service Account");
    const map: MemberMap = { 101: loaded([noEmail]), 102: loaded([noEmail]) };
    expect(membershipIndex([101, 102], map).get(14)).toEqual([101, 102]);
  });

  it("does not double-count a group listed twice", () => {
    const map: MemberMap = { 101: loaded([ALICE]) };
    expect(membershipIndex([101, 101], map).get(ALICE.id)).toEqual([101]);
  });

  it("contributes nothing for an unreadable group", () => {
    const map: MemberMap = { 101: loaded([ALICE]), 102: { state: "error", message: "x" } };
    expect(membershipIndex([101, 102], map).get(ALICE.id)).toEqual([101]);
  });
});

describe("otherAllowedGroups — 'did my removal work?'", () => {
  const map: MemberMap = { 101: loaded([ALICE, SITI]), 102: loaded([SITI]) };
  const idx = membershipIndex([101, 102], map);

  it("names the group that keeps the person in", () => {
    expect(otherAllowedGroups(SITI.id, 101, idx, titleOf)).toEqual(["GHO_GF_CORU_APR"]);
  });

  it("excludes the group being removed from", () => {
    expect(otherAllowedGroups(SITI.id, 101, idx, titleOf)).not.toContain("GHO_GF_CORU_UPL");
  });

  it("is empty when this is the person's only allowed group", () => {
    expect(otherAllowedGroups(ALICE.id, 101, idx, titleOf)).toEqual([]);
  });

  it("is empty for a person in no allowed group at all", () => {
    expect(otherAllowedGroups(999, 101, idx, titleOf)).toEqual([]);
  });

  it("drops a group whose title cannot be resolved rather than listing a blank", () => {
    const m: MemberMap = { 101: loaded([ALICE]), 555: loaded([ALICE]) };
    const i = membershipIndex([101, 555], m);
    expect(otherAllowedGroups(ALICE.id, 101, i, titleOf)).toEqual([]);
  });
});

describe("unreadableGroups", () => {
  it("names only the groups that failed, not the ones still loading", () => {
    const map: MemberMap = {
      101: loaded([ALICE]),
      102: { state: "error", message: "HTTP 403" },
      103: { state: "loading" },
    };
    expect(unreadableGroups([101, 102, 103], map, titleOf)).toEqual(["GHO_GF_CORU_APR"]);
  });

  it("is empty when everything read cleanly", () => {
    expect(unreadableGroups([101], { 101: loaded([]) }, titleOf)).toEqual([]);
  });
});

describe("buildRemoval + removalVerdict — the three answers", () => {
  const base = { allowedGroupIds: [101, 102], titleOf };

  it("ENDS access when the person is in no other allowed group", () => {
    const map: MemberMap = { 101: loaded([ALICE, BAYA]), 102: loaded([]) };
    const r = buildRemoval({ ...base, user: ALICE, groupId: 101, groupName: TITLES[101], map });
    expect(removalVerdict(r)).toBe("ends");
    expect(r.survivingGroups).toEqual([]);
  });

  it("SURVIVES when another allowed group also contains them", () => {
    // The case that makes a correct removal look like a failed one.
    const map: MemberMap = { 101: loaded([SITI]), 102: loaded([SITI]) };
    const r = buildRemoval({ ...base, user: SITI, groupId: 101, groupName: TITLES[101], map });
    expect(removalVerdict(r)).toBe("survives");
    expect(r.survivingGroups).toEqual(["GHO_GF_CORU_APR"]);
  });

  it("is UNKNOWN — not 'ends' — when another allowed group could not be read", () => {
    // Asserting "this ends their access" from an incomplete picture is the failure worth avoiding,
    // and the incomplete half is invisible: a failed read renders as a collapsed row like any other.
    const map: MemberMap = { 101: loaded([ALICE]), 102: { state: "error", message: "HTTP 500" } };
    const r = buildRemoval({ ...base, user: ALICE, groupId: 101, groupName: TITLES[101], map });
    expect(removalVerdict(r)).toBe("unknown");
    expect(r.unreadable).toEqual(["GHO_GF_CORU_APR"]);
  });

  it("prefers SURVIVES over UNKNOWN — a known survivor is the more actionable fact", () => {
    const map: MemberMap = {
      101: loaded([SITI]),
      102: loaded([SITI]),
      103: { state: "error", message: "x" },
    };
    const r = buildRemoval({
      user: SITI, groupId: 101, groupName: TITLES[101], map,
      allowedGroupIds: [101, 102, 103], titleOf,
    });
    expect(removalVerdict(r)).toBe("survives");
  });

  it("flags the LAST member, because the grant does not leave with them", () => {
    const map: MemberMap = { 101: loaded([ALICE]), 102: loaded([]) };
    const r = buildRemoval({ ...base, user: ALICE, groupId: 101, groupName: TITLES[101], map });
    expect(r.lastMember).toBe(true);
  });

  it("does NOT claim last member when the list could not be read", () => {
    // Understating is the safe direction for a sentence about what is left behind.
    const map: MemberMap = { 101: { state: "error", message: "x" } };
    const r = buildRemoval({ ...base, user: ALICE, groupId: 101, groupName: TITLES[101], map });
    expect(r.lastMember).toBe(false);
  });
});

describe("memberLabel — never blank", () => {
  it("uses the title", () => {
    expect(memberLabel(ALICE)).toBe("Alice Tan");
  });

  it("falls back to the email when the title is blank", () => {
    // A SharePoint principal can legitimately have an empty Title; a row with no visible label
    // cannot be acted on with any confidence.
    expect(memberLabel(user(20, "   ", "ghost@example.com"))).toBe("ghost@example.com");
  });

  it("falls back to the id when there is nothing else", () => {
    expect(memberLabel(user(21, "", ""))).toBe("User 21");
  });
});

describe("sortMembers", () => {
  it("orders by label so the list never reshuffles between renders", () => {
    expect(sortMembers([SITI, ALICE, BAYA]).map((m) => m.id)).toEqual([11, 12, 13]);
  });

  it("does not mutate the input", () => {
    const rows = [SITI, ALICE];
    sortMembers(rows);
    expect(rows.map((m) => m.id)).toEqual([13, 11]);
  });
});

describe("memberCountLabel", () => {
  it("says 'no members' ONLY about a list that was read", () => {
    expect(memberCountLabel(loaded([]))).toBe("no members");
    expect(memberCountLabel({ state: "error", message: "x" })).toBe("could not read members");
    expect(memberCountLabel({ state: "loading" })).toBe("loading…");
  });

  it("singularises one member", () => {
    expect(memberCountLabel(loaded([ALICE]))).toBe("1 member");
    expect(memberCountLabel(loaded([ALICE, BAYA]))).toBe("2 members");
  });
});

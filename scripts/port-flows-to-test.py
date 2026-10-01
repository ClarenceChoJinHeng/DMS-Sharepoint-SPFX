"""Copies CRS flow exports into import packages pointed at the test site (ClarenceDMSTesting).
Reads PowerAutomateFlowsSDG/*.zip (never changes them), writes C:/tmp/flows-for-test/TEST_<name>.zip.
Usage: python scripts/port-flows-to-test.py
"""
import glob
import json
import os
import re
import zipfile

SRC = "PowerAutomateFlowsSDG"
OUT = r"C:\tmp\flows-for-test"

CRS_SITE = "https://sdguthrie.sharepoint.com/sites/CRS"
TEST_SITE = "https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting"
CRS_HOST = "sdguthrie.sharepoint.com"
TEST_HOST = "dcidigitalcom.sharepoint.com"
CRS_ACCOUNT = "gdc@sdguthrie.com"
TEST_ACCOUNT = "clarence@trinergydigital.com"

# CRS list GUID -> test list GUID (from the 2026-10-01 schema dumps of both sites)
LISTS = {
    "eeb1bb19-ec53-4c41-8dc9-ecf255979c9b": "a9342528-66be-458c-ba70-a6e3248a5133",  # Approval for Document
    "d935aa6d-dc48-4832-ba7e-eca485ecdff3": "2311cd83-90aa-4077-a647-65d251a5f426",  # Approval for HC Document
    "fbc062dd-fcd0-4458-9161-2bfc19c917fa": "e322e3a5-3687-4da3-94f8-e0b06c01dd7e",  # Restricted & Confidential Document
    "122a4aa9-65fb-479a-90a7-3866d19f51b4": "e4fb3b2c-6f20-4bbb-a4f2-e9a5301d4080",  # Highly Confidential Document
    "647fd644-ab00-4358-9b3f-49031083fcf0": "8c28ca71-4edb-445b-8ebd-5711dacd71d4",  # Archive
    "fc301054-abf5-4e53-957a-6f8c58a410b8": "bfcd6735-7605-47e7-a0d6-bbfff41bab7b",  # HC Archive
    "dd1bc5bd-daa7-491a-9dc7-b81212c391ab": "f413972a-49ae-476a-8a8f-df4562349918",  # CRS Submissions
    "1bcb1aa1-fbf1-4d1b-a6d2-88436824c297": "e47c85d5-f0ee-4190-99b8-b3c5575f95a4",  # CRS Requests
    "4ac219e5-b855-46a3-be4b-4e96bba546c5": "84ed065f-14e2-49b2-8b37-d40324587825",  # CRS Audit Log
}
# CRS view ids have no test equivalent: drop them (the step then reads the default view).
DROP_VIEWS = {"f3adfffa-09a1-473a-9e6a-3349a8e0e5ab"}  # CRS Audit Log / All Items

# Test site runs the bundled approver email, so the per-file email only handles files without a record.
SWITCH_OVER = "@empty(coalesce(triggerOutputs()?['body/SubmissionFileId'], ''))"
SWITCH_OVER_FLOWS = ("NotifyApprovers_", "HCNotifyApprovers_")

AUDIT_ACTOR = '"item/ActorEmail": "' + CRS_ACCOUNT + '"'
KEEP = "\u0000KEEP_AUDIT_ACTOR\u0000"


def drop_views(node):
    if isinstance(node, dict):
        for k in [k for k, v in node.items() if k in ("view", "parameters/view") and v in DROP_VIEWS]:
            del node[k]
        for v in node.values():
            drop_views(v)
    elif isinstance(node, list):
        for v in node:
            drop_views(v)


def add_switch_over(definition):
    for trigger in definition["triggers"].values():
        conditions = trigger.setdefault("conditions", [])
        if not any(c.get("expression") == SWITCH_OVER for c in conditions):
            conditions.append({"expression": SWITCH_OVER})


def port_definition(raw, file_name):
    doc = json.loads(raw)
    definition = doc["properties"]["definition"]
    drop_views(definition)
    if os.path.basename(file_name).startswith(SWITCH_OVER_FLOWS):
        add_switch_over(definition)
    text = json.dumps(doc, ensure_ascii=False)
    text = text.replace(AUDIT_ACTOR, KEEP)  # audit rows keep showing as Guthrie Document Centre
    text = text.replace(CRS_SITE, TEST_SITE).replace(CRS_HOST, TEST_HOST)
    text = re.sub(re.escape(CRS_ACCOUNT), TEST_ACCOUNT, text, flags=re.I)
    for crs, test in LISTS.items():
        text = re.sub(crs, test, text, flags=re.I)
    return text.replace(KEEP, AUDIT_ACTOR)


def leftovers(text):
    found = [g for g in list(LISTS) + list(DROP_VIEWS) if g in text.lower()]
    found += sorted(set(re.findall(r"sdguthrie[^\"' ]*", text)) - {"sdguthrie.com"})
    if CRS_ACCOUNT in text.replace(AUDIT_ACTOR, ""):
        found.append(CRS_ACCOUNT)
    return found


def main():
    os.makedirs(OUT, exist_ok=True)
    for src in sorted(glob.glob(os.path.join(SRC, "*.zip"))):
        out = os.path.join(OUT, "TEST_" + os.path.basename(src))
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.endswith("definition.json"):
                    text = port_definition(data.decode("utf8"), src)
                    left = leftovers(text)
                    data = text.encode("utf8")
                zout.writestr(info, data)
        print(("OK   " if not left else "LEFT ") + os.path.basename(out) + ("" if not left else "  " + str(left)))


if __name__ == "__main__":
    main()

"""Merges several legacy flow packages (.zip) into one, so one Import updates them all.
Usage: python scripts/merge-flow-packages.py <out.zip> <package1.zip> <package2.zip> ...
"""
import json
import sys
import zipfile

FLOWS_MANIFEST = "Microsoft.Flow/flows/manifest.json"


def main(out_path, paths):
    root = None
    asset_paths = []
    entries = {}
    for path in paths:
        with zipfile.ZipFile(path) as z:
            manifest = json.loads(z.read("manifest.json").decode("utf-8-sig"))
            if root is None:
                root = manifest
            else:
                root["resources"].update(manifest["resources"])
            asset_paths += json.loads(z.read(FLOWS_MANIFEST).decode("utf-8-sig"))["flowAssets"]["assetPaths"]
            for name in z.namelist():
                if name not in ("manifest.json", FLOWS_MANIFEST):
                    entries[name] = z.read(name)
    flows = [r for r in root["resources"].values() if r.get("type") == "Microsoft.Flow/flows"]
    root["details"]["displayName"] = "CRS flows for the test site (%d flows)" % len(flows)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as out:
        out.writestr("manifest.json", json.dumps(root))
        out.writestr(FLOWS_MANIFEST, json.dumps({"packageSchemaVersion": "1.0", "flowAssets": {"assetPaths": asset_paths}}))
        for name, data in entries.items():
            out.writestr(name, data)
    print("wrote %s: %d flows, %d resources" % (out_path, len(flows), len(root["resources"])))
    for r in flows:
        print("  -", r["details"]["displayName"])


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2:])

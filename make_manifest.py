"""Developer tool: regenerate resources/skeleton_manifest.json.

Run this after you change anything inside resources/skeleton/ (for example when the
WE branding is updated):

    python make_manifest.py

The app uses the manifest to check that every template file is present and unmodified
before it builds a Word document.
"""
import hashlib
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
SKELETON = os.path.join(HERE, "resources", "skeleton")
OUT = os.path.join(HERE, "resources", "skeleton_manifest.json")


def main():
    files = {}
    for root, _, names in os.walk(SKELETON):
        for name in names:
            full = os.path.join(root, name)
            rel = os.path.relpath(full, SKELETON).replace(os.sep, "/")
            with open(full, "rb") as f:
                files[rel] = hashlib.sha256(f.read()).hexdigest()
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump({"files": dict(sorted(files.items()))}, f, indent=1)
        f.write("\n")
    print(f"Wrote {OUT} ({len(files)} files)")


if __name__ == "__main__":
    main()

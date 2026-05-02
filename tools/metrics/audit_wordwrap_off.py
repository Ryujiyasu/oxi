"""
Scan baseline docx corpus for <w:wordWrap w:val="off"/> or val="0".
Report per-doc instance counts plus total. Used to scope the
break_into_lines CJK-gating fix at layout/mod.rs:4580.
"""
import os
import re
import sys
import zipfile
from collections import OrderedDict

DOCS_DIR = os.path.join(
    os.path.dirname(__file__), "..", "..", "tools", "golden-test", "documents", "docx"
)
DOCS_DIR = os.path.abspath(DOCS_DIR)

PAT = re.compile(rb'<w:wordWrap\s+w:val="(off|0)"\s*/?>')


def scan_one(path):
    try:
        with zipfile.ZipFile(path) as zf:
            count = 0
            for name in zf.namelist():
                if name.endswith(".xml"):
                    try:
                        data = zf.read(name)
                    except Exception:
                        continue
                    count += len(PAT.findall(data))
            return count
    except zipfile.BadZipFile:
        return -1


def main():
    rows = []
    for fn in sorted(os.listdir(DOCS_DIR)):
        if not fn.endswith(".docx"):
            continue
        cnt = scan_one(os.path.join(DOCS_DIR, fn))
        if cnt > 0:
            rows.append((fn, cnt))
    rows.sort(key=lambda r: -r[1])
    total = sum(c for _, c in rows)
    print(f"# wordWrap=off audit ({DOCS_DIR})")
    print(f"# docs with hits: {len(rows)}, total instances: {total}\n")
    for fn, cnt in rows:
        print(f"{cnt:5d}  {fn}")


if __name__ == "__main__":
    main()

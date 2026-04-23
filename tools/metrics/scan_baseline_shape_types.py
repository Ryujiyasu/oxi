"""Scan all baseline docx files for DrawingML / VML shape types present.

Purpose: identify what shape types Oxi must render beyond rect/roundRect/
bracketPair/straightConnector1/bentConnector3 to support forms and
office documents.
"""
import os
import re
import zipfile
from collections import Counter
from pathlib import Path

DOCX_DIR = "tools/golden-test/documents/docx"


def scan_docx(path: str) -> Counter:
    c = Counter()
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if not name.endswith(".xml"): continue
                try:
                    data = z.read(name)
                except Exception:
                    continue
                s = data.decode("utf-8", errors="ignore")
                # DrawingML preset geometries
                for m in re.finditer(r'prstGeom\s+prst="([^"]+)"', s):
                    c[f"DML:{m.group(1)}"] += 1
                # VML shape types (last segment of "#_x0000_t<NN>")
                for m in re.finditer(r'type="#_x0000_(t\d+)"', s):
                    c[f"VML:{m.group(1)}"] += 1
                # tailEnd / headEnd arrow styles
                for m in re.finditer(r'tailEnd[^/]*type="([^"]+)"', s):
                    c[f"tailEnd:{m.group(1)}"] += 1
                for m in re.finditer(r'headEnd[^/]*type="([^"]+)"', s):
                    c[f"headEnd:{m.group(1)}"] += 1
                # endarrow VML attr
                for m in re.finditer(r'endarrow="([^"]+)"', s):
                    c[f"VendArr:{m.group(1)}"] += 1
                # shape presets in oleObject preview
                # Custom geometry / curved paths
                if '<a:custGeom' in s:
                    c["DML:custGeom"] += 1
    except Exception as e:
        print(f"[NG] {path}: {e}")
    return c


def main():
    total = Counter()
    doc_count = Counter()
    files = sorted(Path(DOCX_DIR).glob("*.docx"))
    for p in files:
        c = scan_docx(str(p))
        total.update(c)
        for k in c:
            doc_count[k] += 1
    print(f"Scanned {len(files)} docx files\n")
    print(f"{'type':45s}  {'total':>8s}  {'docs':>5s}")
    print("-" * 65)
    for k, n in total.most_common():
        print(f"{k:45s}  {n:>8d}  {doc_count[k]:>5d}")


if __name__ == "__main__":
    main()

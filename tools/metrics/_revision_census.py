# -*- coding: utf-8 -*-
"""Which benchmark documents carry TRACKED CHANGES?

Word's default view for a docx with revisions is All Markup: deleted text is
still drawn (struck through) and insertions are underlined, so Word's page
count includes text that is not in the final document. Oxi renders the final
text only. Any pagination or SSIM comparison on such a document is therefore
measuring a missing FEATURE, not a layout law -- and it invents sites.

Found 2026-08-28 while dissecting `legal__0010437a7f75f636` (the pcd=-1 head of
the EN Phase-1 queue): its remaining "defect" sites are paragraphs where Word
draws deleted words (`proceedingproceedings`, `(4) to (6and (5)`) and so needs
one line more than Oxi.

    python tools/metrics/_revision_census.py            # every set
    python tools/metrics/_revision_census.py golden     # one set
"""
import json
import os
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "en_benchmark"

TAGS = ("w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange", "w:pPrChange",
        "w:tblPrChange", "w:trPrChange", "w:tcPrChange", "w:sectPrChange")


def counts(path):
    """Revision-element counts for one docx, over every body-ish part."""
    out = {}
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return {"ERROR": str(e)[:40]}
    for name in z.namelist():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        if not re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", name):
            continue
        try:
            x = z.read(name).decode("utf-8", "replace")
        except Exception:
            continue
        for t in TAGS:
            n = len(re.findall("<" + t + r"[ >]", x))
            if n:
                out[t] = out.get(t, 0) + n
    return out


def golden_docs():
    d = REPO / "tools" / "golden-test" / "documents" / "docx"
    return [(p.stem, p) for p in sorted(d.glob("*.docx")) if not p.name.startswith("~$")]


def en_docs():
    sets = {
        "first50": "_final.json", "next50": "_final_next50.json",
        "blind50": "_final_blind50.json", "blindB50": "_final_blindB50.json",
        "blindC50": "_final_blindC50.json",
    }
    out = []
    for sname, fn in sets.items():
        f = BENCH / fn
        if not f.exists():
            continue
        data = json.load(open(f, encoding="utf-8"))
        for kind, items in data.items():
            for it in items:
                p = Path(it["path"])
                out.append((f"{sname}/{kind}__{p.stem}", p))
    return out


which = sys.argv[1] if len(sys.argv) > 1 else "all"
groups = []
if which in ("all", "golden"):
    groups.append(("golden", golden_docs()))
if which in ("all", "en"):
    groups.append(("en_benchmark", en_docs()))

for gname, docs in groups:
    hits = []
    total = 0
    for name, path in docs:
        if not path.exists():
            continue
        total += 1
        c = counts(path)
        if c:
            hits.append((name, c))
    print(f"\n=== {gname}: {len(hits)} of {total} documents carry revisions ===")
    hits.sort(key=lambda kv: -sum(v for k, v in kv[1].items() if isinstance(v, int)))
    for name, c in hits:
        n = sum(v for v in c.values() if isinstance(v, int))
        body = " ".join(f"{k.split(':')[1]}={v}" for k, v in sorted(c.items()))
        print(f"   {n:6d}  {name:58s} {body}")

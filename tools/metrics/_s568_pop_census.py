# -*- coding: utf-8 -*-
"""Page counts for every doc that reaches the S568 legacy-oikomi gate.

S568 was derived when the gate admitted ONE document (harassmanual, compat 11).
The corpus now has 12, nine of them compat 14 -- a population the discriminator
was never tested against. Rather than cut the regime finer (S1234 -> S1236 ->
S1237, each a carve-out on the same gate), this asks the population question:
does the oikomi belong to compat 11 only?

Renders each gate doc under several arms and scores page count against Word.

Usage: _s568_pop_census.py [flag ...]     (default arms: S568, S1237)
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
GDI = str(REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe")
ARMS = sys.argv[1:] or ["OXI_S568_DISABLE", "OXI_S1237_DISABLE"]

WORD_SRC = [
    REPO / "pipeline_data" / "pagination_word",
    REPO / "pipeline_data" / "ja_benchmark" / "p1_blind50" / "word",
    REPO / "pipeline_data" / "ja_benchmark" / "p1_blindB50" / "word",
]


def word_pages(docx):
    stem = Path(docx).stem
    for d in WORD_SRC:
        if not d.exists():
            continue
        for f in d.glob("*.json"):
            if f.stem == stem or f.stem.endswith("__" + stem) or stem.startswith(f.stem):
                try:
                    return json.loads(f.read_text(encoding="utf-8"))["n_pages"]
                except Exception:
                    pass
    return None


def compat_of(docx):
    z = zipfile.ZipFile(docx)
    try:
        st = z.read("word/settings.xml").decode("utf-8", "replace")
    except KeyError:
        return None
    m = re.search(r'<w:compatSetting w:name="compatibilityMode"[^>]*w:val="(\d+)"', st)
    return int(m.group(1)) if m else None


def pages(docx, on):
    env = dict(os.environ)
    for a in ARMS:
        env.pop(a, None)
    for a in on:
        env[a] = "1"
    with tempfile.TemporaryDirectory(prefix="s568_") as t:
        dj = os.path.join(t, "l.json")
        r = subprocess.run([GDI, docx, os.path.join(t, "p"), "--dump-layout=" + dj],
                           capture_output=True, env=env, timeout=300)
        if r.returncode != 0 or not os.path.exists(dj):
            return None
        with open(dj, encoding="utf-8") as f:
            return len(json.load(f)["pages"])


# The gate docs, from _s568_gate_census.
GATE = []
for root in (REPO / "pipeline_data" / "docx_corpus" / "ja",
             REPO / "tools" / "golden-test" / "documents" / "docx"):
    for dp, _d, fs in os.walk(root):
        for f in fs:
            if not f.endswith(".docx") or f.startswith("~$"):
                continue
            p = os.path.join(dp, f)
            try:
                z = zipfile.ZipFile(p)
                doc = z.read("word/document.xml").decode("utf-8", "replace")
                st = z.read("word/settings.xml").decode("utf-8", "replace")
            except Exception:
                continue
            if not re.search(r'<w:docGrid[^>]*w:type="linesAndChars"', doc):
                continue
            m = re.search(r'<w:characterSpacingControl w:val="(\w+)"', st)
            if not m or not m.group(1).startswith("compressPunctuation"):
                continue
            c = compat_of(p)
            if c is None or c >= 15:
                continue
            GATE.append((c, p))

GATE.sort()
hdr = "%-34s %-7s %-6s %-8s" % ("doc", "compat", "word", "default")
for a in ARMS:
    hdr += " %-10s" % a.replace("OXI_", "").replace("_DISABLE", "-off")
print(hdr)
tot = {"default": 0}
for a in ARMS:
    tot[a] = 0
n = 0
for c, p in GATE:
    w = word_pages(p)
    base = pages(p, [])
    row = "%-34s %-7s %-6s %-8s" % (Path(p).stem[:34], c, w if w else "-", base)
    vals = {}
    for a in ARMS:
        v = pages(p, [a])
        vals[a] = v
        mark = ""
        if w:
            mark = "=" if v == w else ("%+d" % (v - w))
        row += " %-10s" % ("%s %s" % (v, mark))
    if w:
        n += 1
        tot["default"] += abs(base - w)
        for a in ARMS:
            tot[a] += abs(vals[a] - w)
        row += "  <<" if base != w else ""
    print(row)
print("\nsum|pcd| over %d scored docs:  default=%d  %s"
      % (n, tot["default"], "  ".join("%s=%d" % (a.replace("OXI_", ""), tot[a]) for a in ARMS)))

# -*- coding: utf-8 -*-
"""Who actually reaches the S568 legacy-oikomi gate?

S568's note says "the ONLY compat<15 linesAndChars compressPunctuation doc in
the corpus is harassmanual, so this is a single-doc-scoped change." That was
true of the dev corpus of 2026-06-14. jaBlindB50 joined dev on 2026-08-31, and
S1237 (a discriminator layered on this same gate) moves three of its docs.

Prints the four gate inputs plus the size regime, straight from the package:

    docGrid@type == linesAndChars
    settings.xml characterSpacingControl == compressPunctuation[AndJapaneseKana]
    compatSetting compatibilityMode < 15
    docDefaults sz (the S1236/S1237 regime reference) and the body run sizes

Usage: _s568_gate_census.py <docx|dir> ...
"""
import os
import re
import sys
import zipfile
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def read(z, name):
    try:
        return z.read(name).decode("utf-8", "replace")
    except KeyError:
        return ""


def probe(path):
    try:
        z = zipfile.ZipFile(path)
    except Exception:
        return None
    doc = read(z, "word/document.xml")
    st = read(z, "word/settings.xml")
    sty = read(z, "word/styles.xml")
    if not doc:
        return None

    grid = re.findall(r'<w:docGrid[^>]*w:type="(\w+)"', doc)
    lac = any(g == "linesAndChars" for g in grid)

    m = re.search(r'<w:characterSpacingControl w:val="(\w+)"', st)
    csc = m.group(1) if m else None
    cp = csc in ("compressPunctuation", "compressPunctuationAndJapaneseKana")

    m = re.search(r'<w:compatSetting w:name="compatibilityMode"[^>]*w:val="(\d+)"', st)
    compat = int(m.group(1)) if m else None

    m = re.search(r'<w:docDefaults>.*?<w:rPrDefault>.*?<w:sz w:val="(\d+)"', sty, re.S)
    dflt = int(m.group(1)) / 2.0 if m else None

    sizes = Counter(int(v) / 2.0 for v in re.findall(r'<w:sz w:val="(\d+)"/>', doc))
    return dict(lac=lac, csc=csc, cp=cp, compat=compat, dflt=dflt,
                sizes=sizes.most_common(4))


targets = []
for a in sys.argv[1:]:
    if os.path.isdir(a):
        for dp, _d, fs in os.walk(a):
            targets += [os.path.join(dp, f) for f in fs
                        if f.endswith(".docx") and not f.startswith("~$")]
    else:
        targets.append(a)

hits = 0
for p in sorted(targets):
    r = probe(p)
    if not r:
        continue
    gate = r["lac"] and r["cp"] and (r["compat"] is not None and r["compat"] < 15)
    if gate:
        hits += 1
        print("GATE-OPEN  %-40s compat=%-4s csc=%-32s dflt=%-5s sizes=%s"
              % (os.path.basename(p)[:40], r["compat"], r["csc"], r["dflt"], r["sizes"]))
print("\n%d of %d docs reach the S568 legacy-oikomi gate" % (hits, len(targets)))

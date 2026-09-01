# -*- coding: utf-8 -*-
"""Who reaches the S475 yakumono capacity break?

`s475_break`'s first branch fires on

    compress_punctuation && compat_mode >= 15 && !lines_and_chars && !natural_break_jc

i.e. every justified body paragraph of a compat-15 compressPunctuation document
whose docGrid is not linesAndChars. S592 later carved ONE exception out of it
(proportional CJK fonts, which are off-grid), but the gate still ASSUMES Word
demand-compresses the 約物 at break time.

`ohnoikuji_03` falsifies that assumption: Word's own PDF puts 40 characters on
the line at a per-char advance of 10.56..10.71 for a 10.5pt font -- larger than
the em, i.e. Word EXPANDED the line to justify it and compressed nothing. Oxi
credited a bracket pair 6.0pt of phantom compression and fitted a 41st char,
overrunning the text column by 5.3pt.

This counts the population the assumption is applied to, so the next probe can
sample it rather than guess. Same discipline as `_s568_gate_census.py`: a gate
justified by "the only doc in the corpus is X" needs an instrument that keeps
counting X.

Usage: _s475_gate_census.py <dir> [<dir> ...]
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
    if not doc:
        return None
    st = read(z, "word/settings.xml")
    sty = read(z, "word/styles.xml")

    m = re.search(r'<w:characterSpacingControl w:val="(\w+)"', st)
    cp = bool(m) and m.group(1).startswith("compressPunctuation")

    m = re.search(r'<w:compatSetting w:name="compatibilityMode"[^>]*w:val="(\d+)"', st)
    compat = int(m.group(1)) if m else None

    lac = bool(re.search(r'<w:docGrid[^>]*w:type="linesAndChars"', doc))

    # justified: either a doc-default/style jc=both, or paragraphs carrying it
    body_both = doc.count('<w:jc w:val="both"')
    style_both = sty.count('<w:jc w:val="both"')

    faces = Counter(re.findall(r'w:eastAsia="([^"]{1,16})"', doc))
    prop = any("Ｐ" in f or f.startswith("HGP") or f.startswith("HGSP")
               for f in faces)
    return dict(cp=cp, compat=compat, lac=lac, body_both=body_both,
                style_both=style_both, prop=prop, faces=faces.most_common(3))


targets = []
for a in sys.argv[1:]:
    if os.path.isdir(a):
        for dp, _d, fs in os.walk(a):
            targets += [os.path.join(dp, f) for f in fs
                        if f.endswith(".docx") and not f.startswith("~$")]
    else:
        targets.append(a)

hits = []
for p in sorted(targets):
    r = probe(p)
    if not r:
        continue
    gate = (r["cp"] and r["compat"] is not None and r["compat"] >= 15
            and not r["lac"] and (r["body_both"] or r["style_both"]))
    if gate:
        hits.append((p, r))

print("%-42s %-7s %-6s %-6s %s" % ("doc", "compat", "jc=both", "prop", "eastAsia faces"))
for p, r in hits:
    print("%-42s %-7s %-6s %-6s %s"
          % (os.path.basename(p)[:42], r["compat"],
             r["body_both"] or ("sty:%d" % r["style_both"]),
             "yes" if r["prop"] else "-",
             ", ".join("%s(%d)" % f for f in r["faces"])))
print("\n%d of %d docs reach the S475 capacity break (S592 excludes the proportional ones)"
      % (len(hits), len(targets)))

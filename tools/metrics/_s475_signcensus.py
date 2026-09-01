# -*- coding: utf-8 -*-
"""Does Word COMPRESS or EXPAND the line it justifies?  (the S475 premise test)

S475/S476 credit a justified CJK line with "demand 約物 compression": if the
line overruns, shrink its punctuation and fit one more character. That premise
was never checked against Word's own output for the population it fires on
(`_s475_gate_census.py`: 223 of 569 docs).

`ohnoikuji_03` falsifies it on the line that decides a page break -- Word puts
40 chars at a per-char advance of 10.56..10.71 for a 10.5pt font. The advance is
LARGER than the em: Word expanded the line and compressed nothing, while Oxi
credited a bracket pair 6.0pt and fitted a 41st char 5.3pt past the column.

Method (Word's PDF, rawdict, per character):
  full  = the line's dominant advance / font_size  -- >1 expanded, <1 compressed
  yaku  = each 約物's own advance / font_size      -- ~0.5 means halved
Only FULL lines are scored (a line ending well short of the column is the
paragraph's last line and is not justified).

Usage: _s475_signcensus.py <docx> <truth.pdf> [<docx> <truth.pdf> ...]
"""
import re
import sys
import zipfile
from collections import Counter

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

YAKU = "、。，．（）「」『』〔〕［］｛｝〈〉《》・：；？！"
# Characters JIS X 4051 forbids at the START of a line -- the ones an oikomi
# has a reason to pull back onto the previous line.
KINSOKU_START = "、。，．）」』〕］｝〉》・：；？！ぁぃぅぇぉっゃゅょァィゥェォッャュョーヽヾ々"


def is_cjk(c):
    """A FULL-WIDTH ideograph or kana -- the only chars whose advance is one em.

    A first version averaged every non-yakumono char and reported 269 of 352
    lines "compressed": ASCII advances half an em, so any line with a latin
    word in it fell below 1.0 no matter what Word did. The most-compressed
    line it found was the running header `モデル規程(2025.9)`. Mixing widths
    into one mean inverts the sign this instrument exists to read.
    """
    o = ord(c)
    return (0x4E00 <= o <= 0x9FFF or 0x3040 <= o <= 0x30FF
            or 0x3400 <= o <= 0x4DBF or 0xF900 <= o <= 0xFAFF)


def geometry(docx):
    z = zipfile.ZipFile(docx)
    d = z.read("word/document.xml").decode("utf-8", "replace")
    m = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"([^/]*)/>', d)
    w, h = int(m.group(1)) / 20.0, int(m.group(2)) / 20.0
    if "landscape" in m.group(3):
        w, h = max(w, h), min(w, h)
    mm = re.search(r"<w:pgMar ([^/]*)/>", d)
    a = dict(re.findall(r'w:(\w+)="(-?\d+)"', mm.group(1)))
    left = int(a.get("left", 1440)) / 20.0
    right = int(a.get("right", 1440)) / 20.0
    return w, left, w - right


def scan(docx, pdf):
    page_w, col_l, col_r = geometry(docx)
    doc = fitz.open(pdf)
    exp = comp = flat = 0
    buckets = {True: [], False: []}
    yaku_ratios = []
    worst = []
    for pi, pg in enumerate(doc):
        for b in pg.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    ch = s.get("chars") or []
                    if len(ch) < 12:
                        continue
                    fs = s["size"]
                    if fs <= 0:
                        continue
                    x0 = ch[0]["origin"][0]
                    xr = ch[-1]["bbox"][2]
                    # full line only: reaches within one em of the column edge
                    if xr < col_r - fs or x0 < col_l - 1.0:
                        continue
                    advs = [ch[i + 1]["origin"][0] - ch[i]["origin"][0]
                            for i in range(len(ch) - 1)]
                    if not advs:
                        continue
                    body = [a for a, c in zip(advs, ch)
                            if is_cjk(c["c"]) and c["c"] not in YAKU]
                    if len(body) < 8:
                        continue
                    ratio = (sum(body) / len(body)) / fs
                    # Does the line END on a char that may not START a line?
                    # Japanese oikomi exists to pull exactly such a char back;
                    # a line ending on an ordinary ideograph has nothing to
                    # pull, so Word should justify it by EXPANDING instead.
                    tail = "".join(c["c"] for c in ch).rstrip()
                    ends_kinsoku = bool(tail) and tail[-1] in KINSOKU_START
                    buckets[ends_kinsoku].append(ratio)
                    if ratio > 1.004:
                        exp += 1
                    elif ratio < 0.996:
                        comp += 1
                        worst.append((ratio, pi + 1, "".join(c["c"] for c in ch)[:26]))
                    else:
                        flat += 1
                    for a, c in zip(advs, ch):
                        if c["c"] in YAKU:
                            yaku_ratios.append(a / fs)
    doc.close()
    n = exp + comp + flat
    print("%-34s full lines=%-5d  EXPANDED=%-5d  compressed=%-4d  flat=%d"
          % (docx.split("\\")[-1].split("/")[-1], n, exp, comp, flat))
    if yaku_ratios:
        b = Counter()
        for r in yaku_ratios:
            b["%.2f" % (round(r * 20) / 20.0)] += 1
        print("      yaku advance / em :", ", ".join(
            "%s=%d" % kv for kv in sorted(b.items())[:8]))
        half = sum(1 for r in yaku_ratios if r < 0.75)
        print("      yaku halved (<0.75em): %d / %d = %.1f%%"
              % (half, len(yaku_ratios), 100.0 * half / len(yaku_ratios)))
    for k, name in ((True, "line ENDS on a kinsoku char"),
                    (False, "line ends on an ordinary char")):
        v = buckets[k]
        if not v:
            continue
        v2 = sorted(v)
        below = sum(1 for r in v if r < 0.996)
        print("      %-30s n=%-5d mean=%.4f  median=%.4f  compressed=%d (%.0f%%)"
              % (name, len(v), sum(v) / len(v), v2[len(v2) // 2], below,
                 100.0 * below / len(v)))
    worst.sort()
    for r, pi, t in worst[:3]:
        print("      most-compressed line ratio=%.4f p%d %r" % (r, pi, t))


args = sys.argv[1:]
for i in range(0, len(args) - 1, 2):
    scan(args[i], args[i + 1])

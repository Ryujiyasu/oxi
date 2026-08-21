# -*- coding: utf-8 -*-
"""Per-table audit of the FOOT gap in a real document.

For every top-level table, pair the LAST row's text with the text of the
paragraph that follows the table, then look both up in Word's PDF and in an Oxi
--dump-layout, and print Word's gap beside Oxi's (with and without a flag).
The point is to see WHICH tables Word charges its bottom rule to, since
`_pb_tblfoot_gen.py` says it always does and forms__0020466f says otherwise.

  python tools/metrics/_tblfoot_audit.py <docx> [FLAG=1]
"""
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
DOCX = Path(sys.argv[1]).resolve()
FLAG = sys.argv[2] if len(sys.argv) > 2 else None
PDF = Path(tempfile.gettempdir()) / (DOCX.name + ".truth.pdf")


def norm(s):
    return re.sub(r"[^0-9A-Za-z]", "", s)[:22]


def tables():
    """(idx, last-row text, following paragraph text, bottom decl, lastrow decls)."""
    x = zipfile.ZipFile(DOCX).read("word/document.xml").decode("utf8", "replace")
    out = []
    # top-level tables only: split the body on depth-0 <w:tbl>
    depth = 0
    spans = []
    start = None
    for m in re.finditer(r"<w:tbl>|</w:tbl>|<w:tc>|<w:tc [^>]*>|</w:tc>", x):
        g = m.group(0)
        if g.startswith("<w:tc"):
            depth += 1
        elif g == "</w:tc>":
            depth -= 1
        elif g == "<w:tbl>":
            if depth == 0:
                start = m.start()
        elif g == "</w:tbl>" and depth == 0 and start is not None:
            spans.append((start, m.end()))
            start = None
    for i, (a, b) in enumerate(spans):
        t = x[a:b]
        rows = re.findall(r"<w:tr[ >].*?</w:tr>", t, re.S)
        if not rows:
            continue
        last = "".join(re.findall(r"<w:t(?: [^>]*)?>(.*?)</w:t>", rows[-1], re.S))
        after = ""
        for mp in re.finditer(r"<w:p[ >].*?</w:p>", x[b:b + 8000], re.S):
            cand = "".join(re.findall(r"<w:t(?: [^>]*)?>(.*?)</w:t>", mp.group(0), re.S))
            if cand.strip():
                after = cand
                break
        tb = re.search(r"<w:tblBorders>(.*?)</w:tblBorders>", t, re.S)
        bot = None
        if tb:
            mb = re.search(r'<w:bottom w:val="(\w+)"(?: w:sz="(\d+)")?', tb.group(1))
            if mb:
                bot = (mb.group(1), int(mb.group(2) or 0))
        cells = re.findall(r"<w:tc>.*?</w:tc>|<w:tc [^>]*>.*?</w:tc>", rows[-1], re.S)
        decls = []
        for c in cells:
            mc = re.search(r"<w:tcBorders>(.*?)</w:tcBorders>", c, re.S)
            d = None
            if mc:
                md = re.search(r'<w:bottom w:val="(\w+)"(?: w:sz="(\d+)")?', mc.group(1))
                if md:
                    d = (md.group(1), int(md.group(2) or 0))
            decls.append(d)
        out.append((i, last.strip(), after.strip(), bot, decls))
    return out


def word_map():
    if not PDF.exists() or "--reexport" in sys.argv:
        import win32com.client as win32
        w = win32.DispatchEx("Word.Application")
        w.Visible = False
        try:
            d = w.Documents.Open(str(DOCX), ReadOnly=True)
            d.ExportAsFixedFormat(str(PDF), 17)
            d.Close(False)
        finally:
            w.Quit()
    import fitz
    doc = fitz.open(PDF)
    m = {}
    for pi in range(doc.page_count):
        for blk in doc[pi].get_text("dict")["blocks"]:
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    m.setdefault(norm(t), []).append(
                        pi * 10000 + min(s["bbox"][1] for s in ln["spans"]))
    return m


def oxi_map(flag):
    exe = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
    tmp = Path(tempfile.mkdtemp())
    dump = tmp / "d.json"
    env = dict(os.environ)
    env.pop("OXI_S1191", None)
    if flag:
        k, _, v = flag.partition("=")
        env[k] = v or "1"
    subprocess.run([str(exe), str(DOCX), str(tmp / "p"), "110", "--dump-layout=%s" % dump],
                   check=True, capture_output=True, env=env)
    import json
    d = json.load(open(dump, encoding="utf-8"))
    m = {}
    for pi, pg in enumerate(d["pages"]):
        rows = {}
        for e in pg["elements"]:
            if e.get("type") == "text" and (e.get("text") or "").strip():
                rows.setdefault(round(e.get("y", 0.0), 2), []).append((e.get("x", 0.0), e["text"]))
        for y, v in rows.items():
            t = "".join(s for _, s in sorted(v)).strip()
            if t:
                m.setdefault(norm(t), []).append(pi * 10000 + y)
    return m


W = word_map()
A = oxi_map(None)
B = oxi_map(FLAG) if FLAG else None
print("%-4s %-9s %-9s %-9s  %-16s %s" % ("tbl", "word_gap", "oxi_off", "oxi_flag",
                                         "tbl_bottom", "lastrow cell bottoms"))
for i, last, after, bot, decls in tables():
    kl, ka = norm(last), norm(after)
    if len(kl) < 6 or len(ka) < 6 or kl not in W or ka not in W:
        continue

    def gap(m):
        if kl not in m or ka not in m:
            return None
        for yl in sorted(m[kl]):
            nxt = [y for y in m[ka] if y > yl]
            if nxt:
                return min(nxt) - yl
        return None
    gw, ga = gap(W), gap(A)
    gb = gap(B) if B else None
    if gw is None or ga is None or gw > 400:
        continue
    print("%-4d %-9.2f %-9s %-9s  %-16s %s"
          % (i, gw, "%.2f" % ga if ga is not None else "-",
             "%.2f" % gb if gb is not None else "-", str(bot), str(decls)[:60]))

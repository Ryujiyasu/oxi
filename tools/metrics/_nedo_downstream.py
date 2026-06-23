# -*- coding: utf-8 -*-
"""Localize the COMPENSATING downstream over-fit (the -1x3 exposed by fixing para
333): a nedo para where OXI has FEWER display lines than WORD (Oxi oikomi where
Word oidashi — OPPOSITE of 333). Robust difflib char-stream alignment (the raw
monotonic anchor breaks; para_idx != word_i). Builds the full Word-PDF char stream
(line boundaries tagged) and the full Oxi-dump char stream (para boundaries tagged),
aligns via SequenceMatcher, then per Oxi para counts the Word LINES its chars span.
"""
import json, sys, fitz, difflib
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")

PDF = r"C:\tmp\nedocontract_word.pdf"
HEADER = "一般再委託用"


def word_stream():
    """Full Word char stream; per char store (page, line_id). Skip the running header."""
    doc = fitz.open(PDF)
    chars = []          # list of char
    meta = []           # parallel (page, line_id)
    lid = 0
    for pi in range(doc.page_count):
        d = doc.load_page(pi).get_text("rawdict")
        lines = []
        for blk in d["blocks"]:
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                cs = []
                for sp in ln.get("spans", []):
                    for c in sp.get("chars", []):
                        cs.append((c["c"], c["bbox"][0]))
                if cs:
                    cs.sort(key=lambda t: t[1])
                    lines.append((round(ln["bbox"][1], 1), [c[0] for c in cs]))
        lines.sort()
        for _, cl in lines:
            txt = "".join(cl).strip()
            if txt.startswith(HEADER) or not txt:
                continue
            lid += 1
            for ch in cl:
                if ch.strip() and ch not in ("　", " "):
                    chars.append(ch)
                    meta.append((pi + 1, lid))
    doc.close()
    return chars, meta


def oxi_stream():
    """Full Oxi char stream; per char store para_idx. Also para -> nlines, first-line."""
    d = json.load(open("C:/tmp/n_def.json", encoding="utf-8"))
    rows = defaultdict(lambda: defaultdict(list))
    for pg in d["pages"]:
        for e in pg["elements"]:
            if e.get("type") == "text" and e.get("para_idx") is not None:
                rows[e["para_idx"]][round(e["y"], 1)].append(e)
    chars, meta = [], []
    info = {}
    for pi in sorted(rows):
        ys = sorted(rows[pi])
        first = "".join(c["text"] for c in sorted(rows[pi][ys[0]], key=lambda c: c["x"]))
        info[pi] = {"nlines": len(ys), "first": first}
        for y in ys:
            for c in sorted(rows[pi][y], key=lambda c: c["x"]):
                ch = c["text"]
                if ch.strip() and ch not in ("　", " "):
                    chars.append(ch)
                    meta.append(pi)
    return chars, meta, info


def main():
    wch, wmeta = word_stream()
    och, ometa, oinfo = oxi_stream()
    sm = difflib.SequenceMatcher(None, och, wch, autojunk=False)
    # per oxi para: set of word line ids its matched chars fall on
    para_wlines = defaultdict(set)
    for a, b, size in sm.get_matching_blocks():
        for k in range(size):
            pi = ometa[a + k]
            wpage, wlid = wmeta[b + k]
            para_wlines[pi].add(wlid)
            oinfo[pi].setdefault("wpage", wpage)
    print(f"{'oxiPI':>5} {'wpg':>3} {'Wln':>4} {'Oxiln':>5}  first-line")
    for pi in sorted(oinfo):
        wp = oinfo[pi].get("wpage")
        if wp is None or not (19 <= wp <= 27):
            continue
        wln = len(para_wlines.get(pi, ()))
        oln = oinfo[pi]["nlines"]
        mark = ""
        if oln < wln:
            mark = f"  <<< OXI FEWER {oln}<{wln}  OVER-FIT (compensating?)"
        elif oln > wln:
            mark = f"  (oxi more {oln}>{wln})"
        print(f"{pi:>5} {wp:>3} {wln:>4} {oln:>5}  {oinfo[pi]['first'][:32]}{mark}")


if __name__ == "__main__":
    main()

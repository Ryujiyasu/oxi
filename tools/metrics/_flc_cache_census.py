# -*- coding: utf-8 -*-
"""firstLineChars: does Word draw the cached w:firstLine or recompute from chars?

For single-section corpus docs whose cached firstLine disagrees with
chars x (sz + 2 x spacing) by >= 1.5pt, render the Word PDF, find the body
paragraph's first glyph x, and print the drawn indent next to the candidates
(cache / chars x sz / chars x (sz + 2sp) / + docGrid extra).

    python tools/metrics/_flc_cache_census.py [max_docs]
"""
import glob, html, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
PDFD = REPO / "pipeline_data/jshrink_corpus/pdf"
MAXD = int(sys.argv[1]) if len(sys.argv) > 1 else 8


def cases():
    out = []
    for p in sorted(glob.glob(str(REPO / "pipeline_data/docx_corpus/*/*/*.docx"))):
        try:
            z = zipfile.ZipFile(p); d = z.read("word/document.xml").decode("utf-8", "ignore")
            st = z.read("word/styles.xml").decode("utf-8", "ignore")
        except Exception:
            continue
        if "firstLineChars" not in d or d.count("<w:sectPr") != 1:
            continue
        sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", d, re.S).group(0)
        pm = re.search(r"<w:pgMar [^>]*/>", sect); pgw = re.search(r'<w:pgSz [^>]*w:w="(\d+)"', sect)
        if not pm or "<w:cols " in sect and 'w:num="' in sect:
            continue
        left = int(re.search(r'w:left="(\d+)"', pm.group(0)).group(1)) / 20
        gut = re.search(r'w:gutter="(\d+)"', pm.group(0)); left += int(gut.group(1)) / 20 if gut else 0
        grid = re.search(r"<w:docGrid [^>]*/>", sect)
        gtype = re.search(r'w:type="(\w+)"', grid.group(0)) if grid else None
        cs = re.search(r'w:charSpace="(-?\d+)"', grid.group(0)) if grid else None
        dd = re.search(r"<w:docDefaults>.*?</w:docDefaults>", st, re.S)
        dsz = re.search(r'<w:sz w:val="(\d+)"', dd.group(0)) if dd else None
        dsz = int(dsz.group(1)) / 2 if dsz else 10.5
        for m in re.finditer(r"<w:p [^>]*>.*?</w:p>", d, re.S):
            para = m.group(0)
            if d.rfind("<w:tbl>", 0, m.start()) > d.rfind("</w:tbl>", 0, m.start()):
                continue
            ind = re.search(r"<w:ind [^>]*/>", para)
            if not ind or "firstLineChars" not in ind.group(0) or "hangingChars" in ind.group(0) or "left" in ind.group(0) or "numPr" in para or "<w:tab/>" in para or "<w:drawing" in para or "<w:pict" in para:
                continue
            fc = re.search(r'firstLineChars="(-?\d+)"', ind.group(0)); fl = re.search(r'firstLine="(-?\d+)"', ind.group(0))
            if not (fc and fl) or int(fc.group(1)) <= 0:
                continue
            jc = re.search(r'<w:jc w:val="(\w+)"', para)
            if jc and jc.group(1) != "left":
                continue
            body = re.sub(r"<w:rt>.*?</w:rt>", "", para, flags=re.S)
            runs = [r for r in re.findall(r"<w:r[ >].*?</w:r>", body, re.S) if re.search(r"<w:t(?:\s[^>]*)?>", r)]
            if not runs:
                continue
            text = html.unescape("".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", body)))
            if len(text.strip()) < 4 or text[0] in " 　":
                continue
            rpr = re.search(r"<w:rPr>.*?</w:rPr>", runs[0], re.S); rpr = rpr.group(0) if rpr else ""
            sz = re.search(r'<w:sz w:val="(\d+)"', rpr); sp = re.search(r'<w:spacing w:val="(-?\d+)"', rpr)
            szp = int(sz.group(1)) / 2 if sz else dsz; spp = int(sp.group(1)) / 20 if sp else 0.0
            chars = int(fc.group(1)) / 100
            cache = int(fl.group(1)) / 20
            if abs(cache - chars * (szp + 2 * spp)) < 1.5 and abs(cache - chars * szp) < 1.5:
                continue
            out.append(dict(doc=p, left=left, text=text, chars=chars, cache=cache, sz=szp, sp=spp,
                            grid=(gtype.group(1) if gtype else None, int(cs.group(1)) / 20 if cs else None),
                            fit="<w:fitText" in rpr, key=Path(p).parts[-3] + "/" + Path(p).parts[-2] + "/" + Path(p).stem[:8]))
    return out


def main():
    import fitz, win32com.client
    cs = cases()
    docs = []
    for c in cs:
        if c["doc"] not in docs:
            docs.append(c["doc"])
    docs = docs[:MAXD]
    print(f"{len(cs)} candidate paragraphs in {len({c['doc'] for c in cs})} single-section docs; measuring {len(docs)} docs")
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for p in docs:
            key = next(c["key"] for c in cs if c["doc"] == p)
            pdf = PDFD / (key.replace("/", "__") + ".pdf")
            if not pdf.exists():
                tmp = os.path.join(tempfile.mkdtemp(), "t.docx"); shutil.copy(p, tmp)
                try:
                    doc = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False); doc.SaveAs2(str(pdf), 17); doc.Close(0)
                except Exception as e:
                    print("ERR", key, e); continue
            glyphs = []
            for pno, page in enumerate(fitz.open(pdf)):
                for b in page.get_text("rawdict")["blocks"]:
                    for l in b.get("lines", []):
                        chars = [c for s in l["spans"] for c in s["chars"]]
                        if chars:
                            glyphs.append(("".join(c["c"] for c in chars), chars[0]["origin"][0], pno + 1))
            for c in [c for c in cs if c["doc"] == p][:4]:
                probe = c["text"][:8]
                hit = next((g for g in glyphs if g[0].startswith(probe)), None)
                if not hit:
                    hit = next((g for g in glyphs if probe in g[0]), None)
                    if hit:
                        hit = None  # the line starts elsewhere: not the paragraph's first line
                drawn = (hit[1] - c["left"]) if hit else None
                f1, f2 = c["chars"] * c["sz"], c["chars"] * (c["sz"] + 2 * c["sp"])
                print(f"{c['key']} p{hit[2] if hit else '?'} chars={c['chars']:.2f} sz={c['sz']} sp={c['sp']} grid={c['grid']} fit={c['fit']} | "
                      f"drawn={drawn if drawn is None else round(drawn, 2)}  cache={c['cache']:.2f}  chars*sz={f1:.2f}  chars*(sz+2sp)={f2:.2f} | {c['text'][:14]!r}")
    finally:
        w.Quit()


if __name__ == "__main__":
    main()

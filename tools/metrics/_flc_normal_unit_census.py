# -*- coding: utf-8 -*-
"""firstLineChars cache written at the NORMAL size: does Word draw it?

The settle step (OXI_STYLE_IND_CHARS) keeps the cached w:firstLine only while it lies within a
quarter unit of chars x the paragraph's unit (its mark/run size), else it recomputes from chars.
blind-J ja/reports__7a8daf36 stores firstLineChars=3650 firstLine=7665 (= 36.5 x 10.5, the Normal
size) on a 12pt paragraph: Word draws 383.25 (the cache), Oxi 438 (36.5 x 12) and wraps «n= 64».
This lists every corpus paragraph whose cache equals chars x the Normal size but is more than a
quarter unit off chars x its own unit, and reads the drawn first-line indent from the Word PDF.

    python _flc_normal_unit_census.py
"""
import glob, os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
REAL = r"C:/Users/ryuji/oxi"


def pdf_for(stem):
    for pat in ("pipeline_data/*_benchmark/*/word_pdf/*%s*.pdf", "pipeline_data/jshrink_corpus/pdf/*%s*.pdf"):
        g = glob.glob(os.path.join(REAL, pat % stem))
        if g:
            return g[0]
    return None


def text_of(xml):
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", xml))


def main():
    import fitz
    for p in sorted(glob.glob(os.path.join(REAL, "pipeline_data/docx_corpus/ja/*/*.docx"))):
        try:
            z = zipfile.ZipFile(p)
            d = z.read("word/document.xml").decode("utf8", "ignore")
            st = z.read("word/styles.xml").decode("utf8", "ignore")
        except Exception:
            continue
        if "firstLineChars" not in d:
            continue
        m = re.search(r'<w:style [^>]*w:type="paragraph"[^>]*w:default="1"[^>]*>.*?</w:style>', st, re.S)
        nsz = re.search(r'<w:sz w:val="(\d+)"', m.group(0)) if m else None
        dd = re.search(r"<w:docDefaults>.*?</w:docDefaults>", st, re.S)
        dsz = re.search(r'<w:sz w:val="(\d+)"', dd.group(0)) if dd else None
        normal = int(nsz.group(1)) / 2 if nsz else (int(dsz.group(1)) / 2 if dsz else 10.5)
        sects = re.findall(r"<w:sectPr\b.*?</w:sectPr>", d, re.S)
        if len(sects) != 1:
            continue
        pm_ = re.search(r'<w:pgMar [^>]*w:left="(\d+)"', sects[0])
        margin = int(pm_.group(1)) / 20 if pm_ else 0
        stem = os.path.basename(p)[:-5]
        pdf = None
        for para in re.findall(r"<w:p[ >].*?</w:p>", d, re.S):
            if "<w:tbl" in para or "txbxContent" in para:
                continue
            ppr = para.split("</w:pPr>")[0]
            ind = re.search(r"<w:ind [^>]*/>", ppr)
            if not ind or "firstLineChars" not in ind.group(0):
                continue
            c = re.search(r'w:firstLineChars="(\d+)"', ind.group(0))
            fl = re.search(r'w:firstLine="(\d+)"', ind.group(0))
            if not c or not fl or int(c.group(1)) == 0:
                continue
            c = int(c.group(1)) / 100
            t = int(fl.group(1)) / 20
            jc = re.search(r'<w:jc w:val="(\w+)"', ppr)
            if jc and jc.group(1) in ("center", "right", "end"):
                continue
            msz = re.search(r"<w:rPr>.*?<w:sz w:val=\"(\d+)\"", ppr, re.S)
            unit = int(msz.group(1)) / 2 if msz else normal
            if not (abs(t - c * normal) < 0.06 and abs(t - c * unit) > 0.25 * unit):
                continue
            left = re.search(r'w:left="(-?\d+)"', ind.group(0))
            left = int(left.group(1)) / 20 if left else 0
            txt = text_of(para).replace("\u3000", "").strip()
            if len(txt) < 3:
                continue
            pdf = pdf or pdf_for(stem)
            if not pdf:
                print(stem, "no Word PDF")
                break
            key = txt[:4]
            x0 = None
            for pg in fitz.open(pdf):
                for b in pg.get_text("dict")["blocks"]:
                    for l in b.get("lines", []):
                        lt = "".join(s["text"] for s in l["spans"]).replace("\u3000", "").replace(" ", "")
                        if lt.startswith(key.replace(" ", "")):
                            x0 = l["bbox"][0]
                            break
                    if x0 is not None:
                        break
                if x0 is not None:
                    break
            drawn = None if x0 is None else round(x0 - margin - left, 2)
            print("%-34s chars %5.2f cache %7.2f  chars x unit(%4.1f) %7.2f  drawn %s  -> %s" % (
                stem, c, t, unit, c * unit, drawn,
                "?" if drawn is None else ("CACHE" if abs(drawn - t) < abs(drawn - c * unit) else "CHARSxUNIT")))


if __name__ == "__main__":
    main()

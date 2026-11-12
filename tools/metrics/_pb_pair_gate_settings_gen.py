# -*- coding: utf-8 -*-
"""Which setting halves a punctuation pair WITHOUT w:kern?

S547 (kern 0 -> no halving) re-measured by PDF on the current Word: holds.  Yet the slice of
ja/policies__07543a6b halves 」） ）」 、「 with every w:kern removed.  That document's compat
block carries useFELayout and compatSetting enableOpenTypeFeatures=1 (and
overrideTableStyleFontSizeAndJustification / doNotFlipMirrorIndents /
differentiateMultirowTableHeaders).  Arms: S547's kern-0 compat-15 doc (+ each setting),
and the slice's no-kern arm (- enableOpenTypeFeatures).  Read: advance of the first mark of
«、（» / «（「» (S547 docs, 10.5pt) and «」）» (slice, 12pt) from Word's PDF.

    python _pb_pair_gate_settings_gen.py gen | word | oxi EXE
"""
import json, os, re, subprocess, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_pair_gate_settings")
S547 = r"C:/Users/ryuji/oxi/tools/golden-test/repros/s547_kern/s547b_k0_c0_v1.docx"
SLICE = os.path.join(REPO, "pipeline_data", "_pb_cell_pair_kern", "body_nokern.docx")
OT = '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
ARMS = {
    "s547_k0": (S547, None),
    "s547_k0_ot": (S547, ("</w:compat>", OT + "</w:compat>")),
    "s547_k0_fe": (S547, ("<w:compat>", "<w:compat><w:useFELayout/>")),
    "s547_k0_sampler": (S547, None, ("国、（国国", "投」）落、「投。」落）」投）、落、（国")),
    "s547_k0_sampler12": (S547, None, ("国、（国国", "投」）落、「投。」落）」投）、落、（国"), ('<w:sz w:val="21"/>', '<w:sz w:val="24"/>')),
    "s547_k0_sampler_zh": (S547, ("</w:compat>", '</w:compat><w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>'), ("国、（国国", "投」）落、「投。」落）」投）、落、（国")),
    "s547_k0_sampler_ja": (S547, ("</w:compat>", '</w:compat><w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>'), ("国、（国国", "投」）落、「投。」落）」投）、落、（国")),
    "slice_nokern_tja": (SLICE, ('w:eastAsia="zh-CN"/>', 'w:eastAsia="ja-JP"/>')),
    "slice_kern_tja": (os.path.join(REPO, "pipeline_data", "_pb_cell_pair_kern", "body.docx"), ('w:eastAsia="zh-CN"/>', 'w:eastAsia="ja-JP"/>')),
    "slice_nokern_s547styles": (SLICE, None),
    "slice_nokern_s547settings": (SLICE, None),
    "slice_nokern": (SLICE, None),
    "slice_nokern_noot": (SLICE, (OT, "")),
}


PARTSWAP = {
    # slice with one whole part taken from S547's kern-0 doc
    "slice_nokern_s547styles": "word/styles.xml",
    "slice_nokern_s547settings": "word/settings.xml",
}


def gen():
    os.makedirs(OUT, exist_ok=True)
    for a, (src, rep, *more) in ARMS.items():
        z = zipfile.ZipFile(src)
        with zipfile.ZipFile(os.path.join(OUT, a + ".docx"), "w", zipfile.ZIP_DEFLATED) as o:
            for n in z.namelist():
                b = z.read(n)
                if n == "word/settings.xml" and rep:
                    t = b.decode("utf8")
                    assert t.count(rep[0]) == 1, (a, rep[0])
                    b = t.replace(rep[0], rep[1]).encode("utf8")
                if PARTSWAP.get(a) == n:
                    b = zipfile.ZipFile(S547).read(n)
                for old, new in more:
                    if n in ("word/document.xml", "word/styles.xml") and old.encode("utf8") in b:
                        b = b.replace(old.encode("utf8"), new.encode("utf8"), 1)
                o.writestr(n, b)
    print("wrote", list(ARMS))


def pairs(pdf):
    import fitz
    out = {}
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [(c["c"], c["bbox"][0]) for s in l["spans"] for c in s["chars"]]
                for i in range(len(ch) - 1):
                    k = ch[i][0] + ch[i + 1][0]
                    if k in ("、（", "（「", "」）", "）」", "、「", "。」", "）、"):
                        out.setdefault(k, round(ch[i + 1][1] - ch[i][1], 2))
    return out


def word():
    import win32com.client as w
    res = {}
    for a in ARMS:
        if "--only" in sys.argv and a not in sys.argv:
            continue
        app = w.DispatchEx("Word.Application"); app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, a + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, a + "_word.pdf"), 17); d.Close(False)
        finally:
            app.Quit()
        res[a] = pairs(os.path.join(OUT, a + "_word.pdf")); print(a, res[a], flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf8"), ensure_ascii=False, indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf8"))
    for a in ARMS:
        pdf = os.path.join(OUT, a + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, a + ".docx"), pdf], capture_output=True)
        print("%-18s word %s\n%-18s oxi  %s" % (a, want[a], "", pairs(pdf)))


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()

# -*- coding: utf-8 -*-
"""Does a whitespace-only run in a larger font raise the line?

blind-H en/forms__00613e6b: footer paragraph = a bare « » run (style font
Times New Roman 12) followed by Arial 10 bold text. Word's line is 13.8
(the TNR-12 box); Oxi's is 11.5 (Arial 10 only), so the footer's first line
sits 2.36pt high.

Arms: lead (space run TNR12 then Arial10 text), trail (text then the space
run), mid (text, space run, text), tab (a TNR12 tab run then text), none
(Arial 10 only), big (lead space run at 20pt).  Each arm = three identical
paragraphs «Ann / Bnn / Cnn»; read the baseline pitch.

    python tools/metrics/_pb_wsrun_lh_gen.py gen
    python tools/metrics/_pb_wsrun_lh_gen.py word
    python tools/metrics/_pb_wsrun_lh_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_wsrun_lh"
ARIAL = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="20"/></w:rPr>'
TNR = lambda sz=24: f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="{sz}"/></w:rPr>'
ARMS = ["lead", "trail", "mid", "tab", "none", "big"]


def runs(arm, label):
    txt = lambda t: f'<w:r>{ARIAL}<w:t xml:space="preserve">{t}</w:t></w:r>'
    sp = lambda sz=24: f'<w:r>{TNR(sz)}<w:t xml:space="preserve"> </w:t></w:r>'
    return {
        "lead": sp() + txt(f"{label} Somerset Adult ADHD Service"),
        "trail": txt(f"{label} Somerset Adult ADHD Service") + sp(),
        "mid": txt(f"{label} Somerset Adult") + sp() + txt("ADHD Service"),
        "tab": f'<w:r>{TNR()}<w:tab/></w:r>' + txt(f"{label} Somerset Adult ADHD Service"),
        "none": txt(f"{label} Somerset Adult ADHD Service"),
        "big": sp(40) + txt(f"{label} Somerset Adult ADHD Service"),
    }[arm]


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    styles = T.STYLES.replace('w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic" w:cs="MS Gothic"',
                              'w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"')
    for arm in ARMS:
        paras = "".join(f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{ARIAL}</w:pPr>'
                        f'{runs(arm, lab)}</w:p>' for lab in ("A01", "B01", "C01"))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + paras +
               '<w:sectPr><w:pgSz w:w="11909" w:h="16834"/><w:pgMar w:top="743" w:right="964" w:bottom="1474"'
               ' w:left="964" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", styles)
            z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _pitch(lines):
    ys = sorted(l["y"] for l in lines if l["t"][:3] in ("A01", "B01", "C01") or l["t"].lstrip()[:3] in ("A01", "B01", "C01"))
    return round((ys[-1] - ys[0]) / 2, 2) if len(ys) == 3 else None


def word():
    import win32com.client, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        res[arm] = _pitch(L.word_lines(OUT / f"{arm}.pdf"))
        print(f"{arm:6s} word pitch {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        p = _pitch(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        flag = "" if p is not None and word.get(arm) is not None and abs(p - word[arm]) < 0.3 else "  <-- differs"
        print(f"{arm:6s} oxi pitch {p}  word {word.get(arm)}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()

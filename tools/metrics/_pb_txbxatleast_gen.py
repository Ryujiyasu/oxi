# -*- coding: utf-8 -*-
"""A text box whose paragraphs carry `line=0 atLeast` (and bodyPr upright="1",
insets 74295/8890) on a lines docGrid: where do Word's first two baselines sit?

blind-G ja/correspondence__139e74eb: the title box (355x80.5, posV paragraph
+4.8) holds «自立活動通信» (HGS創英角ﾎﾟｯﾌﾟ体 32, line=0 atLeast, centred) and a
10.5 reading line; Word baselines 93.62 / 117.02, Oxi 86.90 / 105.28.

Arms are faithful slices: the document's own anchor XML (sectPr too), with
regex variants -- orig / noatleast (spacing removed) / noupright / definsets
(lIns 91440 tIns 45720) / lines0 (no docGrid). Read: baseline − box top.

    python tools/metrics/_pb_txbxatleast_gen.py gen
    python tools/metrics/_pb_txbxatleast_gen.py word
    python tools/metrics/_pb_txbxatleast_gen.py oxi <renderer.exe>
"""
import json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import feature_census as FC  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_txbxatleast"
ARMS = ["orig", "noatleast", "noupright", "definsets", "nogrid", "nogrid_noatleast", "nofittext", "nospacing", "plain", "nogrid_plain"]
BOX_TOP = 45.35 + 4.8  # top margin 907tw + posV paragraph offset


def src_doc():
    d = FC.load()
    return d[[k for k in d if "139e74eb" in k][0]]["path"]


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    z = zipfile.ZipFile(src_doc())
    x = z.read("word/document.xml").decode("utf-8")
    # the run holding the title box: back to its start, forward over the nested
    # AlternateContent elements to its end
    i = x.find("つーしん")  # the reading line is unique to the title box
    a = x.rfind("<mc:AlternateContent>", 0, i)
    start = max(x.rfind("<w:r>", 0, a), x.rfind("<w:r ", 0, a))
    assert i > 0 and start > 0, "title box not found"
    depth, pos = 0, start
    while True:
        o = x.find("<mc:AlternateContent>", pos); c = x.find("</mc:AlternateContent>", pos)
        if o != -1 and o < c:
            depth += 1; pos = o + 1
        else:
            depth -= 1; pos = c + len("</mc:AlternateContent>")
            if depth == 0:
                break
    end = x.find("</w:r>", pos) + len("</w:r>")
    run = x[start:end]
    assert "自立活動" in run and "つーしん" in run, "title box not found"
    sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", x, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    styles = z.read("word/styles.xml")
    theme = z.read("word/theme/theme1.xml") if "word/theme/theme1.xml" in z.namelist() else None
    for arm in ARMS:
        r, s = run, sect
        if "noatleast" in arm:
            r = r.replace('<w:spacing w:line="0" w:lineRule="atLeast"/>', "")
        if arm == "noupright":
            r = r.replace(' upright="1"', "")
        if arm == "definsets":
            r = r.replace('lIns="74295" tIns="8890" rIns="74295" bIns="8890"', 'lIns="91440" tIns="45720" rIns="91440" bIns="45720"')
        if arm in ("nofittext", "plain", "nogrid_plain"):
            r = re.sub(r"<w:fitText [^>]*/>", "", r)
        if arm in ("nospacing", "plain", "nogrid_plain"):
            r = r.replace('<w:spacing w:val="192"/>', "")
        if arm.startswith("nogrid"):
            s = re.sub(r"<w:docGrid [^>]*/>", "", s)
        body = "<w:p>" + r + "</w:p>"
        # the source document's own root: its namespace declarations cover the
        # run's wp14 / w14 attributes (a minimal root makes Word call the file corrupt)
        head = x[: x.find("<w:body>") + len("<w:body>")]
        doc = head + body + s + "</w:body></w:document>"
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          + ('<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' if theme else "") + "</Types>")
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                                + ('<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' if theme else "") + "</Relationships>")
        # the source settings part makes Word call the slice corrupt; keep only its
        # compat block (what layout depends on) in a minimal settings part
        src_settings = z.read("word/settings.xml").decode("utf-8")
        compat = re.search(r"<w:compat>.*?</w:compat>", src_settings, re.S)
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                    '<w:characterSpacingControl w:val="compressPunctuation"/>' + (compat.group(0) if compat else "") + "</w:settings>")
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as o:
            o.writestr("[Content_Types].xml", ct); o.writestr("_rels/.rels", T.RELS)
            o.writestr("word/_rels/document.xml.rels", drels); o.writestr("word/styles.xml", styles)
            o.writestr("word/settings.xml", settings); o.writestr("word/document.xml", doc)
            if theme:
                o.writestr("word/theme/theme1.xml", theme)
    print("ok", len(ARMS))


def _read(lines):
    ls = sorted([l for l in lines if l["t"].strip()], key=lambda l: l["y"])
    return {"lines": [(round(l["y"] - BOX_TOP, 2), round(l["x0"], 1), l["t"][:8]) for l in ls[:3]]}


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
        res[arm] = _read(L.word_lines(OUT / f"{arm}.pdf"))
        print(f"{arm:17s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        print(f"{arm:17s} oxi {o}   word {word.get(arm)}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()

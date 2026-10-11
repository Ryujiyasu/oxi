# -*- coding: utf-8 -*-
"""What does Word do with a tab whose next stop lies past the right edge?

en/forms__00619ac4: a 281pt cell «☐ Yes ☐ No» + 8 <w:tab/> + «If yes, insert UNGM
Vendor Number»: Word sets six tabs on line 1 (default 36pt stops), wraps, and starts
line 2 with the remaining tabs before the text; Oxi keeps everything on one line
(one line short -> the table below sits 12pt high).

Arms: <where>_k<n>: «No» + n tabs + «END»; where = body (468pt measure) / cell
(a 1-cell table, 281pt). Read the line (PDF baseline) of «No» and «END» and END's x.

    python tools/metrics/_pb_tab_overflow_gen.py gen
    python tools/metrics/_pb_tab_overflow_gen.py word
    python tools/metrics/_pb_tab_overflow_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_tab_overflow"
ARMS = {f"{w}_k{n}": (w, n) for w in ("body", "cell") for n in range(4, 16)}
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr>'
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '>'
            '<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def para(n):
    tabs = "".join(f"<w:r>{RPR}<w:tab/></w:r>" for _ in range(n))
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>{RPR}</w:pPr><w:r>{RPR}<w:t>No</w:t></w:r>{tabs}'
            f'<w:r>{RPR}<w:t>END</w:t></w:r></w:p>')


def body(where, n):
    if where == "body":
        return para(n)
    return ('<w:tbl><w:tblPr><w:tblW w:w="5625" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/>'
            '<w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
            '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="5625"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="5625" w:type="dxa"/></w:tcPr>{para(n)}</w:tc></w:tr></w:tbl>'
            '<w:p/>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    for arm, (where, n) in ARMS.items():
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" +
               body(where, n) +
               '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", T.STYLES)
            z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def word():
    import win32com.client, fitz
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
        pos = {}
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    for k in ("No", "END"):
                        if k in t and k not in pos:
                            c = s["chars"][t.index(k)]
                            pos[k] = (round(c["origin"][0], 2), round(c["origin"][1], 2))
        res[arm] = pos
        nl = round(pos["END"][1] - pos["No"][1], 2) if "END" in pos and "No" in pos else None
        print(f"{arm:9s} END x {pos.get('END', (None,))[0]}  dy(END-No) {nl}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import subprocess
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "l.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-layout=" + out],
                       capture_output=True, timeout=300)
        els = json.load(open(out, encoding="utf-8"))["pages"][0]["elements"]
        pos = {}
        for e in els:
            if e["type"] == "text" and e.get("text") in ("No", "END") and e["text"] not in pos:
                pos[e["text"]] = (round(e["x"], 2), round(e["y"], 2))
        nl = round(pos["END"][1] - pos["No"][1], 2) if len(pos) == 2 else None
        w = word[arm]
        wnl = round(w["END"][1] - w["No"][1], 2) if "END" in w and "No" in w else None
        print(f"{arm:9s} oxi END x {pos.get('END', (None,))[0]} dy {nl}   word x {w.get('END', (None,))[0]} dy {wnl}",
              flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()

# -*- coding: utf-8 -*-
"""Does a direct w:spacing before/after switch off a STYLE's HTML autospacing?

en/administrative__0069bb27: style «paragraph» = before 100 beforeAutospacing=1,
after 100 afterAutospacing=1; body paragraphs add a DIRECT <w:spacing w:before="0"
w:after="0"/>. Word's PDF keeps ~14pt between them; Oxi (S895's per-side guard:
a direct before/after stops the style's autospacing) puts them 13.4pt apart.

Arms (compat 15, Arial 12, single spacing, four paragraphs P1..P4):
  none     no direct spacing
  d00      direct before=0 after=0             (the real document)
  d120     direct before=120 after=120
  dauto0   direct beforeAutospacing=0 afterAutospacing=0
  d00a0    direct before=0 after=0 + both autospacing=0
Read Word's PDF baseline steps P1->P2->P3.

    python tools/metrics/_pb_autosp_direct_gen.py gen
    python tools/metrics/_pb_autosp_direct_gen.py word
    python tools/metrics/_pb_autosp_direct_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_autosp_direct"
DIRECT = {
    "none": "",
    "d00": '<w:spacing w:before="0" w:after="0"/>',
    "d120": '<w:spacing w:before="120" w:after="120"/>',
    "dauto0": '<w:spacing w:beforeAutospacing="0" w:afterAutospacing="0"/>',
    "d00a0": '<w:spacing w:before="0" w:beforeAutospacing="0" w:after="0" w:afterAutospacing="0"/>',
}
TOKENS = ("P1x", "P2x", "P3x", "P4x")
STYLES = T.STYLES.replace(
    "</w:styles>",
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="paragraph"><w:name w:val="paragraph"/>'
    '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100"'
    ' w:afterAutospacing="1" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial"'
    ' w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    for arm, sp in DIRECT.items():
        paras = "".join(f'<w:p><w:pPr><w:pStyle w:val="paragraph"/>{sp}</w:pPr><w:r><w:t>{t} body text</w:t></w:r></w:p>'
                        for t in TOKENS)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + paras +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    print("ok", list(DIRECT))


def _steps(ys):
    return [round(ys[TOKENS[i + 1]] - ys[TOKENS[i]], 2) for i in range(3)] if all(t in ys for t in TOKENS) else None


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in DIRECT:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in DIRECT:
        ys = {}
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"])
                for k in TOKENS:
                    if t.startswith(k) and k not in ys:
                        ys[k] = round(l["spans"][0]["origin"][1], 3)
        res[arm] = ys
        print(f"{arm:7s} first {ys.get(TOKENS[0])}  steps {_steps(ys)}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in DIRECT:
        ys = {}
        for o in L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"):
            for k in TOKENS:
                if o["t"].startswith(k) and k not in ys:
                    ys[k] = round(o["y"], 3)
        print(f"{arm:7s} oxi steps {_steps(ys)}  word {_steps(word[arm])}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()

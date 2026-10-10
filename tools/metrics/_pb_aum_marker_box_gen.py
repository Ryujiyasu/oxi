# -*- coding: utf-8 -*-
"""OXI_AUM_MARKER_BOX probe: how far does a numbering marker set in the
(missing) Arial Unicode MS raise its line?

One docx per arm. Each holds, in the body and then in a 1-column table cell,
the triple  «Before line» / «<numbered> Heading line» / «After line»  in
Calibri 10. The heading's marker is decimal "%1" with level rPr
rFonts ascii/hAnsi/eastAsia = Arial Unicode MS, hint=eastAsia, size MSZ.
Control arm NONUM: the same heading with no numbering.
Read from Word's PDF: baseline gap Before -> Heading (growth sits above the
baseline) and Heading -> After.

    python tools/metrics/_pb_aum_marker_box_gen.py gen
    python tools/metrics/_pb_aum_marker_box_gen.py word
    python tools/metrics/_pb_aum_marker_box_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "tests/fixtures/aum_marker_box"
ARMS = {"NONUM": None, "AUM8": 16, "AUM10": 20, "AUM12": 24, "AUM16": 32, "ARIAL10": "arial"}

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
DOCRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"""
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>"""
FONTS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts {W}><w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font><w:font w:name="Arial Unicode MS"><w:altName w:val="Yu Gothic"/><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="80"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="F7FFAFFF" w:usb1="E9DFFFFF" w:usb2="0000003F" w:usb3="00000000" w:csb0="003F01FF" w:csb1="00000000"/></w:font></w:fonts>"""
SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""


def numbering(arm):
    v = ARMS[arm]
    if v == "arial":
        rpr = '<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:hint="eastAsia"/><w:sz w:val="20"/>'
    else:
        rpr = f'<w:rFonts w:ascii="Arial Unicode MS" w:eastAsia="Arial Unicode MS" w:hAnsi="Arial Unicode MS" w:hint="eastAsia"/><w:sz w:val="{v or 20}"/>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering {W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="720"/></w:pPr><w:rPr>{rpr}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"""


def para(text, num=False, bold=False):
    ppr = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' if num else ""
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def document(arm):
    num = ARMS[arm] is not None
    trip = lambda tag: (para(f"Before {tag} line") + para(f"Heading {tag} line", num, True) + para(f"After {tag} line"))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
           f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{trip("cell")}</w:tc></w:tr></w:tbl>')
    body = trip("body") + para("Gap") + tbl + para("End")
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DOCRELS); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/numbering.xml", numbering(arm)); z.writestr("word/fontTable.xml", FONTS)
            z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", document(arm))
    print("ok", list(ARMS))


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    import fitz
    for arm in ARMS:
        ys = {}
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(c["c"] for s in l["spans"] for c in s["chars"])
                for tag in ("body", "cell"):
                    for k in ("Before", "Heading", "After"):
                        if f"{k} {tag}" in t:
                            ys[f"{k}_{tag}"] = l["spans"][-1]["origin"][1]
        res[arm] = gaps(ys)
        print(arm, res[arm], flush=True)
    (OUT / "word.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


def gaps(ys):
    return {tag: (round(ys[f"Heading_{tag}"] - ys[f"Before_{tag}"], 2), round(ys[f"After_{tag}"] - ys[f"Heading_{tag}"], 2))
            for tag in ("body", "cell") if all(f"{k}_{tag}" in ys for k in ("Before", "Heading", "After"))}


def oxi(exe):
    sys.path.insert(0, str(REPO / "tools/metrics"))
    import line_census as L
    res = {}
    for arm in ARMS:
        ys = {}
        for o in L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"):
            for tag in ("body", "cell"):
                for k in ("Before", "Heading", "After"):
                    if f"{k}{tag}line" in o["t"]:
                        ys[f"{k}_{tag}"] = o["y"]
        res[arm] = gaps(ys)
        print(arm, res[arm], flush=True)
    (OUT / "oxi.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()

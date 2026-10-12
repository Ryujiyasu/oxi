# -*- coding: utf-8 -*-
"""A numbered paragraph whose DIRECT w:ind carries firstLine="0" (or a positive
firstLine) while the numbering level says hanging=360/420: where does Word put
the label, and where does the suffix tab send the text?

golden 3a4f (就業規則 model, ＭＳ 明朝 10.5, docGrid lines 360, dts 840): p470
«① 一般勤務» (level left=1836 hanging=360, direct leftChars=202 left=424
firstLine=0, in a tblInd=250 cell) — Word label at L=21.2 from the cell text
edge, text at the 42pt default stop; Oxi label at L−18 (the level's hanging
kept) and text at L. The same for p478 / p940 (num tab 490) / p974 (num tab
426, level 2316/420) / p1058 (level 586/360) in the body.

Arms (one paragraph each, its own numbering so every label is ①; the label
and the text x are read from the PDF, margin-relative):
    A_fl0_lc202   level 1836/360, ind leftChars=202 left=424 firstLine=0   (p470)
    A_fl0_lc0     level 1836/360, ind leftChars=0 left=426 firstLine=0     (p478)
    A_nofl        level 1836/360, ind leftChars=0 left=426                 (control: hanging from the level)
    A_fl210       level 1836/360, ind leftChars=0 left=426 firstLine=210   (positive first line)
    A_fl0_l0      level 1836/360, ind left=0 firstLine=0
    A_fl0_l840    level 1836/360, ind left=840 firstLine=0                 (label ends past the 42 stop)
    A_hang210     level 1836/360, ind left=426 hanging=210                 (control: direct hanging)
    B_fl0         level 586/360,  ind leftChars=0 left=426 firstLine=0     (p1058)
    C_fl0_numtab  level 2316/420 + level tab 2316; para tabs clear 2316 + num 426; ind leftChars=0 left=426 firstLine=0 (p974)
    D_fl0_tab490  level 1680/420 + level tab 1680; para tabs clear 1680 + num 490; ind leftChars=203 left=426 firstLine=0 (p940)
    A_fl0_wide    level 1836/360 with lvlText «%1．» decimalFullWidth start 10 (label 2 cells wide), ind left=426 firstLine=0
CELL=1 wraps the paragraphs in 3a4f's 1-column table (tblInd 250, cell margin
108); LATIN=1 renders a Calibri-11 English family without a docGrid.

    python tools/metrics/_pb_numfl0_gen.py gen
    python tools/metrics/_pb_numfl0_gen.py word
    python tools/metrics/_pb_numfl0_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_numfl0"
CELL = bool(os.environ.get("CELL"))
LATIN = bool(os.environ.get("LATIN"))
NAME = "numfl0" + ("_cell" if CELL else "") + ("_latin" if LATIN else "")
# (level left, level hanging, level tab or None, lvlText/numFmt/start, paragraph tabs xml, paragraph ind xml)
CIRC = ("%1", "decimalEnclosedCircle", 1)
ARMS = {
    "A_fl0_lc202": (1836, 360, None, CIRC, "", '<w:ind w:leftChars="202" w:left="424" w:firstLine="0"/>'),
    "A_fl0_lc0": (1836, 360, None, CIRC, "", '<w:ind w:leftChars="0" w:left="426" w:firstLine="0"/>'),
    "A_nofl": (1836, 360, None, CIRC, "", '<w:ind w:leftChars="0" w:left="426"/>'),
    "A_fl210": (1836, 360, None, CIRC, "", '<w:ind w:leftChars="0" w:left="426" w:firstLine="210"/>'),
    "A_fl0_l0": (1836, 360, None, CIRC, "", '<w:ind w:left="0" w:firstLine="0"/>'),
    "A_fl0_l840": (1836, 360, None, CIRC, "", '<w:ind w:left="840" w:firstLine="0"/>'),
    "A_hang210": (1836, 360, None, CIRC, "", '<w:ind w:left="426" w:hanging="210"/>'),
    "B_fl0": (586, 360, None, CIRC, "", '<w:ind w:leftChars="0" w:left="426" w:firstLine="0"/>'),
    "C_fl0_numtab": (2316, 420, 2316, CIRC, '<w:tabs><w:tab w:val="clear" w:pos="2316"/><w:tab w:val="num" w:pos="426"/></w:tabs>',
                     '<w:ind w:leftChars="0" w:left="426" w:firstLine="0"/>'),
    "D_fl0_tab490": (1680, 420, 1680, CIRC, '<w:tabs><w:tab w:val="clear" w:pos="1680"/><w:tab w:val="num" w:pos="490"/></w:tabs>',
                     '<w:ind w:leftChars="203" w:left="426" w:firstLine="0"/>'),
    "A_fl0_wide": (1836, 360, None, ("%1．", "decimalFullWidth", 10), "", '<w:ind w:left="426" w:firstLine="0"/>'),
    # explicit stops past the label end: a left tab at 700 (35pt) / a num tab at 700 / a left tab at 1100 (55pt, past the 42 default)
    "A_fl0_tab700": (1836, 360, None, CIRC, '<w:tabs><w:tab w:val="left" w:pos="700"/></w:tabs>', '<w:ind w:leftChars="0" w:left="426" w:firstLine="0"/>'),
    "A_fl0_ntab700": (1836, 360, None, CIRC, '<w:tabs><w:tab w:val="num" w:pos="700"/></w:tabs>', '<w:ind w:leftChars="0" w:left="426" w:firstLine="0"/>'),
    "A_fl0_tab1100": (1836, 360, None, CIRC, '<w:tabs><w:tab w:val="left" w:pos="1100"/></w:tabs>', '<w:ind w:leftChars="0" w:left="426" w:firstLine="0"/>'),
    # a level whose num tab lies past the label end (level tab 700), no paragraph tabs
    "E_fl0_ltab700": (700, 360, 700, CIRC, "", '<w:ind w:leftChars="0" w:left="426" w:firstLine="0"/>'),
    # no first line at all on a level WITHOUT hanging (left=426 firstLine=0 in the level): where is the label?
    "F_lvlfl0": (426, 0, None, CIRC, "", ""),
}
FONT = "Calibri" if LATIN else "ＭＳ 明朝"
EA = "Calibri" if LATIN else "ＭＳ 明朝"
SZ = "22" if LATIN else "21"
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{EA}" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="{SZ}"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="a7"><w:name w:val="List Paragraph"/><w:basedOn w:val="a"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style>'
          '<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
          '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
          '</w:styles>')
MARGIN = 85.05
LEFT = MARGIN + (12.5 + 5.4 if CELL else 0.0)


def numbering():
    parts = []
    for i, (left, hang, ltab, (txt, fmt, start), _, _) in enumerate(ARMS.values()):
        tabs = f'<w:tabs><w:tab w:val="num" w:pos="{ltab}"/></w:tabs>' if ltab else ""
        ind = f'<w:ind w:left="{left}" w:hanging="{hang}"/>' if hang else f'<w:ind w:left="{left}" w:firstLine="0"/>'
        parts.append(f'<w:abstractNum w:abstractNumId="{i}"><w:multiLevelType w:val="hybridMultilevel"/>'
                     f'<w:lvl w:ilvl="0"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{txt}"/><w:lvlJc w:val="left"/>'
                     f'<w:pPr>{tabs}{ind}</w:pPr><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:lvl></w:abstractNum>')
    for i in range(len(ARMS)):
        parts.append(f'<w:num w:numId="{i + 1}"><w:abstractNumId w:val="{i}"/></w:num>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + T.NS + ">" + "".join(parts) + "</w:numbering>"


def para(i, ptabs, ind):
    text = f"Q{i} shift work hours" if LATIN else f"Ｑ{i}一般勤務の時刻"
    return (f'<w:p><w:pPr><w:pStyle w:val="a7"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{i + 1}"/></w:numPr>{ptabs}{ind}</w:pPr>'
            f'<w:r><w:t>{text}</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    paras = "".join(para(i, ptabs, ind) for i, (_, _, _, _, ptabs, ind) in enumerate(ARMS.values()))
    if CELL:
        paras = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="250" w:type="dxa"/><w:tblBorders>'
                 '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                 '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                 '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="8244"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="8458" w:type="dxa"/></w:tcPr>'
                 + paras + '</w:tc></w:tr></w:tbl><w:p/>')
    grid = "" if LATIN else '<w:docGrid w:type="lines" w:linePitch="360"/>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + paras +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701"'
           f' w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>{grid}</w:sectPr></w:body></w:document>')
    ct = T.CT.replace("</Types>", '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                            '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    with zipfile.ZipFile(OUT / f"{NAME}.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", numbering()); z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", NAME, len(ARMS))


def _read(lines):
    res = {}
    for i, arm in enumerate(ARMS):
        key = f"Q{i}" if LATIN else f"Ｑ{i}"
        tx = [l for l in lines if key in l["t"]]
        if not tx:
            res[arm] = None
            continue
        l = tx[0]
        gx = l.get("gx") or []
        k = l["t"].find(key)
        txt_x = gx[k] if gx and k < len(gx) else l["x0"]
        res[arm] = {"label_x": round(l["x0"] - LEFT, 2), "text_x": round(txt_x - LEFT, 2), "t": l["t"][:12]}
    return res


def _lines_with_glyphs(items):
    import line_census as L
    out = L.group_lines(list(items))
    for ln in out:
        ln["gx"] = [x for (x, _) in sorted((x0, ch) for (y, x0, x1, ch) in items if abs(y - ln["y"]) < 1.5 and ch.strip())]
    return out


def word_items(pdf):
    import fitz
    page = fitz.open(pdf)[0]; items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            if abs(l["dir"][1]) > 1e-3:
                continue
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((c["origin"][1], c["origin"][0], c["bbox"][2], c["c"]))
    return items


def oxi_items(exe, docx):
    import line_census as L, subprocess
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    d = json.load(open(out, encoding="utf-8")); items = []
    for g in d["pages"][0]["glyphs"]:
        fs = g.get("font_size", 0) or 0; base = g.get("baseline")
        if base is None:
            base = g["top"] + L.win_ascent(g.get("font_family", "")) * fs
        items.append((base, g["x"], g["x"] + fs, g["char"]))
    items.sort(key=lambda it: (round(it[0]), it[1]))
    return items


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        t = os.path.join(tempfile.mkdtemp(), f"{NAME}.docx"); shutil.copy(OUT / f"{NAME}.docx", t)
        d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
        d.SaveAs2(str((OUT / f"{NAME}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = _read(_lines_with_glyphs(word_items(OUT / f"{NAME}.pdf")))
    for arm in ARMS:
        print(f"{arm:13s} word {res[arm]}", flush=True)
    (OUT / f"word_{NAME}.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    word = json.loads((OUT / f"word_{NAME}.json").read_text(encoding="utf-8"))
    res = _read(_lines_with_glyphs(oxi_items(os.path.abspath(exe), OUT / f"{NAME}.docx")))
    for arm in ARMS:
        o, wd = res[arm], word.get(arm)
        flag = ""
        if o and wd and (abs(o["label_x"] - wd["label_x"]) > 1.0 or abs(o["text_x"] - wd["text_x"]) > 1.0):
            flag = "  <-- differs"
        print(f"{arm:13s} oxi {o}   word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()

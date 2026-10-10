# -*- coding: utf-8 -*-
"""Where does Word put a CJK line's baseline inside its 83/64 natural box?

ja/forms__02281020: BIZ UD明朝 Medium 12pt cell lines (snapToGrid=0, docGrid
lines/360) sit 0.6pt lower in Word than in Oxi relative to the row's top rule,
while the line pitch (15.6 = 12 x 83/64) matches.  S1322 left the BIZ UD /
Meiryo ascent side unresolved.

One docx per (face, size): body docGrid lines/360; a 1x1 table, top+bottom
borders 0.5pt, cell margins 0, holding two paragraphs «漢字あ» / «二行目»
(snapToGrid 0), then a body paragraph «後».  Read from Word's PDF: the top
rule's y (drawing), each line's baseline -> top->baseline1, baseline1->2.

    python tools/metrics/_pb_cjk_cell_top_gen.py gen
    python tools/metrics/_pb_cjk_cell_top_gen.py word
    python tools/metrics/_pb_cjk_cell_top_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_aum_marker_box_gen as G  # noqa: E402  (package parts)

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "tests/fixtures/cjk_cell_top"
FACES = {"bizm": "BIZ UD明朝 Medium", "bizg": "BIZ UDゴシック", "msm": "ＭＳ 明朝", "yum": "游明朝", "mei": "メイリオ"}
SIZES = {"105": 21, "12": 24, "16": 32}
ARMS = {f"{f}_{s}": (face, hp) for f, face in FACES.items() for s, hp in SIZES.items()}
# regime arms (12pt): "ng" = no docGrid at all, "nt" = docGrid linePitch without w:type
GRID = {"": '<w:docGrid w:type="lines" w:linePitch="360"/>', "ng": "", "nt": '<w:docGrid w:linePitch="360"/>'}
for f in ("msm", "bizm", "mei"):
    for g in ("ng", "nt"):
        ARMS[f"{f}_12_{g}"] = (FACES[f], 24)


def document(face, hp, grid=GRID[""]):
    rpr = f'<w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}" w:hint="eastAsia"/><w:sz w:val="{hp}"/></w:rPr>'
    p = lambda t: f'<w:p><w:pPr><w:snapToGrid w:val="0"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>'
    bd = '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>' + bd + '</w:tblBorders>'
           '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
           '</w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>'
           + p("漢字あ") + p("二行目") + "</w:tc></w:tr></w:tbl>")
    body = '<w:p><w:r><w:t>Top</w:t></w:r></w:p>' + tbl + p("後")
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {G.W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            + grid + '</w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    styles = G.STYLES.replace('<w:lang w:val="en-US" w:eastAsia="en-US"', '<w:lang w:val="en-US" w:eastAsia="ja-JP"')
    for arm, (face, hp) in ARMS.items():
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", G.CT); z.writestr("_rels/.rels", G.RELS)
            z.writestr("word/_rels/document.xml.rels", G.DOCRELS); z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", G.numbering("NONUM")); z.writestr("word/fontTable.xml", G.FONTS)
            z.writestr("word/settings.xml", G.SETTINGS); z.writestr("word/document.xml", document(face, hp, GRID[arm.split("_")[2]] if arm.count("_") == 2 else GRID[""]))
    print("ok", len(ARMS))


def measure(top, b1, b2):
    return {"top_to_b1": round(b1 - top, 2), "pitch": round(b2 - b1, 2)}


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
    for arm, (face, hp) in ARMS.items():
        pg = fitz.open(OUT / f"{arm}.pdf")[0]
        rules = []
        for dr in pg.get_drawings():
            for it in dr["items"]:
                if it[0] == "l" and abs(it[1].y - it[2].y) < 0.3:
                    rules.append(it[1].y)
                elif it[0] == "re" and it[1].height < 1.5:
                    rules.append(it[1].y0 + it[1].height / 2)
        ys = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"] in "漢二" and c["c"] not in ys:
                            ys[c["c"]] = c["origin"][1]
        res[arm] = dict(measure(min(rules), ys["漢"], ys["二"]), face=face, pt=hp / 2,
                        pdf_font=pg.get_text("rawdict")["blocks"][1]["lines"][0]["spans"][0]["font"] if len(pg.get_text("rawdict")["blocks"]) > 1 else "")
        print(arm, res[arm], flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")


def oxi(exe):
    sys.path.insert(0, str(REPO / "tools/metrics"))
    import line_census as L
    wj = json.loads((OUT / "word.json").read_text(encoding="utf-8")) if (OUT / "word.json").exists() else {}
    for arm in ARMS:
        t = tempfile.mkdtemp(); lay = os.path.join(t, "l.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-layout=" + lay], capture_output=True)
        top = min(e["y"] for e in json.load(open(lay, encoding="utf-8"))["pages"][0]["elements"] if e["type"] == "border" and e["w"] > 20)
        ys = {}
        for o in L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"):
            for ch in "漢二":
                if o["t"].startswith(ch) and ch not in ys:
                    ys[ch] = o["y"]
        m = measure(top, ys["漢"], ys["二"])
        w = wj.get(arm, {})
        print(arm, m, "word", {k: w.get(k) for k in ("top_to_b1", "pitch")},
              "d_top %+.2f" % (m["top_to_b1"] - w["top_to_b1"]) if w else "", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()

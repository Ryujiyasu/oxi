# -*- coding: utf-8 -*-
"""An EMPTY body paragraph whose mark carries an explicit sz: whose line?

policies__060b605eaef40085 (jablindC50): two empty centred paragraphs with
`<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="22"/></w:rPr>`
on a lines grid (360). Word COM pitch 13.5 / 13.5 (a 10.5pt ＭＳ 明朝 line,
13.6); Oxi prices the 11pt mark: 14.27. Day-33 measured the OPPOSITE in a
table cell (explicit mark sz applies).

Arms (docDefaults Century/ＭＳ 明朝 10.5, Normal no size, lines grid 360,
anchor A1 / two empties / anchor A2, Info6 pitch A1->A2 = 2 empties + 1 line):
  A_sz22_grid        mark sz=22, docGrid lines 360
  B_nosz_grid        mark no sz
  C_sz22_nogrid      mark sz=22, no docGrid
  D_sz22_snap0       mark sz=22, snapToGrid=0 on the empties
  E_sz28_grid        mark sz=28 (14pt: above the 18pt pitch? no, 18.15 > 18)
  F_sz22_cell        the A shape inside a table cell (the Day-33 case)
"""
import importlib.util, os, sys, json, subprocess, tempfile
from pathlib import Path
spec = importlib.util.spec_from_file_location("lp", "tools/metrics/_pb_line_pitch.py")
lp = importlib.util.module_from_spec(spec); spec.loader.exec_module(lp)
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/emptymark_sz'); OUT.mkdir(parents=True, exist_ok=True)
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/><w:szCs w:val="24"/></w:rPr></w:style>'
          '</w:styles>')
def sect(grid):
    g = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''
    return f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1021" w:right="1418" w:bottom="454" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>{g}</w:sectPr>'
def empty(sz, snap0):
    s = f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>' if sz else ''
    sn = '<w:snapToGrid w:val="0"/>' if snap0 else ''
    return f'<w:p><w:pPr>{sn}<w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>{s}</w:rPr></w:pPr></w:p>'
def anchor(t):
    return f'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="24"/></w:rPr><w:t>{t}</w:t></w:r></w:p>'
def body(sz, snap0, cell=False):
    inner = anchor('A1') + empty(sz, snap0) + empty(sz, snap0) + anchor('A2')
    if cell:
        inner = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>' + inner + '</w:tc></w:tr></w:tbl>' + anchor('TAIL'))
    return inner
def doc(sz, snap0, grid, cell=False):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            + body(sz, snap0, cell) + sect(grid) + '</w:body></w:document>')
arms = {"A_sz22_grid": (22, False, True, False), "B_nosz_grid": (None, False, True, False), "C_sz22_nogrid": (22, False, False, False),
        "D_sz22_snap0": (22, True, True, False), "E_sz28_grid": (28, False, True, False), "F_sz22_cell": (22, False, True, True)}
import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, (sz, snap0, grid, cell) in arms.items():
    at = OUT / f"{name}.docx"; lp.write(at, doc(sz, snap0, grid, cell), STYLES)
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        ys = []
        for i in range(1, d.Paragraphs.Count + 1):
            r = d.Paragraphs(i).Range; c = d.Range(r.Start, r.Start); ys.append((round(c.Information(6), 2), r.Text.strip()[:4]))
    finally: d.Close(False)
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8'))
    oxi = {el['text'][:4]: round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and el.get('text', '').strip()}
    w = {t: y for y, t in ys}
    print(f"=== {name}: WORD {ys[:4]} gap A1->A2 {round(w.get('A2', 0) - w.get('A1', 0), 2)}   OXI gap {round(oxi.get('A2', 0) - oxi.get('A1', 0), 2)}")
finally: app.Quit()

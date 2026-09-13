# -*- coding: utf-8 -*-
"""Does a NON-BINDING atLeast trHeight row split when only its first line fits?

technical__00afb3e6b2bb1a5a (blindD50): a 2-column spec table, every row
`<w:trHeight w:val="300" w:hRule="atLeast"/>` with 2pt cell margins, 7pt
lines (9.84). The last row on page 1 holds 2 paragraphs (56pt) and starts with
14.8pt left on the page: Word keeps its first line on page 1 and continues on
page 2. Oxi's `minimum_requires_page` demands trHeight + borders + margins
(~20pt) before it will split, so the whole row moves.

Probe: K filler Calibri-11 lines, then a 2-col table whose single row has one
short paragraph in cell 1 and FOUR 7pt Arial paragraphs in cell 2 (trHeight
300 atLeast, tcMar 40). K is swept so the remaining space R at the row top
walks from ~35 down to ~5. Readout per K: R (page_bottom - row top, row top =
first cell paragraph y - 2pt), the number of cell-2 paragraphs Word leaves on
page 1, and the same from Oxi's dump.
"""
import importlib.util, os, sys, json, subprocess, tempfile
from pathlib import Path
spec = importlib.util.spec_from_file_location("lp", "tools/metrics/_pb_line_pitch.py")
lp = importlib.util.module_from_spec(spec); spec.loader.exec_module(lp)
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/rowsplit_trh'); OUT.mkdir(parents=True, exist_ok=True)
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="864" w:right="1440" w:bottom="864" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
def cp(t):
    rpr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/></w:rPr>'
    return f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r></w:p>'
MAR = '<w:tcMar><w:top w:w="40" w:type="dxa"/><w:left w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tcMar>'
def table():
    c1 = f'<w:tc><w:tcPr><w:tcW w:w="1830" w:type="dxa"/>{MAR}</w:tcPr>{cp("Label")}</w:tc>'
    c2 = (f'<w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/>{MAR}</w:tcPr>'
          + cp('Line one of the cell') + cp('Line two of the cell') + cp('Line three of the cell') + cp('Line four of the cell') + '</w:tc>')
    return ('<w:tbl><w:tblPr><w:tblW w:w="8830" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="1830"/><w:gridCol w:w="7000"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="300" w:hRule="atLeast"/></w:trPr>' + c1 + c2 + '</w:tr>'
            '<w:tr><w:trPr><w:trHeight w:val="300" w:hRule="atLeast"/></w:trPr>' + c1 + c2 + '</w:tr></w:tbl>')
def doc(k):
    body = ''.join(f'<w:p><w:r><w:t>Filler line {i}</w:t></w:r></w:p>' for i in range(1, k + 1)) + table() + '<w:p><w:r><w:t>TAIL</w:t></w:r></w:p>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + SECT + '</w:body></w:document>')
PAGE_BOTTOM = 792 - 43.2
import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for k in range(48, 54):
    at = OUT / f"k{k}.docx"
    lp.write(at, doc(k), STYLES)
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        info = []
        for i in range(1, d.Paragraphs.Count + 1):
            r = d.Paragraphs(i).Range; c = d.Range(r.Start, r.Start)
            info.append((c.Information(3), round(c.Information(6), 2), r.Text.strip()[:9]))
    finally: d.Close(False)
    lab = [x for x in info if x[2] == 'Label']
    lines = [x for x in info if x[2].startswith('Line ')][:4]
    w_top = lab[0][1] - 2.0 if lab else None
    w_p1 = sum(1 for x in lines if x[0] == 1)
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8'))
    o_lines = [(pi + 1, round(el['y'], 2)) for pi, pg in enumerate(dd['pages']) for el in pg['elements'] if el.get('type') == 'text' and el.get('text', '').startswith('Line')]
    o_lab = [(pi + 1, round(el['y'], 2)) for pi, pg in enumerate(dd['pages']) for el in pg['elements'] if el.get('type') == 'text' and el.get('text', '') == 'Label']
    o_p1 = sum(1 for p_, _ in o_lines[:4] if p_ == 1)
    print(f"k={k}: WORD row top {w_top} R={round(PAGE_BOTTOM - w_top, 2) if w_top else None} lines on p1 = {w_p1} (label p{lab[0][0] if lab else '?'})   OXI label {o_lab[:1]} lines on p1 = {o_p1}")
finally: app.Quit()

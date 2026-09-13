# -*- coding: utf-8 -*-
"""An EMPTY paragraph whose mark names a Latin-only eastAsia face: whose line?

reports__003862302b660a86 (blindD50): cell paragraphs carry
`<w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial"/>`
with sz=20. Word renders the empty cell paragraphs at Arial 10's line (row
24.6 = 12 + 12.5); Oxi's actual cell-empty consumer sees an EXPLICIT eastAsia,
skips the S989 ascii preference, and prices the mark through the CJK
substitute (12.97 = a 10pt 83/64 box) -- +1.47 per such row.

Arms (Calibri 11 body, mark sz=20, an empty paragraph between two anchors,
once in the body and once inside a table cell):
  A_ea_tnr       ascii=Arial hAnsi=Arial eastAsia="Times New Roman"
  B_ea_msmincho  ascii=Arial hAnsi=Arial eastAsia="ＭＳ 明朝"
  C_no_ea        ascii=Arial hAnsi=Arial
  D_ea_yu        ascii=Arial hAnsi=Arial eastAsia="Yu Mincho"
Readout: Word COM Information(6) gap anchor1 -> anchor2 (= empty line), body
and cell; PDF lines; Oxi dump.
"""
import importlib.util, os, sys, json, subprocess, tempfile
from pathlib import Path
spec = importlib.util.spec_from_file_location("lp", "tools/metrics/_pb_line_pitch.py")
lp = importlib.util.module_from_spec(spec); spec.loader.exec_module(lp)
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/markea_latin'); OUT.mkdir(parents=True, exist_ok=True)
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
def rf(ea):
    return '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"' + (f' w:eastAsia="{ea}"' if ea else '') + '/>'
def p(t, ea):
    rpr = f'<w:rPr>{rf(ea)}<w:sz w:val="20"/></w:rPr>'
    run = f'<w:r>{rpr}<w:t>{t}</w:t></w:r>' if t else ''
    return f'<w:p><w:pPr>{rpr}</w:pPr>{run}</w:p>'
def cell(ea):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
            + p('C1', ea) + p('', ea) + p('C2', ea) + '</w:tc></w:tr></w:tbl>')
def doc(ea):
    body = p('A1', ea) + p('', ea) + p('A2', ea) + p('gap', None) + cell(ea) + p('TAIL', None)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + SECT + '</w:body></w:document>')
arms = {"A_ea_tnr": "Times New Roman", "B_ea_msmincho": "ＭＳ 明朝", "C_no_ea": None, "D_ea_yu": "Yu Mincho"}
import pymupdf, win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, ea in arms.items():
    at = OUT / f"{name}.docx"
    lp.write(at, doc(ea), STYLES)
    pdf = str(at)[:-5] + '.pdf'
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        d.ExportAsFixedFormat(OutputFileName=str(Path(pdf).resolve()), ExportFormat=17)
        ys = []
        for i in range(1, d.Paragraphs.Count + 1):
            r = d.Paragraphs(i).Range; c = d.Range(r.Start, r.Start)
            ys.append((round(c.Information(6), 2), r.Text.strip()[:6]))
    finally: d.Close(False)
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8'))
    oxi = sorted({(round(el['y'], 2), el['text'][:4]) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and el.get('text', '').strip()})
    wy = {t: y for y, t in ys}
    oy = {t: y for y, t in oxi}
    print(f"=== {name}: WORD body gap {round(wy.get('A2',0)-wy.get('A1',0),2)}  cell gap {round(wy.get('C2',0)-wy.get('C1',0),2)}   OXI body gap {round(oy.get('A2',0)-oy.get('A1',0),2)}  cell gap {round(oy.get('C2',0)-oy.get('C1',0),2)}")
finally: app.Quit()

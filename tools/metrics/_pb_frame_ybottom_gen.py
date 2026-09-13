# -*- coding: utf-8 -*-
"""Where does Word put a framePr paragraph run with yAlign=bottom?

correspondence__0059143bed49147b (blindD50): nineteen consecutive 6.5pt
paragraphs carry `<w:framePr w:w="2659" w:wrap="around" w:hAnchor="page"
w:x="8971" w:yAlign="bottom" w:anchorLock="1"/>` -- an address column in the
right page margin. Word COM puts them at y 612..792 (the group's bottom on the
bottom margin, x = 448.5) and the body starts at 159.75; Oxi lays them in the
body flow at the top and the body starts 189pt low.

Arms (Calibri 11 body of six lines, three 9pt frame paragraphs, A4 margins
top 72 / bottom 72 / left 72 / right 170 so the text column ends at 425):
  A_margin_right   hAnchor=page x=8971(448.5pt) yAlign=bottom       outside the column
  B_overlap        hAnchor=margin x=2000(100pt) yAlign=bottom       inside the column (wrap around)
  C_vpage          as A + vAnchor=page                              bottom of the PAGE
  D_single         as A, one frame paragraph only
  E_top            as A but yAlign=top
For each: COM Information(6)/(5) of every paragraph, PDF line boxes, Oxi dump.
"""
import importlib.util, os, sys, json, subprocess, tempfile
from pathlib import Path
spec = importlib.util.spec_from_file_location("lp", "tools/metrics/_pb_line_pitch.py")
lp = importlib.util.module_from_spec(spec); spec.loader.exec_module(lp)
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/frame_ybottom'); OUT.mkdir(parents=True, exist_ok=True)

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="3402" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')

def p(t):
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'

def fp(t, attrs):
    return (f'<w:p><w:pPr><w:framePr {attrs}/><w:rPr><w:sz w:val="18"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{t}</w:t></w:r></w:p>')

def doc(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            + body + SECT + '</w:body></w:document>')

BODY = [f'Body line {i} of the letter text that fills the column width nicely.' for i in range(1, 7)]
A = 'w:w="2659" w:wrap="around" w:hAnchor="page" w:x="8971" w:yAlign="bottom" w:anchorLock="1"'
B = 'w:w="2659" w:wrap="around" w:hAnchor="margin" w:x="2000" w:yAlign="bottom" w:anchorLock="1"'
C = 'w:w="2659" w:wrap="around" w:hAnchor="page" w:vAnchor="page" w:x="8971" w:yAlign="bottom" w:anchorLock="1"'
E = 'w:w="2659" w:wrap="around" w:hAnchor="page" w:x="8971" w:yAlign="top" w:anchorLock="1"'
FR = ['Frame one', 'Frame two', 'Frame three']
def arm(attrs, nfr=3):
    return p(BODY[0]) + ''.join(fp(t, attrs) for t in FR[:nfr]) + ''.join(p(t) for t in BODY[1:])
arms = {"A_margin_right": arm(A), "B_overlap": arm(B), "C_vpage": arm(C), "D_single": arm(A, 1), "E_top": arm(E)}

import pymupdf, win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, body in arms.items():
    at = OUT / f"{name}.docx"
    lp.write(at, doc(body), STYLES)
    pdf = str(at)[:-5] + '.pdf'
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        d.ExportAsFixedFormat(OutputFileName=str(Path(pdf).resolve()), ExportFormat=17)
        info = []
        for i in range(1, d.Paragraphs.Count + 1):
            r = d.Paragraphs(i).Range
            c = d.Range(r.Start, r.Start)
            info.append((c.Information(3), round(c.Information(6), 2), round(c.Information(5), 1), r.Text.strip()[:10]))
    finally: d.Close(False)
    pg = pymupdf.open(pdf)
    lines = sorted((round(l['bbox'][1], 2), round(l['bbox'][0], 1), round(l['bbox'][2], 1), ''.join(s['text'] for s in l['spans'])[:10])
                   for page in pg for b in page.get_text('dict')['blocks'] for l in b.get('lines', []))
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8'))
    rows = {}
    for el in dd['pages'][0]['elements']:
        if el.get('type') == 'text' and el.get('text', '').strip():
            rows.setdefault((round(el['y'], 2), round(el['x'], 1)), []).append(el['text'])
    print(f"=== {name}")
    print("  WORD COM  :", [f"{y}/{x}:{t}" for _, y, x, t in info])
    print("  WORD PDF  :", [f"{y}/{x0}-{x1}:{t}" for y, x0, x1, t in lines])
    print("  OXI       :", [f"{y}/{x}:{''.join(v)[:10]}" for (y, x), v in sorted(rows.items())])
finally: app.Quit()

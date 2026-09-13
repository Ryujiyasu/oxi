# -*- coding: utf-8 -*-
"""Does an EMPTY paragraph that carries a CONTINUOUS section break take a line?

technical__00ac06fef95cc36c (blindD50): a 2-column section of six one-line
paragraphs closed by an empty sect-end paragraph. Word balances 3/3 and the
next section starts three lines down; Oxi counts the empty sect-end as a
seventh line, balances 4/3, and everything below sits one line low (the last
column overflows to page 2). Word COM Information(6) of that empty paragraph
equals the y of the line before it -- zero height. This probe isolates it.

Arms (all continuous section breaks, Calibri 11, no docGrid):
  A_1col_empty      HEAD | P1 P2 P3 | ''(sect) | TAIL          gap P3->TAIL ?
  B_2col_empty      HEAD(sect) | P1..P6 | ''(sect cols=2) | TAIL   3 or 4 lines?
  C_2col_text       same as B, sect-end paragraph has text        4/3 -> 4 lines
  D_1col_2empty     HEAD | P1 P2 P3 | '' | ''(sect) | TAIL       1 or 2 lines?
  E_2col_5_empty    HEAD(sect) | P1..P5 | ''(sect cols=2) | TAIL  3 lines either way (control)
  F_2col_empty_np   as B but the sect-end break is nextPage       control
"""
import importlib.util, os, sys, json, subprocess, tempfile
from pathlib import Path
spec = importlib.util.spec_from_file_location("lp", "tools/metrics/_pb_line_pitch.py")
lp = importlib.util.module_from_spec(spec); spec.loader.exec_module(lp)
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/contsect_empty'); OUT.mkdir(parents=True, exist_ok=True)

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '</w:styles>')
PG = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'

def sect(cols, kind="continuous"):
    c = f'<w:cols w:num="{cols}" w:space="720"/>' if cols > 1 else '<w:cols w:space="720"/>'
    return f'<w:sectPr><w:type w:val="{kind}"/>{PG}{c}</w:sectPr>'

def p(t="", sectpr=""):
    ppr = f'<w:pPr>{sectpr}</w:pPr>' if sectpr else ''
    run = f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>' if t else ''
    return f'<w:p>{ppr}{run}</w:p>'

def doc(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            + body + sect(1) + '</w:body></w:document>')

P = [f'P{i} line' for i in range(1, 7)]
arms = {
    "A_1col_empty":   p('HEAD') + ''.join(p(t) for t in P[:3]) + p('', sect(1)) + p('TAIL'),
    "B_2col_empty":   p('HEAD', sect(1)) + ''.join(p(t) for t in P[:6]) + p('', sect(2)) + p('TAIL'),
    "C_2col_text":    p('HEAD', sect(1)) + ''.join(p(t) for t in P[:6]) + p('SECTEND', sect(2)) + p('TAIL'),
    "D_1col_2empty":  p('HEAD') + ''.join(p(t) for t in P[:3]) + p('') + p('', sect(1)) + p('TAIL'),
    "E_2col_5_empty": p('HEAD', sect(1)) + ''.join(p(t) for t in P[:5]) + p('', sect(2)) + p('TAIL'),
    "F_2col_empty_np": p('HEAD', sect(1)) + ''.join(p(t) for t in P[:6]) + p('', sect(2, "nextPage")) + p('TAIL'),
}

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
            info.append((i, c.Information(3), round(c.Information(6), 2), round(c.Information(5), 1), r.Text.strip()[:8]))
    finally: d.Close(False)
    pg = pymupdf.open(pdf)
    lines = sorted((pi + 1, round(l['bbox'][1], 2), round(l['bbox'][0], 1), ''.join(s['text'] for s in l['spans']))
                   for pi, page in enumerate(pg) for b in page.get_text('dict')['blocks'] for l in b.get('lines', []))
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8'))
    oxi = [(pi + 1, round(el['y'], 2), round(el['x'], 1), el['text'][:8]) for pi, pgd in enumerate(dd['pages'])
           for el in pgd['elements'] if el.get('type') == 'text' and el.get('text', '').strip()]
    print(f"=== {name}")
    print("  WORD COM  :", [f"p{pg_}:{y}/{x}:{t}" for _, pg_, y, x, t in info])
    print("  WORD PDF  :", [f"p{pg_}:{y}/{x}:{t[:8]}" for pg_, y, x, t in lines])
    print("  OXI       :", [f"p{pg_}:{y}/{x}:{t}" for pg_, y, x, t in sorted(oxi)])
finally: app.Quit()

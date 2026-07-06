# -*- coding: utf-8 -*-
"""Controlled sweep: the bd90b00 multi-para cell row (+0.36 residual under
ROWBOX2). Row = one cell with [exact-200 8pt note / auto 10.5 heading /
auto 10.5 line], docDefaults spacing line=254 exact, docGrid lines 330.
Variants isolate each factor. Word PDF border pitches + text line ys.

Run: python tools/metrics/_bd90_cell_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), "bd90_sweep")
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, "bd90_sweep.docx")
PDF = os.path.join(OUTDIR, "bd90_sweep.pdf")

esc = pg.esc
MINCHO = pg.MINCHO

def rpr(sz):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')

def para(txt, sz, spacing):
    r = rpr(sz)
    return (f'<w:p><w:pPr>{spacing}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

EXACT200 = '<w:spacing w:line="200" w:lineRule="exact"/>'
AUTO240 = '<w:spacing w:line="240" w:lineRule="auto"/>'

def table(cell_paras):
    return ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="6000"/><w:gridCol w:w="3000"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{cell_paras}</w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para("あ", "21", AUTO240)}</w:tc></w:tr>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{para("参照行いろは", "21", AUTO240)}</w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para("い", "21", AUTO240)}</w:tc></w:tr>'
            '</w:tbl>')

NOTE = '※上記以外の事項は実際の年月日とする'
HEAD = '８規則第27条関係'
INSTR = '以下の各事項に該当する場合に付す'

configs = [
    ("V1_replica",   table(para(NOTE, "16", EXACT200) + para(HEAD, "21", AUTO240) + para(INSTR, "21", AUTO240))),
    ("V2_all_auto",  table(para(NOTE, "16", AUTO240) + para(HEAD, "21", AUTO240) + para(INSTR, "21", AUTO240))),
    ("V3_two_auto",  table(para(HEAD, "21", AUTO240) + para(INSTR, "21", AUTO240))),
    ("V4_one_auto",  table(para(HEAD, "21", AUTO240))),
    ("V5_exact_one", table(para(NOTE, "16", EXACT200) + para(HEAD, "21", AUTO240))),
    ("V6_nodefsp",   None),  # placeholder — second docx without docDefaults spacing
]

def marker(i):
    r = rpr("21")
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:br w:type="page"/></w:r>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>M{i}</w:t></w:r></w:p>')

body = []
for i, (tag, tbl) in enumerate(c for c in configs if c[1] is not None):
    if i > 0:
        body.append(marker(i))
    else:
        body.append(para("M0", "21", AUTO240))
    body.append(tbl)

# doc() uses _probe_gen's skeleton; override docDefaults + docGrid via custom styles/sectPr
SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" w:header="851" w:footer="567" w:gutter="0"/>'
          '<w:docGrid w:type="lines" w:linePitch="330"/></w:sectPr>')

DOCDEF_SPACING = '<w:spacing w:line="254" w:lineRule="exact"/>'

SETTINGS_ALIT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:characterSpacingControl w:val="compressPunctuation"/>'
    '<w:adjustLineHeightInTable/><w:useFELayout/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')

def build(docx, defsp, alit=False):
    ppr_def = f'<w:pPr>{defsp}<w:jc w:val="both"/></w:pPr>' if defsp else '<w:pPr><w:jc w:val="both"/></w:pPr>'
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
              '<w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
              '</w:rPr></w:rPrDefault>'
              f'<w:pPrDefault>{ppr_def}</w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>'
              '</w:styles>')
    doc_xml = pg.doc(''.join(body) + SECTPR)
    extra = {'word/styles.xml': styles}
    if alit:
        extra['word/settings.xml'] = SETTINGS_ALIT
    pg.write_docx(docx, doc_xml, extra_parts=extra)

DOCX2 = os.path.join(OUTDIR, "bd90_sweep_nodef.docx")
PDF2 = os.path.join(OUTDIR, "bd90_sweep_nodef.pdf")
DOCX3 = os.path.join(OUTDIR, "bd90_sweep_alit.docx")
PDF3 = os.path.join(OUTDIR, "bd90_sweep_alit.pdf")
build(DOCX, DOCDEF_SPACING)
build(DOCX2, None)
build(DOCX3, DOCDEF_SPACING, alit=True)
print('wrote', DOCX, DOCX2, DOCX3)

import win32com.client
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
for dx, pf in ((DOCX, PDF), (DOCX2, PDF2), (DOCX3, PDF3)):
    doc = word.Documents.Open(dx, ReadOnly=True, AddToRecentFiles=False)
    doc.ExportAsFixedFormat(pf, 17)
    doc.Close(False)
word.Quit()
print('exported PDFs')

import fitz
tags = [t for t, tbl in configs if tbl is not None]
for pf, label in ((PDF, 'defsp=exact254'), (PDF2, 'defsp=none'), (PDF3, 'defsp=exact254+ALIT')):
    d = fitz.open(pf)
    print(f'== {label} ({pf}) pages={len(d)}')
    for i, tag in enumerate(tags):
        if i >= len(d):
            break
        page = d[i]
        ys = set()
        for dr in page.get_drawings():
            for it in dr["items"]:
                if it[0] == "l":
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 0.2 and abs(p1.x - p2.x) > 60:
                        ys.add(round(p1.y, 2))
                elif it[0] == "re":
                    rr = it[1]
                    if rr.height < 2.0 and rr.width > 60:
                        ys.add(round(rr.y0, 2))
        merged = []
        for y in sorted(ys):
            if merged and abs(merged[-1] - y) < 1.2:
                continue
            merged.append(y)
        pitches = [round(merged[j+1] - merged[j], 2) for j in range(len(merged) - 1)]
        lines = []
        for b in page.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                txt = ''.join(s['text'] for s in l.get('spans', []))
                if txt.strip() and not txt.strip().startswith('M'):
                    lines.append((round(l['bbox'][1], 2), txt.strip()[:10]))
        lines.sort()
        print(f'  {tag:12s} pitches={pitches}')
        print(f'               lines={lines}')

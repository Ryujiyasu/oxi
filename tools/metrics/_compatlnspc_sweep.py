# -*- coding: utf-8 -*-
"""Controlled sweep: compatLnSpc="1" textbox LINE STACKING law for
MIXED-size lines (the 1ec1 wall: steady pitches match ±0.04 but the
size-transition gaps err ±1.5 with opposite signs per box).

One wps textbox per config (compatLnSpc=1, fixed tIns), lines:
  [szA line] [szB line] [szB line]
Sweep (szA, szB) pairs + fonts. Word truth: per-line baseline from the
PDF → first_offset(A) from box top, gap(A→B), gap(B→B).

Run: python tools/metrics/_compatlnspc_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'gridquant')
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, 'clnsp.docx')
PDF = os.path.join(OUTDIR, 'clnsp.pdf')
MINCHO = pg.MINCHO
GOTHIC = "ＭＳ ゴシック"
esc = pg.esc
EMU = 12700

WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def rpr(sz, font):
    return (f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{sz}"/>')


SG0 = os.environ.get('CLNSP_SG0', '1') == '1'

def inner_para(txt, sz, font):
    r = rpr(sz, font)
    sg = '<w:snapToGrid w:val="0"/>' if SG0 else ''
    return (f'<w:p><w:pPr>{sg}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')


def tb(pid, paras, cx_pt=430.0, cy_pt=110.0):
    cx, cy = int(cx_pt * EMU), int(cy_pt * EMU)
    choice = (
        f'<w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>127000</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        f'<wp:docPr id="{pid}" name="TB{pid}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic {A_NS}>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        '</wps:spPr>'
        f'<wps:txbx><w:txbxContent>{paras}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="18000" tIns="18000" '
        'rIns="18000" bIns="18000" anchor="t" anchorCtr="0" compatLnSpc="1">'
        '<a:noAutofit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>')
    return ('<mc:AlternateContent '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<mc:Choice Requires="wps">{choice}</mc:Choice>'
            f'<mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>')


def para(txt, sz="21"):
    r = rpr(sz, MINCHO)
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')


def pagebreak():
    return ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:br w:type="page"/></w:r></w:p>')


# (szA halfpt, szB halfpt, font)
CONFIGS = [
    ("28", "21", GOTHIC),   # 14 -> 10.5 (the 1ec1 case)
    ("28", "21", MINCHO),
    ("36", "24", GOTHIC),   # 18 -> 12
    ("24", "21", GOTHIC),   # 12 -> 10.5
    ("21", "21", GOTHIC),   # control: uniform 10.5
    ("28", "28", GOTHIC),   # control: uniform 14
    ("21", "28", GOTHIC),   # ascending 10.5 -> 14
    ("42", "21", GOTHIC),   # 21 -> 10.5 big jump
]


def build():
    body = []
    pid = 300
    for i, (a, bsz, font) in enumerate(CONFIGS):
        if i:
            body.append(pagebreak())
        marks = (inner_para(f'Ａ{i:02d}甲行', a, font)
                 + inner_para(f'Ｂ{i:02d}乙行', bsz, font)
                 + inner_para(f'Ｃ{i:02d}丙行', bsz, font))
        r = rpr("21", MINCHO)
        body.append(
            f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:noProof/></w:rPr>{tb(pid, marks)}</w:r>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>基準{i:02d}</w:t></w:r></w:p>')
        pid += 1
    body.append(pg.sectpr())
    pg.write_docx(DOCX, pg.doc(''.join(body)))
    print('wrote', DOCX)


if __name__ == '__main__':
    build()
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
    doc.ExportAsFixedFormat(PDF, 17)
    doc.Close(False)
    word.Quit()
    import fitz
    d = fitz.open(PDF)
    sys.stdout.reconfigure(encoding='utf-8')
    print('pages:', len(d), ' SG0:', SG0)
    print(f'{"cfg":>8} {"font":>4} | {"A_off":>7} {"gapAB":>7} {"gapBC":>7}  (szA, szB)')
    for pno in range(len(d)):
        rects = [dr['rect'] for dr in d[pno].get_drawings()
                 if 100 < dr['rect'].width < 500 and 40 < dr['rect'].height < 200]
        marks = {}
        for b in d[pno].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                for sp in l['spans']:
                    t = sp['text'].replace(' ', '')
                    for k in 'ＡＢＣ':
                        if t.startswith(k):
                            marks[k] = (sp['origin'][1], round(sp['size'], 2))
        if not rects or len(marks) < 3:
            print(f'p{pno+1}: incomplete'); continue
        top = rects[0].y0
        a, b_, c = marks['Ａ'], marks['Ｂ'], marks['Ｃ']
        cfg = CONFIGS[pno]
        print(f'{cfg[0]+"/"+cfg[1]:>8} {"G" if cfg[2]==GOTHIC else "M":>4} | '
              f'{a[0]-top:7.2f} {b_[0]-a[0]:7.2f} {c[0]-b_[0]:7.2f}  ({a[1]}, {b_[1]})')

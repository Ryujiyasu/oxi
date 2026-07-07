# -*- coding: utf-8 -*-
"""Controlled sweep: WHERE does Word anchor a paragraph-relative wrapNone
textbox vertically? (The 1ec1 wall: per-box anchor errors −2..+1 / −4.8
dominate its SSIM; boxes sharing one anchor paragraph still differ.)

Per config page: [2 filler paras][ANCHOR para (10.5pt, marker 基準)
carrying the box][follower]. The box (wrapNone, compatLnSpc=1, tIns=0,
border ON so the PDF shows the frame rect) holds one marker para.

Measured per config (Word PDF): frame rect y0 (drawings), box first-line
baseline, anchor-para baseline. Derive: rect_y0 − anchor_baseline vs
posV; first_baseline − rect_y0 vs the law asc(first).

Grid: posV {0, 10, 22.45, 50} × first-para spec {auto10.5, exact480-14,
exact340-10.5} + an OVERLAP pair (two boxes on one anchor, the 1ec1
footer cluster shape).

Run: python tools/metrics/_txbx_anchor_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'gridquant')
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, 'txanchor.docx')
PDF = os.path.join(OUTDIR, 'txanchor.pdf')
MINCHO = pg.MINCHO
GOTHIC = "ＭＳ ゴシック"
esc = pg.esc
EMU = 12700

WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def rpr(sz, font=GOTHIC):
    return (f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{sz}"/>')


def inner(txt, sz, spacing):
    r = rpr(sz)
    return (f'<w:p><w:pPr><w:snapToGrid w:val="0"/>{spacing}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')


def tbox(pid, paras, pos_v_pt, cy_pt=40.0, cx_pt=200.0, pos_h_pt=30.0):
    cx, cy = int(cx_pt * EMU), int(cy_pt * EMU)
    posv = int(pos_v_pt * EMU)
    posh = int(pos_h_pt * EMU)
    choice = (
        f'<w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="margin"><wp:posOffset>{posh}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{posv}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        f'<wp:docPr id="{pid}" name="TA{pid}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic {A_NS}>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        '</wps:spPr>'
        f'<wps:txbx><w:txbxContent>{paras}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="18000" tIns="0" '
        'rIns="18000" bIns="0" anchor="t" anchorCtr="0" compatLnSpc="1">'
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


AUTO = ''
EX480 = '<w:spacing w:line="480" w:lineRule="exact"/>'
EX340 = '<w:spacing w:line="340" w:lineRule="exact"/>'

# (posV, first-para sz halfpt, first-para spacing, label)
CONFIGS = [
    (0.0,   '21', AUTO,  'pv0_auto105'),
    (10.0,  '21', AUTO,  'pv10_auto105'),
    (22.45, '21', AUTO,  'pv22_auto105'),
    (50.0,  '21', AUTO,  'pv50_auto105'),
    (22.45, '28', EX480, 'pv22_ex480_14'),
    (22.45, '21', EX340, 'pv22_ex340_105'),
    (10.0,  '28', EX480, 'pv10_ex480_14'),
]


def build():
    body = []
    pid = 500
    for i, (pv, sz, sp, _label) in enumerate(CONFIGS):
        if i:
            body.append(pagebreak())
        body.append(para(f'前{i:02d}あ'))
        body.append(para(f'前{i:02d}い'))
        marks = inner(f'Ｂ{i:02d}箱行', sz, sp)
        r = rpr('21', MINCHO)
        body.append(
            f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:noProof/></w:rPr>{tbox(pid, marks, pv)}</w:r>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>基準{i:02d}行</w:t></w:r></w:p>')
        pid += 1
        body.append(para(f'後{i:02d}う'))
    # overlap pair page (the 1ec1 footer cluster shape): two boxes, one anchor
    body.append(pagebreak())
    i = len(CONFIGS)
    marks1 = inner(f'Ｂ{i:02d}箱行', '21', AUTO)
    marks2 = inner(f'Ｃ{i:02d}箱行', '21', AUTO)
    r = rpr('21', MINCHO)
    body.append(para(f'前{i:02d}あ'))
    body.append(
        f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:noProof/></w:rPr>{tbox(700, marks1, 20.0, 40.0, 200.0, 30.0)}'
        f'{tbox(701, marks2, 25.0, 40.0, 200.0, 120.0)}</w:r>'
        f'<w:r><w:rPr>{r}</w:rPr><w:t>基準{i:02d}行</w:t></w:r></w:p>')
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
    print('pages:', len(d))
    print(f'{"label":<16} {"posV":>6} | {"anch_base":>9} {"rect_y0":>8} {"box_base":>8} | {"rect-anchb":>10} {"base-rect":>9}')
    labels = [c[3] for c in CONFIGS] + ['overlap']
    for pno in range(len(d)):
        rects = sorted((dr['rect'] for dr in d[pno].get_drawings()
                        if 100 < dr['rect'].width < 350 and 20 < dr['rect'].height < 80),
                       key=lambda r: r.y0)
        anch = None
        boxes = {}
        for b in d[pno].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                for sp in l['spans']:
                    t = sp['text'].replace(' ', '')
                    if t.startswith('基準'):
                        anch = (sp['origin'][1], round(sp['size'], 2))
                    for k in 'ＢＣ':
                        if t.startswith(k):
                            boxes[k] = (sp['origin'][1], round(sp['size'], 2))
        if not rects or anch is None or 'Ｂ' not in boxes:
            print(f'p{pno+1} {labels[pno] if pno < len(labels) else "?"}: incomplete')
            continue
        pv = CONFIGS[pno][0] if pno < len(CONFIGS) else 20.0
        r0 = rects[0]
        bb = boxes['Ｂ'][0]
        print(f'{labels[pno]:<16} {pv:>6.2f} | {anch[0]:>9.2f} {r0.y0:>8.2f} {bb:>8.2f} | '
              f'{r0.y0 - anch[0]:>10.2f} {bb - r0.y0:>9.2f}')
        if 'Ｃ' in boxes and len(rects) > 1:
            r1 = rects[1]
            print(f'{"  overlap-C":<16} {25.0:>6.2f} | {anch[0]:>9.2f} {r1.y0:>8.2f} {boxes["Ｃ"][0]:>8.2f} | '
                  f'{r1.y0 - anch[0]:>10.2f} {boxes["Ｃ"][0] - r1.y0:>9.2f}')

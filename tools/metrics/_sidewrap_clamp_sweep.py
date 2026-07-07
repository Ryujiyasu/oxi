# -*- coding: utf-8 -*-
"""Controlled sweep: wrapSquare float at the page bottom — does Word PUSH
the anchor paragraph to the next page (probeximgfloat behavior) or CLAMP
the float up against the physical page bottom keeping the anchor (2ea81a
level-8 behavior)? Derives the clamp-vs-push discriminator for the S758
far-float wall.

Per config page: [filler exact para parks the cursor][anchor para with a
wrapSquare wps textbox marker + anchor text][3 follower paras]. Word truth
via PDF: box marker page+y (clamped iff drawn bottom == physical page
height and above natural), anchor/follower text page.

Grid: anchor_y {700, 750} x posV {0, 30, 60, 120, 170}pt x cy {58, 113}pt.

Run: python tools/metrics/_sidewrap_clamp_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'gridquant')
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, 'swclamp.docx')
PDF = os.path.join(OUTDIR, 'swclamp.pdf')
MINCHO = pg.MINCHO
esc = pg.esc
EMU = 12700

WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def rpr(sz="21"):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')


def para(txt, spacing=''):
    r = rpr()
    return (f'<w:p><w:pPr>{spacing}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')


def filler(h_pt, tag):
    return para(f'Ｆ{tag}', f'<w:spacing w:line="{int(h_pt*20)}" w:lineRule="exact"/>')


def pagebreak():
    return ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:br w:type="page"/></w:r></w:p>')


def sq_txbx(pid, marker, pos_v_pt, cy_pt, cx_pt=100.0):
    cx, cy = int(cx_pt * EMU), int(cy_pt * EMU)
    posv = int(pos_v_pt * EMU)
    r = rpr("16")
    inner = (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{r}</w:rPr><w:t>{esc(marker)}</w:t></w:r></w:p>')
    choice = (
        f'<w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{posv}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="{pid}" name="SW{pid}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic {A_NS}>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        '</wps:spPr>'
        f'<wps:txbx><w:txbxContent>{inner}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="18000" tIns="18000" '
        'rIns="18000" bIns="18000" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>')
    return ('<mc:AlternateContent '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<mc:Choice Requires="wps">{choice}</mc:Choice>'
            f'<mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>')


ANCHOR_Y = [700.0, 750.0]
POSV = [0.0, 30.0, 60.0, 120.0, 170.0]
CY = [58.0, 113.0]
PAGE_TOP = 56.7  # top margin 1134tw

configs = []


def build():
    body = []
    pid = 100
    first = True
    for ay in ANCHOR_Y:
        for pv in POSV:
            for cy in CY:
                if not first:
                    body.append(pagebreak())
                first = False
                idx = len(configs)
                mk = f'Ｂ{idx:02d}'
                configs.append((ay, pv, cy, mk))
                body.append(filler(ay - PAGE_TOP, f'ｘ{idx:02d}'))
                r = rpr()
                body.append(
                    f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
                    f'<w:r><w:rPr><w:noProof/></w:rPr>{sq_txbx(pid, mk, pv, cy)}</w:r>'
                    f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">Ａ{idx:02d}アンカー行</w:t></w:r></w:p>')
                pid += 1
                for j in range(3):
                    body.append(para(f'Ｃ{idx:02d}続{j}'))
    body.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" '
                'w:header="851" w:footer="567" w:gutter="0"/>'
                '<w:docGrid w:type="lines" w:linePitch="330"/></w:sectPr>')
    pg.write_docx(DOCX, pg.doc(''.join(body)))
    print('wrote', DOCX, len(configs), 'configs')


if __name__ == '__main__':
    build()
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
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
    locs = {}
    for p in range(len(d)):
        for b in d[p].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                txt = ''.join(s['text'] for s in l.get('spans', [])).replace(' ', '')
                for key in ('Ｂ', 'Ａ', 'Ｃ'):
                    if txt.startswith(key):
                        locs.setdefault(txt[:3], []).append((p + 1, round(l['bbox'][1], 2)))
    PAGE_H = 841.9
    print(f'{"cfg":>4} {"a_y":>5} {"posV":>5} {"cy":>5} | box(page,y_top) nat_top clamp_top | anchor_pg follow_pg verdict')
    for idx, (ay, pv, cy, mk) in enumerate(configs):
        cfg_page = idx + 1
        box = (locs.get(mk) or [(None, None)])[0]
        anc = (locs.get(f'Ａ{idx:02d}') or [(None, None)])[0]
        fol = (locs.get(f'Ｃ{idx:02d}') or [(None, None)])[0]
        nat_top = ay + pv  # anchor line top ~ ay; box ink ~ +3.5 for insets
        clamp_top = PAGE_H - cy
        verdict = '?'
        if box[0] is not None:
            box_top = box[1] - 3.5
            on_next = box[0] > cfg_page
            if anc[0] is not None and anc[0] > cfg_page:
                verdict = 'PUSH(anchor moved)'
            elif on_next:
                verdict = 'BOX-NEXT-PAGE(anchor kept)'
            elif abs(box_top - clamp_top) < 5:
                verdict = 'CLAMP'
            elif abs(box_top - nat_top) < 6:
                verdict = 'NATURAL'
            else:
                verdict = f'OTHER(top={box_top:.1f})'
        print(f'{idx:4d} {ay:5.0f} {pv:5.0f} {cy:5.0f} | p{box[0]} y={box[1]}  nat={nat_top:.0f} clamp={clamp_top:.0f} | p{anc[0]} p{fol[0]} {verdict}')

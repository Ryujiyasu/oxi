# -*- coding: utf-8 -*-
"""Controlled sweep: Word's anchored-float page-boundary rule — CLAMP the
float up to fit the page vs PUSH the anchor paragraph to the next page
(the 2ea81a pi26-vs-pi27/28 contradiction: pi26's floats clamp [Ｒ７.６ box
bottom = exactly the page height], pi27/28 get pushed).

Per config page: [filler para, exact spacing to park the cursor] [anchor
para: empty run + wp:anchor wrapNone textbox with marker text] [follower
para]. Word truth: the box marker's rendered page+Y (clamp?) and the
follower's page (push?).

Run: python tools/metrics/_anchorclamp_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), "gridquant")
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, "anchorclamp.docx")
PDF = os.path.join(OUTDIR, "anchorclamp.pdf")

esc = pg.esc
MINCHO = pg.MINCHO
EMU = 12700  # per pt

WP_NS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"')
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

def rpr(sz="21"):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')

def para(txt, extra=''):
    r = rpr()
    return (f'<w:p><w:pPr>{extra}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def anchored_box(pid, marker, pos_v_pt, cy_pt, cx_pt=42.0):
    cx, cy = int(cx_pt * EMU), int(cy_pt * EMU)
    posv = int(pos_v_pt * EMU)
    r = rpr("16")
    inner = (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{r}</w:rPr><w:t>{esc(marker)}</w:t></w:r></w:p>')
    choice = (
        f'<w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:posOffset>3175000</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{posv}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        f'<wp:docPr id="{pid}" name="AB{pid}"/><wp:cNvGraphicFramePr/>'
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

def anchor_para(pid, marker, pos_v_pt, cy_pt):
    r = rpr("28")
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:noProof/></w:rPr>{anchored_box(pid, marker, pos_v_pt, cy_pt)}</w:r></w:p>')

def pagebreak_para():
    r = rpr()
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:br w:type="page"/></w:r></w:p>')

SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="737" w:right="1134" w:bottom="397" w:left="1134" w:header="680" w:footer="907" w:gutter="0"/>'
          '<w:docGrid w:type="lines" w:linePitch="323"/></w:sectPr>')

# park the cursor: filler = one exact-spacing para of height F (tw)
def filler(h_pt):
    return para('Ｆ', f'<w:spacing w:line="{int(h_pt*20)}" w:lineRule="exact"/>')

ANCHOR_TOPS = [740, 770, 795]   # target anchor-para y (page_top=36.85)
POSV = [8.0, 47.5, 128.0]
CYS = [30.4, 57.85]
configs = []
body = []
pid = 100
for at in ANCHOR_TOPS:
    for pv in POSV:
        for cy in CYS:
            if body:
                body.append(pagebreak_para())
            marker = f'Ｂ{len(configs):02d}'
            configs.append((at, pv, cy, marker))
            body.append(filler(at - 36.85))
            body.append(anchor_para(pid, marker, pv, cy))
            pid += 1
            body.append(para('あと続き'))
body.append(SECTPR)

pg.write_docx(DOCX, pg.doc(''.join(body)))
print('wrote', DOCX, 'configs:', len(configs))

import win32com.client
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
doc.ExportAsFixedFormat(PDF, 17)
doc.Close(False)
word.Quit()
print('exported', PDF)

import fitz
d = fitz.open(PDF)
sys.stdout.reconfigure(encoding='utf-8')
print('pages:', len(d))
# locate each marker + the follower per config
locs = {}
for p in range(len(d)):
    for b in d[p].get_text('dict')['blocks']:
        for l in b.get('lines', []):
            txt = ''.join(s['text'] for s in l.get('spans', [])).replace(' ', '')
            y = round(l['bbox'][1], 2)
            for (_, _, _, mk) in configs:
                if mk in txt:
                    locs.setdefault(mk, []).append(('box', p+1, y))
            if 'あと続き' in txt:
                locs.setdefault(f'F{p}', []).append(('follow', p+1, y))
print(f'{"anchor_y":>8} {"posV":>6} {"cy":>6}  box(page,y)      natural_boxY  clamped_to  verdict')
follows = sorted([v for k, vs in locs.items() if k.startswith('F') for v in vs], key=lambda t: (t[1], t[2]))
for idx, (at, pv, cy, mk) in enumerate(configs):
    hits = locs.get(mk, [])
    box = hits[0] if hits else None
    nat = round(at + pv, 1)
    clamp = round(841.9 - cy, 1)
    if box:
        _, bp, by = box
        # box marker ink ~ box top + tIns(1.4pt)+~2 → boxY ≈ by-3.5
        boxy = round(by - 3.5, 1)
        verdict = 'NAT' if abs(boxy - nat) < 4 else ('CLAMP' if abs(boxy - clamp) < 4 else ('PUSHED?' if bp > (idx // 1) % 999 + 0 else '??'))
        print(f'{at:8.0f} {pv:6.1f} {cy:6.2f}  p{bp} y={by:7.2f}  nat={nat:7.1f}  clamp={clamp:7.1f}  {verdict}')
    else:
        print(f'{at:8.0f} {pv:6.1f} {cy:6.2f}  MISSING')

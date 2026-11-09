"""Faithful-slice probe: where is the fit limit for a wrapSquare text box near the page bottom
when the footer is tall? (zh dev forms__00342784 p3: a 151pt box anchored 43.8pt below its
paragraph reaches y 727.2, past the footer's top ~714, and Word keeps box + anchor on p3; Oxi's
S758 push sent both to p4.)

Arms: the document itself, with the box's posOffset (positionV relativeFrom=paragraph) changed so
its natural bottom moves down in steps; for each, Word's PDF gives the box rectangle (its border)
and the page holding the anchor text «本表所填写».
Usage: python _pb_square_box_footer_gen.py <src.docx> <out_dir>
"""
import sys, os, re, zipfile

sys.stdout.reconfigure(encoding='utf-8')
src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
z = zipfile.ZipFile(src)
parts = {it.filename: z.read(it.filename) for it in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
anchor_txt = '本表所填写'
i = doc.find(anchor_txt)
j = doc.rfind('<wp:positionV relativeFrom="paragraph"><wp:posOffset>', 0, i)
k = doc.find('</wp:posOffset>', j)
pre = doc[:j] + '<wp:positionV relativeFrom="paragraph"><wp:posOffset>'
post = doc[k:]
orig = int(doc[j + len('<wp:positionV relativeFrom="paragraph"><wp:posOffset>'):k])
# VML fallback carries its own top; leave it (Word reads the wps choice)
OFFS = [orig + d * 12700 for d in (0, 20, 40, 60, 80, 100, 110, 120)]
import fitz, win32com.client
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    for off in OFFS:
        xml = pre + str(off) + post
        dst = os.path.join(out, f'off{off}.docx')
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items():
                o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
        pdf = os.path.join(out, f'off{off}.pdf')
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
        P = fitz.open(pdf)
        where = None
        for pi, pg in enumerate(P):
            ys = [l['spans'][0]['origin'][1] for b in pg.get_text('dict')['blocks'] for l in b.get('lines', [])
                  if anchor_txt in ''.join(s['text'] for s in l['spans'])]
            if ys:
                boxes = [r for r in (dr['rect'] for dr in pg.get_drawings()) if r.width > 300 and r.height > 100]
                where = (pi + 1, round(ys[0], 1), [(round(b.y0, 1), round(b.y1, 1)) for b in boxes])
        print(f'posOffset {off / 12700:6.1f}pt  pages {len(P)}  anchor page/baseline/box {where}')
finally:
    w.Quit()

# -*- coding: utf-8 -*-
"""Does a paragraph's space before survive at the top of a split row's continuation?

blind-J en/creative__04637c5f: the body is one table (a repeated header row and
one content row running over three pages). On page 2 the right cell continues
with a Heading2 paragraph («REGULAR WEB TRAFFIC», style before=720 = 36pt).
Word draws it at 334.8, 36pt below the row's continuation top (the left cell's
image at 298.8); Oxi draws it at 299 -- the space before dropped as at a body
page top.

Faithful slice (the document itself, styles.xml edited): Heading2's before =
0 / 18 / 36pt. Readout: page-2 top of «REGULAR WEB TRAFFIC» and of the left
cell's first image.
"""
import glob, os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = glob.glob('pipeline_data/docx_corpus/en/creative/04637c5f*.docx')[0]
OUT = 'tests/fixtures/split_row_before'; os.makedirs(OUT, exist_ok=True)
H2 = b'<w:spacing w:before="720"/><w:outlineLvl w:val="1"/>'


def build(before):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, f'b{before}.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/styles.xml':
                assert data.count(H2) == 1
                data = data.replace(H2, b'<w:spacing w:before="%d"/><w:outlineLvl w:val="1"/>' % before)
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    for b in [int(a) for a in sys.argv[1:]] or [0, 360, 720]:
        at = build(b)
        app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
        finally:
            app.Quit()
        doc = pymupdf.open(at[:-5] + '.pdf')
        for pi, P in enumerate(doc):
            hits = [round(l['bbox'][1], 1) for blk in P.get_text('dict')['blocks'] for l in blk.get('lines', [])
                    if 'REGULAR WEB' in ''.join(s['text'] for s in l['spans'])]
            if hits:
                imgs = sorted(round(P.get_image_bbox(im).y0, 1) for im in P.get_images(full=True))
                print(f'before {b / 20:g}pt: page {pi + 1} REGULAR top {hits}, images {imgs[:2]}', flush=True)

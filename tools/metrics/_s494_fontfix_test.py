# -*- coding: utf-8 -*-
"""Test: render gen2_045 p1 via MuPDF with the REAL font (Calibri) vs the current
Mincho-substitution, compare SSIM to word_png. If real-font SSIM >> Mincho SSIM, the
per-glyph gate's biggest 'drops' (Latin docs) are a font-substitution artifact, not
real position errors. cp932-safe."""
import os, json, tempfile
import numpy as np, fitz
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DPI = 150
FONTS = {
    'm': ('C:/Windows/Fonts/msmincho.ttc', 0.8594),
    'g': ('C:/Windows/Fonts/msgothic.ttc', 0.8594),
    'calibri': ('C:/Windows/Fonts/calibri.ttf', 0.75),
    'times': ('C:/Windows/Fonts/times.ttf', 0.8911),
    'arial': ('C:/Windows/Fonts/arial.ttf', 0.9053),
}
_F = {k: fitz.Font(fontfile=v[0]) for k, v in FONTS.items()}


def pick(fam, real):
    fam = fam or ''
    if not real:
        if 'Goth' in fam or 'Gothic' in fam or 'ゴシック' in fam:
            return 'g'
        return 'm'
    if 'Calibri' in fam:
        return 'calibri'
    if 'Times' in fam:
        return 'times'
    if 'Arial' in fam:
        return 'arial'
    if 'Goth' in fam or 'Gothic' in fam or 'ゴシック' in fam:
        return 'g'
    return 'm'


def render(glyphs_json, page_idx, real, w, h):
    g = json.load(open(glyphs_json, encoding='utf-8'))['pages'][page_idx]
    doc = fitz.open()
    pg = doc.new_page(width=w, height=h)
    for k in FONTS:
        pg.insert_font(fontname=k, fontfile=FONTS[k][0])
    for gl in g['glyphs']:
        fn = pick(gl.get('font_family'), real)
        K = FONTS[fn][1]
        fs = gl['font_size']
        try:
            pg.insert_text((gl['x'], gl['top'] + fs * K), gl['char'], fontname=fn, fontsize=fs, color=(0, 0, 0))
        except Exception:
            pass
    png = tempfile.mktemp(suffix='.png')
    pg.get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72)).save(png)
    return png


def rgb(wp, op):
    a = np.array(Image.open(wp).convert('RGB')); b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))


did = 'gen2_045_Training_Report'
oxi_glyphs = 'c:/tmp/og45.json'
wpng = os.path.join(ROOT, 'pipeline_data', 'word_png', did, 'page_0001.png')
wa = Image.open(wpng); w, h = wa.width * 72 / DPI, wa.height * 72 / DPI
mincho = render(oxi_glyphs, 0, False, w, h)
realf = render(oxi_glyphs, 0, True, w, h)
print('gen2_045 p1  Mincho-sub SSIM %.4f  |  real-font SSIM %.4f  (dwrite gate was 0.9385ish)'
      % (rgb(wpng, mincho), rgb(wpng, realf)))

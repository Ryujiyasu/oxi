# -*- coding: utf-8 -*-
"""S494 — per-glyph AA-FREE gate. Render Oxi via the DWRITE --dump-glyphs (gate-render
positions, incl. autoSpace+charGrid) through the SAME PyMuPDF that makes word_png, so
the text AA matches Word by construction and the SSIM measures pure POSITION/LAYOUT
fidelity (the dwrite-gate's DirectWrite-vs-MuPDF AA-texture confound is removed).

This is the AA-free judge for the cell-Y / line-height position track (b35 S451,
de6e32 top-region drift) that the AA-confounded dwrite SSIM masks. Fixed baseline k
(deltas are robust to k since it cancels before/after a fix).

Usage:
  python tools/metrics/per_glyph_gate.py <docid> [env_KEY=VAL ...]   # one doc, SSIM
  python tools/metrics/per_glyph_gate.py --ab <docid> <ENV_DISABLE>  # A/B a fix env
"""
import os, sys, json, subprocess, tempfile, glob
import numpy as np
import fitz
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
MINCHO = 'C:/Windows/Fonts/msmincho.ttc'; GOTHIC = 'C:/Windows/Fonts/msgothic.ttc'
DOCXDIRS = [os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx'),
            os.path.join(ROOT, 'pipeline_data', 'docx')]
DPI = 150
_FMIN = fitz.Font(fontfile=MINCHO); _FGOT = fitz.Font(fontfile=GOTHIC)
K = 0.859  # fixed baseline ratio (fitz MS Mincho ascender); cancels in A/B deltas

def docx_for(did):
    for d in DOCXDIRS:
        for p in glob.glob(os.path.join(d, did + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == did:
                return p
        f = glob.glob(os.path.join(d, did + '*.docx'))
        if f:
            return f[0]
    return None

def _font(fam):
    fam = fam or ''
    if 'Goth' in fam or 'ゴシック' in fam or 'Gothic' in fam:
        return 'g', _FGOT
    return 'm', _FMIN

def render_per_glyph(docx, env=None, out_png=None):
    e = dict(os.environ)
    if env:
        e.update(env)
    gj = tempfile.mktemp(suffix='.json'); lj = tempfile.mktemp(suffix='.json')
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(), str(DPI), '--dump-glyphs=' + gj],
                   capture_output=True, timeout=300, env=e)
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(), str(DPI), '--dump-layout=' + lj],
                   capture_output=True, timeout=300, env=e)
    g = json.load(open(gj, encoding='utf-8')); lay = json.load(open(lj, encoding='utf-8'))
    os.unlink(gj); os.unlink(lj)
    borders = {pi: [el for el in p['elements'] if el.get('type') == 'border']
               for pi, p in enumerate(lay['pages'])}
    out_png = out_png or tempfile.mktemp(suffix='.png')
    doc = fitz.open()
    for pi, page in enumerate(g['pages']):
        pg = doc.new_page(width=page['width'], height=page['height'])
        pg.insert_font(fontname='m', fontfile=MINCHO); pg.insert_font(fontname='g', fontfile=GOTHIC)
        for gl in page['glyphs']:
            fn, fo = _font(gl.get('font_family'))
            fs = gl['font_size']
            try:
                pg.insert_text((gl['x'], gl['top'] + fs * K), gl['char'], fontname=fn, fontsize=fs, color=(0, 0, 0))
            except Exception:
                pass
        for el in borders.get(pi, []):
            x, y, w, h = el['x'], el['y'], el['w'], el['h']
            pg.draw_line((x, y), (x + w, y + h), color=(0, 0, 0), width=0.75)
        pix = pg.get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72))
        if pi == 0:
            pix.save(out_png)
    doc.close()
    return out_png

def gate_ssim(docx, env=None):
    wp = os.path.join(ROOT, 'pipeline_data', 'word_png',
                      os.path.splitext(os.path.basename(docx))[0], 'page_0001.png')
    if not os.path.exists(wp):
        return None
    op = render_per_glyph(docx, env)
    a = np.array(Image.open(wp).convert('RGB')); b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))

if __name__ == '__main__':
    if sys.argv[1] == '--ab':
        did, disable_env = sys.argv[2], sys.argv[3]
        dx = docx_for(did)
        on = gate_ssim(dx, {})                       # fix ON (default)
        off = gate_ssim(dx, {disable_env: '1'})      # fix OFF
        print(f'{did} per-glyph gate: FIX-ON {on:.4f} | FIX-OFF(={disable_env}) {off:.4f} | delta {on-off:+.4f}')
    else:
        did = sys.argv[1]
        env = dict(kv.split('=', 1) for kv in sys.argv[2:] if '=' in kv)
        dx = docx_for(did)
        print(f'{did} per-glyph gate SSIM = {gate_ssim(dx, env):.4f}')

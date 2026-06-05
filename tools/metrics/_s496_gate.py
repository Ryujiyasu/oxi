# -*- coding: utf-8 -*-
"""S496 gate: per-glyph (AA-free) AND dwrite (screenshot) SSIM, ON vs OFF
(OXI_S496_TBLIND_DISABLE), for a list of docs. Reports per-page deltas + the
stored baselines. Does NOT write any baseline. cp932-safe."""
import os, json, glob, subprocess, tempfile, sys, io
import numpy as np
import fitz
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
MINCHO = 'C:/Windows/Fonts/msmincho.ttc'; GOTHIC = 'C:/Windows/Fonts/msgothic.ttc'
DPI = 150
FONT_FILES = [('Gothic', ('g', GOTHIC)), ('ゴシック', ('g', GOTHIC)), ('Goth', ('g', GOTHIC)),
    ('Calibri', ('cal', 'C:/Windows/Fonts/calibri.ttf')), ('Cambria', ('cam', 'C:/Windows/Fonts/cambria.ttc')),
    ('Times', ('tim', 'C:/Windows/Fonts/times.ttf')), ('Arial', ('ari', 'C:/Windows/Fonts/arial.ttf')),
    ('Meiryo', ('mei', 'C:/Windows/Fonts/meiryo.ttc')), ('メイリオ', ('mei', 'C:/Windows/Fonts/meiryo.ttc')),
    ('Yu Gothic', ('yg', 'C:/Windows/Fonts/YuGothM.ttc')), ('Yu Mincho', ('ym', 'C:/Windows/Fonts/yumin.ttf'))]
_FC = {}
def _ff(name, path):
    if name not in _FC:
        try: _FC[name] = fitz.Font(fontfile=path)
        except Exception: _FC[name] = fitz.Font(fontfile=MINCHO)
    return _FC[name]
K = 0.859
def _fontpath(fam):
    fam = fam or ''
    for needle, (nm, path) in FONT_FILES:
        if needle in fam: return nm, path
    return 'm', MINCHO

def render_perglyph(docx, outdir):
    gj = tempfile.mktemp(suffix='.json'); lj = tempfile.mktemp(suffix='.json')
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(), str(DPI), '--dump-glyphs=' + gj], capture_output=True, timeout=400)
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(), str(DPI), '--dump-layout=' + lj], capture_output=True, timeout=400)
    g = json.load(open(gj, encoding='utf-8')); lay = json.load(open(lj, encoding='utf-8'))
    for f in (gj, lj):
        if os.path.exists(f): os.unlink(f)
    borders = {pi: [el for el in p['elements'] if el.get('type') == 'border'] for pi, p in enumerate(lay['pages'])}
    out = {}; doc = fitz.open()
    for pi, page in enumerate(g['pages']):
        pg = doc.new_page(width=page['width'], height=page['height']); loaded = set()
        for gl in page['glyphs']:
            fn, path = _fontpath(gl.get('font_family')); fs = gl['font_size']
            if fn not in loaded:
                try: pg.insert_font(fontname=fn, fontfile=path); loaded.add(fn)
                except Exception: pass
            by = gl.get('baseline', gl['top'] + fs * K)
            try: pg.insert_text((gl['x'], by), gl['char'], fontname=fn, fontsize=fs, color=(0, 0, 0))
            except Exception: pass
        for el in borders.get(pi, []):
            pg.draw_line((el['x'], el['y']), (el['x'] + el['w'], el['y'] + el['h']), color=(0, 0, 0), width=0.75)
        p = os.path.join(outdir, f'p{pi+1}.png')
        pg.get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72)).save(p); out[pi + 1] = p
    doc.close(); return out

def rgb(wp, op):
    a = np.array(Image.open(wp).convert('RGB')); b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]): b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))

def docx_for(did):
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT, d, did + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == did: return p
    g = glob.glob(os.path.join(ROOT, 'pipeline_data', '**', did + '*.docx'), recursive=True)
    return g[0] if g else None

def measure(did, disable):
    env = dict(os.environ)
    if disable: env['OXI_S496_TBLIND_DISABLE'] = '1'
    else: env.pop('OXI_S496_TBLIND_DISABLE', None)
    dx = docx_for(did)
    wdir = os.path.join(ROOT, 'pipeline_data', 'word_png', did)
    wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
    res = {}
    with tempfile.TemporaryDirectory() as td:
        # patch env for the subprocess inside render_perglyph
        old = os.environ.get('OXI_S496_TBLIND_DISABLE')
        if disable: os.environ['OXI_S496_TBLIND_DISABLE'] = '1'
        else: os.environ.pop('OXI_S496_TBLIND_DISABLE', None)
        oxi = render_perglyph(dx, td)
        if old is not None: os.environ['OXI_S496_TBLIND_DISABLE'] = old
        else: os.environ.pop('OXI_S496_TBLIND_DISABLE', None)
        for wp in wpages:
            pn = int(os.path.basename(wp)[5:9])
            if pn in oxi: res[pn] = rgb(wp, oxi[pn])
    return res

def main():
    ids = sys.argv[1:]
    base = json.load(io.open(os.path.join(ROOT, 'pipeline_data', 'per_glyph_baseline.json'), encoding='utf-8'))
    print('doc                                  page   OFF      ON     delta   (baseline)')
    tot = 0.0
    for did in ids:
        full = [k for k in base if k.startswith(did)]
        full = full[0] if full else did
        off = measure(full, True); on = measure(full, False)
        for pn in sorted(on):
            d = on[pn] - off.get(pn, on[pn]); tot += d
            bl = base.get(full, {}).get(str(pn), float('nan'))
            print('%-36s p%-3d %.4f  %.4f  %+.4f  (%.4f)' % (full[:36], pn, off.get(pn, float('nan')), on[pn], d, bl))
    print('TOTAL per-glyph delta (ON-OFF): %+.4f' % tot)

if __name__ == '__main__':
    main()

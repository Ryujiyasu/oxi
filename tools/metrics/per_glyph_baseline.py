# -*- coding: utf-8 -*-
"""S494 (decision A) — compute the full-corpus PER-GLYPH AA-FREE baseline and adopt it
as the ship gate. For every doc in ssim_baseline.json, render Oxi via the DWRITE
--dump-glyphs (gate-render positions incl. autoSpace+charGrid) through the SAME PyMuPDF
that makes word_png, per page, and compute RGB SSIM vs word_png. This measures pure
POSITION fidelity (AA matched by construction) — what the browser/Canvas product
actually renders — removing the DirectWrite-vs-MuPDF AA-texture confound.

Writes pipeline_data/per_glyph_baseline.json and prints the dwrite->per-glyph shift.
Does NOT touch ssim_baseline.json (the old dwrite gate is kept for reference). cp932-safe."""
import os, json, glob, subprocess, tempfile, sys
import numpy as np
import fitz
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
MINCHO = 'C:/Windows/Fonts/msmincho.ttc'; GOTHIC = 'C:/Windows/Fonts/msgothic.ttc'
DOCXDIRS = [os.path.join(ROOT, 'pipeline_data', 'docx'), os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')]
DPI = 150
_FMIN = fitz.Font(fontfile=MINCHO); _FGOT = fitz.Font(fontfile=GOTHIC)
K = 0.859
OLD = json.load(open(os.path.join(ROOT, 'pipeline_data', 'ssim_baseline.json'), encoding='utf-8'))

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

def render_all_pages(docx, outdir):
    """Render every page of Oxi via dwrite dump-glyphs -> MuPDF. Returns {page_no: png}."""
    gj = tempfile.mktemp(suffix='.json'); lj = tempfile.mktemp(suffix='.json')
    try:
        subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(), str(DPI), '--dump-glyphs=' + gj],
                       capture_output=True, timeout=400)
        subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(), str(DPI), '--dump-layout=' + lj],
                       capture_output=True, timeout=400)
        g = json.load(open(gj, encoding='utf-8')); lay = json.load(open(lj, encoding='utf-8'))
    except Exception:
        return {}
    finally:
        for f in (gj, lj):
            if os.path.exists(f):
                os.unlink(f)
    borders = {pi: [el for el in p['elements'] if el.get('type') == 'border'] for pi, p in enumerate(lay['pages'])}
    out = {}
    doc = fitz.open()
    for pi, page in enumerate(g['pages']):
        pg = doc.new_page(width=page['width'], height=page['height'])
        pg.insert_font(fontname='m', fontfile=MINCHO); pg.insert_font(fontname='g', fontfile=GOTHIC)
        for gl in page['glyphs']:
            fn, fo = _font(gl.get('font_family')); fs = gl['font_size']
            try:
                pg.insert_text((gl['x'], gl['top'] + fs * K), gl['char'], fontname=fn, fontsize=fs, color=(0, 0, 0))
            except Exception:
                pass
        for el in borders.get(pi, []):
            x, y, w, h = el['x'], el['y'], el['w'], el['h']
            pg.draw_line((x, y), (x + w, y + h), color=(0, 0, 0), width=0.75)
        png = os.path.join(outdir, f'p{pi+1}.png')
        pg.get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72)).save(png)
        out[pi + 1] = png
    doc.close()
    return out

def rgb(wp, op):
    a = np.array(Image.open(wp).convert('RGB')); b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))

def main():
    ids = sorted(OLD)
    if len(sys.argv) > 1 and sys.argv[1] != '--all':
        ids = [x for x in ids if any(x.startswith(p) for p in sys.argv[1:])]
    new = {}; fail = []
    for i, did in enumerate(ids):
        wdir = os.path.join(ROOT, 'pipeline_data', 'word_png', did)
        dx = docx_for(did)
        wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png'))) if os.path.isdir(wdir) else []
        if not dx or not wpages:
            new[did] = OLD[did]; continue
        with tempfile.TemporaryDirectory() as td:
            oxi = render_all_pages(dx, td)
            if not oxi:
                new[did] = OLD[did]; fail.append(did); continue
            pages = {}
            for wp in wpages:
                pn = int(os.path.basename(wp)[5:9])
                if pn in oxi:
                    pages[str(pn)] = rgb(wp, oxi[pn])
                elif str(pn) in OLD[did]:
                    pages[str(pn)] = OLD[did][str(pn)]
            new[did] = pages if pages else OLD[did]
        if (i + 1) % 20 == 0:
            print('...%d/%d' % (i + 1, len(ids)), flush=True)
    outp = os.path.join(ROOT, 'pipeline_data', 'per_glyph_baseline.json')
    json.dump(new, open(outp, 'w', encoding='utf-8'), indent=2, ensure_ascii=False)

    def dm(d): return sum(d.values()) / len(d) if d else 0
    common = [k for k in OLD if k in new and OLD[k] and new[k]]
    om = {k: dm(OLD[k]) for k in common}; nm = {k: dm(new[k]) for k in common}
    mo = sum(om.values()) / len(common); mn = sum(nm.values()) / len(common)
    deltas = sorted((nm[k] - om[k], k) for k in common)
    print('\n=== PER-GLYPH (AA-free) BASELINE vs old dwrite gate (doc-mean) ===')
    print('docs=%d  dwrite-mean %.4f -> per-glyph-mean %.4f  (%+.4f; per-glyph removes AA "help")' % (len(common), mo, mn, mn - mo))
    print('biggest DROPS (AA was helping most):', [(round(d, 3), k[:24]) for d, k in deltas[:8]])
    print('biggest RISES (positions cleaner than AA):', [(round(d, 3), k[:24]) for d, k in deltas[-6:]])
    bot = sorted(common, key=lambda k: nm[k])[:12]
    print('\nWORST per-glyph POSITION fidelity (the grind targets):')
    for k in bot:
        print('  %-40s dwrite %.4f | per-glyph %.4f' % (k[:40], om[k], nm[k]))
    if fail:
        print('render-failed (kept old):', len(fail), fail[:5])
    print('\nwrote pipeline_data/per_glyph_baseline.json  (dwrite ssim_baseline.json untouched)')
    print('DONE')

if __name__ == '__main__':
    main()

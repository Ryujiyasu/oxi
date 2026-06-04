# -*- coding: utf-8 -*-
"""Render Oxi's layout via PyMuPDF (the SAME rasterizer word_png uses), so the
text AA matches Word by construction and the SSIM measures pure LAYOUT fidelity
(no DirectWrite-vs-MuPDF rasterizer confound).

word_png = Word ExportAsFixedFormat(PDF) -> fitz get_pixmap. This builds a PDF
from Oxi's --dump-layout JSON and rasterizes it the same way, so the AA matches
Word by construction. S494 (2026-06-04): confirmed word_png IS MuPDF (not a
screenshot/GDI/ClearType render); the text bottom-N weight/AA cap = the
DirectWrite-vs-MuPDF rasterizer coverage difference.

S494 step 2: now uses the GDI renderer's --dump-glyphs (EXACT per-char x via
GetCharWidthW + char_spacing cumulative) so glyphs are placed faithfully to Oxi's
GDI render, NOT fitz's natural advance. SCOPE/CAVEAT: the GDI renderer is TIGHT —
it does NOT apply the DirectWrite/Word CJK<->half-width-digit 1/4em autoSpace that
the dwrite gate renderer (and word_png) HAVE. So this matches word_png for pure-CJK
BODY text (0e7af +0.006, 15076df +0.013 vs dwrite) but UNDER-scores digit/table/
charGrid docs (b35 −0.048, a47e −0.010) by the GDI-vs-dwrite autoSpace gap. A FULLY
faithful render needs the DWRITE renderer's cluster positions (IDWriteTextLayout::
GetClusterMetrics, which include DirectWrite's autoSpace) — the next refinement.
NOT a gate yet; use as a render-truth DIAGNOSTIC, AA-confound removed.

Usage: python tools/metrics/oxi_via_mupdf.py <docid>
       python tools/metrics/oxi_via_mupdf.py --bench
"""
import json, os, subprocess, sys, tempfile, glob
import numpy as np
import fitz
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
GDI = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
DW  = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
MINCHO = 'C:/Windows/Fonts/msmincho.ttc'
GOTHIC = 'C:/Windows/Fonts/msgothic.ttc'
DOCXDIRS = [os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx'),
            os.path.join(ROOT, 'pipeline_data', 'docx')]
DPI = 150

# font ascender ratios (calibrated so the MuPDF baseline matches the dwrite glyph row)
_F_MINCHO = fitz.Font(fontfile=MINCHO)
_F_GOTHIC = fitz.Font(fontfile=GOTHIC)

def docx_for(docid):
    for d in DOCXDIRS:
        for p in glob.glob(os.path.join(d, docid + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == docid:
                return p
        f = glob.glob(os.path.join(d, docid + '*.docx'))
        if f:
            return f[0]
    return None

def dump_glyphs(docx):
    """Run --dump-glyphs on the DWRITE renderer: each glyph's EXACT per-char (char,x,
    top) in pt from IDWriteTextLayout::HitTestTextPosition — the gate-render positions
    (DirectWrite, includes autoSpace + charGrid; ≈ Word). Faithful for ALL docs incl.
    tables (the GDI dump-glyphs is also available but uses GetTextExtentExPointW)."""
    fd, jp = tempfile.mkstemp(suffix='.json'); os.close(fd)
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(), str(DPI),
                    '--dump-glyphs=' + jp], capture_output=True, timeout=300)
    d = json.load(open(jp, encoding='utf-8')); os.unlink(jp)
    return d

# also dump borders (the glyph dump is text-only) for a complete page
def dump_layout(docx):
    fd, jp = tempfile.mkstemp(suffix='.json'); os.close(fd)
    subprocess.run([GDI, os.path.abspath(docx), tempfile.mktemp(), str(DPI),
                    '--dump-layout=' + jp], capture_output=True, timeout=300)
    d = json.load(open(jp, encoding='utf-8')); os.unlink(jp)
    return d

def font_for(fam):
    fam = fam or ''
    if 'Goth' in fam or 'ゴシック' in fam or 'Gothic' in fam:
        return 'gothic', _F_GOTHIC
    return 'mincho', _F_MINCHO

def build_pdf_png(docx, out_png):
    """Faithful render: place each glyph at its EXACT per-char (x, top) from the
    GDI --dump-glyphs, baseline = top + fitz_ascender*fs (per-font), + borders from
    --dump-layout. Rasterize via the SAME PyMuPDF as word_png → AA matches by
    construction; SSIM measures Oxi's LAYOUT position fidelity."""
    g = dump_glyphs(docx)
    lay = dump_layout(docx)
    doc = fitz.open()
    borders_by_page = {}
    for pi, page in enumerate(lay['pages']):
        borders_by_page[pi] = [el for el in page['elements'] if el.get('type') == 'border']
    for pi, page in enumerate(g['pages']):
        pg = doc.new_page(width=page['width'], height=page['height'])
        pg.insert_font(fontname='mincho', fontfile=MINCHO)
        pg.insert_font(fontname='gothic', fontfile=GOTHIC)
        for gl in page['glyphs']:
            fn, fobj = font_for(gl.get('font_family'))
            fs = gl['font_size']
            baseline = gl['top'] + fs * fobj.ascender
            try:
                pg.insert_text((gl['x'], baseline), gl['char'], fontname=fn, fontsize=fs, color=(0, 0, 0))
            except Exception:
                pass
        for el in borders_by_page.get(pi, []):
            x, y, w, h = el['x'], el['y'], el['w'], el['h']
            pg.draw_line((x, y), (x + w, y + h), color=(0, 0, 0), width=0.75)
        pix = pg.get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72))
        if pi == 0:
            pix.save(out_png)
    doc.close()
    return out_png

def rgb_ssim(wp, op):
    a = np.array(Image.open(wp).convert('RGB'))
    b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))

def render_dwrite(docx):
    td = tempfile.mkdtemp()
    subprocess.run([DW, os.path.abspath(docx), os.path.join(td, 'o'), str(DPI)], capture_output=True, timeout=300)
    return os.path.join(td, 'o_p1.png')

def bench():
    docs = ['1ec1091177b1_006', '15076df085f5_tokumei_08_09',
            '0e7af1ae8f21_20230331_resources_open_data_contract_sample_00',
            'b35123fe8efc_tokumei_08_01', 'de6e32b5960b_tokumei_08_01-1',
            'a47e6c6b2ca1_order_08', '29dc6e8943fe_order_01']
    print('%-16s %8s %8s %8s' % ('doc', 'dwrite', 'mupdf', 'delta'))
    for did in docs:
        dx = docx_for(did)
        wp = os.path.join(ROOT, 'pipeline_data', 'word_png', did, 'page_0001.png')
        if not dx or not os.path.exists(wp):
            print('%-16s skip' % did[:16]); continue
        mp = build_pdf_png(dx, tempfile.mktemp(suffix='.png'))
        s_mp = rgb_ssim(wp, mp)
        s_dw = rgb_ssim(wp, render_dwrite(dx))
        print('%-16s %8.4f %8.4f %+8.4f' % (did[:16], s_dw, s_mp, s_mp - s_dw))

if __name__ == '__main__':
    if '--bench' in sys.argv:
        bench()
    else:
        did = sys.argv[1]
        dx = docx_for(did)
        out = 'c:/tmp/oxi_mupdf.png'
        build_pdf_png(dx, out)
        wp = os.path.join(ROOT, 'pipeline_data', 'word_png', did, 'page_0001.png')
        if os.path.exists(wp):
            print('SSIM vs word_png:', round(rgb_ssim(wp, out), 4))
        print('saved', out)

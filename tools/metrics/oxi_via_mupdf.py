# -*- coding: utf-8 -*-
"""Render Oxi's layout via PyMuPDF (the SAME rasterizer word_png uses), so the
text AA matches Word by construction and the SSIM measures pure LAYOUT fidelity
(no DirectWrite-vs-MuPDF rasterizer confound).

word_png = Word ExportAsFixedFormat(PDF) -> fitz get_pixmap. This builds a PDF
from Oxi's --dump-layout JSON (per-char x, per-font, real baseline) and rasterizes
it the same way. S494 (2026-06-04): prototype confirmed +0.0198 on 1ec1 vs DirectWrite.

Usage: python tools/metrics/oxi_via_mupdf.py <docid> [--png OUT.png]
       python tools/metrics/oxi_via_mupdf.py --bench   (bottom-N + controls vs dwrite)
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

def build_pdf_png(layout, out_png, asc_cal=None):
    # asc_cal: dict fontname->ratio override; default uses the font ascender
    doc = fitz.open()
    for page in layout['pages']:
        PW, PH = page['width'], page['height']
        pg = doc.new_page(width=PW, height=PH)
        pg.insert_font(fontname='mincho', fontfile=MINCHO)
        pg.insert_font(fontname='gothic', fontfile=GOTHIC)
        for el in page['elements']:
            t = el.get('type')
            if t == 'text':
                txt = el.get('text', '')
                if not txt:
                    continue
                fs = el['font_size']
                fn, fobj = font_for(el.get('font_family'))
                asc = (asc_cal or {}).get(fn, fobj.ascender)
                baseline = el['y'] + el['text_y_off'] + fs * asc
                n = len(txt)
                w = el.get('w', 0.0)
                # per-char x: distribute the element width uniformly (CJK runs are uniform)
                adv = (w / n) if (n > 0 and w > 0) else fs
                x = el['x']
                for i, ch in enumerate(txt):
                    if ch == ' ':
                        continue
                    try:
                        pg.insert_text((x + i * adv, baseline), ch, fontname=fn, fontsize=fs, color=(0, 0, 0))
                    except Exception:
                        pass
            elif t == 'border':
                x, y, w, h = el['x'], el['y'], el['w'], el['h']
                pg.draw_line((x, y), (x + w, y + h), color=(0, 0, 0), width=0.75)
            elif t == 'shading':
                x, y, w, h = el['x'], el['y'], el['w'], el['h']
                # cell shading fill — skip (color not in dump); borders dominate
        pix = pg.get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72))
        if page is layout['pages'][0]:
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
        lay = dump_layout(dx)
        mp = build_pdf_png(lay, tempfile.mktemp(suffix='.png'))
        s_mp = rgb_ssim(wp, mp)
        s_dw = rgb_ssim(wp, render_dwrite(dx))
        print('%-16s %8.4f %8.4f %+8.4f' % (did[:16], s_dw, s_mp, s_mp - s_dw))

if __name__ == '__main__':
    if '--bench' in sys.argv:
        bench()
    else:
        did = sys.argv[1]
        dx = docx_for(did)
        lay = dump_layout(dx)
        out = 'c:/tmp/oxi_mupdf.png'
        build_pdf_png(lay, out)
        wp = os.path.join(ROOT, 'pipeline_data', 'word_png', did, 'page_0001.png')
        if os.path.exists(wp):
            print('SSIM vs word_png:', round(rgb_ssim(wp, out), 4))
        print('saved', out)

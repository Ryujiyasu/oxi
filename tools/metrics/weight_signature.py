# -*- coding: utf-8 -*-
"""Weight/AA signature per doc: render at the current dwrite default (ss2), compare gray stats
Oxi vs Word (position-insensitive globals: very-dark<64 ratio, mean of dark pixels, total ink).
Also parse the doc's dominant eastAsia font. Goal: is the Oxi-vs-Word weight DIRECTION per-FONT
(→ per-font gamma fixable) or per-doc arbitrary? cp932-safe (font names by codepoint-safe parse)."""
import json, os, glob, subprocess, tempfile, zipfile, collections
import xml.etree.ElementTree as ET
import numpy as np
from PIL import Image

DW = os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DOCS = ['0e7af1ae8f21_20230331_resources_open_data_contract_sample_00',
        'd77a58485f16_20240705_resources_data_outline_08',
        '683ffcab86e2_20230331_resources_open_data_contract_addon_00',
        'b837808d0555_20240705_resources_data_guideline_02',
        '15076df085f5_tokumei_08_09',
        '1636d28e2c46_tokumei_08_04',
        '29dc6e8943fe_order_01',
        'de6e32b5960b_tokumei_08_01-1']


def docx_for(d):
    for base in ('tools/golden-test/documents/docx', 'pipeline_data/docx'):
        f = glob.glob(os.path.join(base, d + '*.docx'))
        for p in f:
            if os.path.splitext(os.path.basename(p))[0] == d:
                return p
        if f:
            return f[0]
    return None


def dominant_ea_font(docx):
    try:
        root = ET.fromstring(zipfile.ZipFile(docx).read('word/document.xml'))
    except Exception:
        return '?'
    cnt = collections.Counter()
    for rf in root.iter(W + 'rFonts'):
        ea = rf.get(W + 'eastAsia')
        if ea:
            cnt[ea] += 1
    # also default from styles
    if not cnt:
        try:
            sroot = ET.fromstring(zipfile.ZipFile(docx).read('word/styles.xml'))
            for rf in sroot.iter(W + 'rFonts'):
                ea = rf.get(W + 'eastAsia')
                if ea:
                    cnt[ea] += 1
        except Exception:
            pass
    return cnt.most_common(1)[0][0] if cnt else '(theme)'


def weight_sig(docx, docid):
    wdir = f'pipeline_data/word_png/{docid}'
    wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
    if not wpages:
        return None
    with tempfile.TemporaryDirectory() as td:
        pref = os.path.join(td, 'o')
        subprocess.run([DW, os.path.abspath(docx), pref, '150'], capture_output=True, timeout=400)
        odark = []; wdark = []; ovd = []; wvd = []; inkr = []
        for wp in wpages:
            pg = int(os.path.basename(wp)[5:9])
            op = os.path.join(td, f'o_p{pg}.png')
            if not os.path.exists(op):
                continue
            a = np.asarray(Image.open(wp).convert('L')).astype(float)
            b = Image.open(op).convert('L')
            if b.size != (a.shape[1], a.shape[0]):
                b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
            b = np.asarray(b).astype(float)
            wdark.append(a[a < 200].mean() if (a < 200).any() else 255)
            odark.append(b[b < 200].mean() if (b < 200).any() else 255)
            wvd.append((a < 64).mean()); ovd.append((b < 64).mean())
            inkr.append((255 - b).sum() / max((255 - a).sum(), 1))
    n = len(odark)
    if n == 0:
        return None
    return {'n': n,
            'oxi_dark': sum(odark) / n, 'word_dark': sum(wdark) / n,
            'dark_diff': sum(odark) / n - sum(wdark) / n,  # <0 = Oxi darker (harder)
            'vd_ratio': (sum(ovd) / n) / max(sum(wvd) / n, 1e-9),  # >1 = Oxi more very-dark
            'ink_ratio': sum(inkr) / n}


print("%-30s %-14s %5s %8s %8s %8s %8s" % ('doc', 'ea_font', 'n', 'oxiDk', 'wordDk', 'darkDiff', 'vdRatio'))
print('-' * 95)
for d in DOCS:
    dx = docx_for(d)
    if not dx:
        print("%-30s (no docx)" % d[:30]); continue
    font = dominant_ea_font(dx)
    sig = weight_sig(dx, d)
    if not sig:
        print("%-30s %-14s (no png)" % (d[:30], font[:14])); continue
    print("%-30s %-14s %5d %8.1f %8.1f %+8.1f %8.3f" % (
        d[:30], font[:14], sig['n'], sig['oxi_dark'], sig['word_dark'], sig['dark_diff'], sig['vd_ratio']))
print('-' * 95)
print("darkDiff <0 => Oxi cores DARKER/harder (supersample helps); >0 => Oxi LIGHTER (supersample hurts)")
print("vdRatio >1 => Oxi more very-dark pixels")

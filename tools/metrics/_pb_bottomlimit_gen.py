# -*- coding: utf-8 -*-
"""Word's page-bottom limit as a FUNCTION (size / line rule / grid pitch).

MEASURED (content bottom 770.00, A4, last accepted BASELINE):
    MS Mincho  9pt   765.60  reserve 4.400 = 0.4889 em
    MS Mincho 10.5   764.88          5.120 = 0.4876 em
    MS Mincho 12     764.16          5.840 = 0.4867 em
    MS Mincho 14     763.08          6.920 = 0.4943 em
    MS Gothic 10.5   764.88          5.120 = 0.4876 em
    Times NR  10.5   767.04          2.960 = 0.2819 em   <- face-dependent
    exact 13pt       767.52          2.480 = 0.191 of the line height
    atLeast 13pt     764.88          5.120  (same as auto)
    pitch 300/420/absent: identical -- the GRID DOES NOT ENTER.
    space_before shifts the flip by exactly its own value; the limit is unchanged.

CAVEAT: the Oxi column of this probe is NOT Oxi's page-bottom rule.  The spacer
is a single ~680pt exact line, and Oxi's flip is governed by how it handles that
giant line, not by the test line's bottom test.  Use it for the WORD side only.

Same construction as bottomlimit.py -- an exact-height spacer followed by one
test line -- but the flip is found by binary search (~14 renders per arm) so a
whole matrix is affordable.  Reports Word's last accepted BASELINE and the
reserve below it, plus Oxi's flip for the same arm.
"""
import os, sys, zipfile, subprocess, json, io
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, 'C:/Users/ryuji/oxi-main/tools/metrics')
import _pb_gridbottom_gen as G

OUT = os.path.dirname(os.path.abspath(__file__))
EXE = 'C:/Users/ryuji/oxi-main/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'
TEXT = '測定行'

# (tag, font, half-pt size, line rule, line value, grid pitch, space_before_xml)
ARMS = [
    ('mincho9',    'ＭＳ明朝', 18, 'auto', 240, 360),
    ('mincho105',  'ＭＳ明朝', 21, 'auto', 240, 360),
    ('mincho12',   'ＭＳ明朝', 24, 'auto', 240, 360),
    ('mincho14',   'ＭＳ明朝', 28, 'auto', 240, 360),
    ('gothic105',  'ＭＳゴシック', 21, 'auto', 240, 360),
    ('times105',   'Times New Roman', 21, 'auto', 240, 360),
    ('m105_exact', 'ＭＳ明朝', 21, 'exact', 260, 360),
    ('m105_atleast', 'ＭＳ明朝', 21, 'atLeast', 260, 360),
    ('m105_p300',  'ＭＳ明朝', 21, 'auto', 240, 300),
    ('m105_p420',  'ＭＳ明朝', 21, 'auto', 240, 420),
    ('m105_nogrid', 'ＭＳ明朝', 21, 'auto', 240, None),
    ('m105_before120', 'ＭＳ明朝', 21, 'auto', 240, 360, 'w:before="120"'),
    ('m105_bl50', 'ＭＳ明朝', 21, 'auto', 240, 360, 'w:beforeLines="50" w:before="120"'),
]


def build(arm, spacer_tw):
    tag, font, sz, rule, lval, pitch = arm[:6]
    sb = arm[6] if len(arm) > 6 else 'w:before="0"'
    rpr = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/>'
    spacer = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
              f'w:line="{spacer_tw}" w:lineRule="exact"/>'
              f'<w:rPr>{rpr}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{rpr}</w:rPr><w:t>詰</w:t></w:r></w:p>')
    test = ('<w:p><w:pPr>'
            f'<w:spacing {sb} w:after="0" w:line="{lval}" w:lineRule="{rule}"/>'
            f'<w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{TEXT}</w:t></w:r></w:p>')
    grid = f'<w:docGrid w:linePitch="{pitch}"/>' if pitch else ''
    sect = (f'<w:sectPr><w:pgSz w:w="{G.PAGE_W}" w:h="{G.PAGE_H}" w:code="9"/>'
            f'<w:pgMar w:top="1440" w:right="{G.SIDE}" w:bottom="{G.BOTTOM}" '
            f'w:left="{G.SIDE}" w:header="720" w:footer="720" w:gutter="0"/>'
            f'{grid}</w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {G.W_NS}><w:body>{spacer}{test}{sect}</w:body></w:document>')
    path = os.path.join(OUT, f'bl2_{tag}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', G.CT)
        z.writestr('_rels/.rels', G.RELS)
        z.writestr('word/_rels/document.xml.rels', G.DOC_RELS)
        z.writestr('word/settings.xml', G.SETTINGS)
        z.writestr('word/styles.xml', G.styles(font, sz))
        z.writestr('word/document.xml', doc)
    return path


def word_page(w, path):
    import pymupdf
    pdf = path[:-5] + '.pdf'
    if os.path.exists(pdf):
        os.remove(pdf)
    d = w.w.Documents.Open(os.path.abspath(path), ReadOnly=True)
    d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
    d.Close(False)
    doc = pymupdf.open(pdf)
    for pi in range(len(doc)):
        for b in doc[pi].get_text('dict')['blocks']:
            if b.get('type') != 0:
                continue
            for l in b.get('lines', []):
                if TEXT in ''.join(s['text'] for s in l['spans']):
                    return pi + 1, max(s['origin'][1] for s in l['spans'])
    return None, None


def oxi_page(path):
    subprocess.run([EXE, path, 'bl2', '110', '--dump-layout=bl2.json'],
                   capture_output=True, cwd=OUT)
    d = json.load(io.open(os.path.join(OUT, 'bl2.json'), encoding='utf-8'))
    for pi, pg in enumerate(d['pages'], 1):
        for e in pg['elements']:
            if e.get('text') and TEXT[0] in e['text']:
                return pi, e['y'], e['y'] + e.get('text_y_off', 0)
    return None, None, None


def flip(probe, arm, lo=9000, hi=14400):
    """largest spacer with the test line still on page 1."""
    if probe(build(arm, lo))[0] != 1:
        return None
    if probe(build(arm, hi))[0] == 1:
        return None
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if probe(build(arm, mid))[0] == 1:
            lo = mid
        else:
            hi = mid
    return lo


if __name__ == '__main__':
    only = sys.argv[1:] or None
    w = G.Word()
    try:
        print('%-14s %5s %-8s %5s | %9s %8s %8s | %9s %9s'
              % ('arm', 'fs', 'rule', 'pitch', 'W_spacer', 'W_basel', 'reserve', 'em', 'O_spacer'))
        for arm in ARMS:
            if only and arm[0] not in only:
                continue
            fs = arm[2] / 2.0
            ws = flip(lambda p: word_page(w, p), arm)
            if ws is None:
                print('%-14s  no flip in range' % arm[0])
                continue
            _, bl = word_page(w, build(arm, ws))
            os_ = flip(lambda p: oxi_page(p), arm)
            res = G.CONTENT_BOTTOM - bl
            print('%-14s %5.1f %-8s %5s | %9d %8.2f %8.3f | %9.4f %9s'
                  % (arm[0], fs, arm[3], arm[5], ws, bl, res, res / fs,
                     os_ if os_ is not None else '-'))
    finally:
        w.close()

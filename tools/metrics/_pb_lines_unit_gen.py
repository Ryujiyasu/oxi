"""What is one "line" for w:beforeLines / w:afterLines?

MEASURED 2026-09-19 (three-paragraph probe, MS Mincho, beforeLines=100,
advance minus the natural line height = the unit):

    docGrid type="linesAndChars" pitch 403   unit = 20.15 = the pitch (0.3 -> 6.04)
    docGrid WITHOUT type, pitch 403          unit = 11.90
    ... pitch 286 / 355 / 360                unit = 11.90
    ... compatibilityMode 11 / 14 / 15 / no settings.xml   unit = 11.90
    ... docDefaults 10 / 12 / 14pt, run 10 / 14pt          unit = 12.05..12.16
    no docGrid                               unit = 12.0  (S697)

So a docGrid without w:type is not a line grid for spacing: the unit is the
fixed 12pt no-grid value.  Oxi used the pitch for every non-360 no-type grid
(S571 keeps the pitch for LINE snapping); S1480 switches the spacing unit.

The same probe also shows Word does NOT snap lines to a no-type grid
(advance 13.68 = natural for 10.5pt) under any compat mode -- that conflicts
with S571's ikujidetail render truth and is left for its own derivation.

The font name MUST be 'ＭＳ 明朝' (full-width space); without it Word silently
substitutes Yu Gothic and every number moves by ~30%.

Usage:  python _pb_lines_unit_gen.py [arm ...]
"""
import os
import sys
import zipfile
import subprocess
import json
import io

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_gridbottom_gen as G  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..',
                   'pipeline_data', '_pb_lines_unit')
EXE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..',
                   'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
FONT = 'ＭＳ 明朝'

# (tag, compat or None, docGrid xml, beforeLines)
ARMS = [
    ('lc403', 15, '<w:docGrid w:type="linesAndChars" w:linePitch="403" w:charSpace="-2156"/>', 100),
    ('c15_p403', 15, '<w:docGrid w:linePitch="403"/>', 100),
    ('c14_p403', 14, '<w:docGrid w:linePitch="403"/>', 100),
    ('c11_p403', 11, '<w:docGrid w:linePitch="403"/>', 100),
    ('nosettings_p403', None, '<w:docGrid w:linePitch="403"/>', 100),
    ('c11_p286', 11, '<w:docGrid w:linePitch="286"/>', 100),
    ('c15_p355', 15, '<w:docGrid w:linePitch="355"/>', 100),
    ('c14_p360', 14, '<w:docGrid w:linePitch="360"/>', 100),
    ('nogrid', 15, '', 100),
]


def settings(cm):
    if cm is None:
        return None
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {G.W_NS}><w:compat><w:compatSetting w:name="compatibilityMode" '
            f'w:uri="http://schemas.microsoft.com/office/word" w:val="{cm}"/></w:compat></w:settings>')


def build(tag, cm, grid, bl):
    rpr = f'<w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="21"/>'

    def p(t, extra='w:before="0"'):
        return (f'<w:p><w:pPr><w:spacing {extra} w:after="0" w:line="240" w:lineRule="auto"/>'
                f'<w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>{t}</w:t></w:r></w:p>')

    body = p('一行目') + p('二行目') + p('三行目', f'w:beforeLines="{bl}"') + p('四行目', f'w:beforeLines="{bl}"')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
            '<w:pgMar w:top="697" w:right="697" w:bottom="697" w:left="697" '
            f'w:header="45" w:footer="142" w:gutter="0"/>{grid}</w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {G.W_NS}><w:body>{body}{sect}</w:body></w:document>')
    st = settings(cm)
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, f'lu_{tag}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        ct = G.CT if st else G.CT.replace(
            '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>', '')
        rels = G.DOC_RELS if st else G.DOC_RELS.replace(
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>', '')
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', G.RELS)
        z.writestr('word/_rels/document.xml.rels', rels)
        if st:
            z.writestr('word/settings.xml', st)
        z.writestr('word/styles.xml', G.styles(FONT, 24))
        z.writestr('word/document.xml', doc)
    return path


def main():
    import pymupdf
    only = sys.argv[1:] or None
    w = G.Word()
    try:
        print('%-16s %9s %9s %9s | %9s' % ('arm', 'adv(line)', 'adv(bl)', 'W_unit', 'O_unit'))
        for arm in ARMS:
            if only and arm[0] not in only:
                continue
            p = build(*arm)
            pdf = p[:-5] + '.pdf'
            if os.path.exists(pdf):
                os.remove(pdf)
            d = w.w.Documents.Open(os.path.abspath(p), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
            d.Close(False)
            bl = sorted(max(s['origin'][1] for s in l['spans'])
                        for b in pymupdf.open(pdf)[0].get_text('dict')['blocks'] if b.get('type') == 0
                        for l in b['lines'] if ''.join(s['text'] for s in l['spans']).strip())
            line, adv = bl[1] - bl[0], bl[2] - bl[1]
            subprocess.run([EXE, p, 'lu', '110', '--dump-layout=lu.json'], capture_output=True, cwd=OUT)
            dd = json.load(io.open(os.path.join(OUT, 'lu.json'), encoding='utf-8'))
            ys = sorted({round(e['y'], 2) for e in dd['pages'][0]['elements'] if e.get('text') and e['text'].strip()})
            o_unit = (ys[2] - ys[1]) - (ys[1] - ys[0]) if len(ys) > 2 else float('nan')
            print('%-16s %9.2f %9.2f %9.2f | %9.2f' % (arm[0], line, adv, adv - line, o_unit))
    finally:
        w.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

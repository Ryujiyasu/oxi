# -*- coding: utf-8 -*-
"""S506 oidashi discriminator repro: confirm whether docGrid PRESENCE decides hang
(burasagari) vs oidashi for a jc=left line whose trailing 。 lands at the wrap boundary.
Build variants (no-docGrid / docGrid type=lines / docGrid linesAndChars), COM->PDF->fitz,
and report for each: does 。 stay on line 1 (HANG/overflow) or wrap to line 2 (OIDASHI)?
Expected per the S505 discriminator: no-docGrid HANG, docGrid present OIDASHI. cp932-safe."""
import os, sys, zipfile, subprocess, tempfile, io, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'oidashi')
os.makedirs(OUT, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
# 37 kanji + period: at width ~450pt / 12pt, 37 fit (444) and 。 (38th -> 456) overflows.
TEXT = '亜' * 37 + '。'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
WRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings %s><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>' % (NS, compat))
RPR = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'


def doc_xml(grid):
    para = '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (RPR, TEXT)
    if grid == 'none':
        dg = ''
    elif grid == 'lines':
        dg = '<w:docGrid w:type="lines" w:linePitch="360"/>'
    else:
        dg = '<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="0"/>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1450" w:bottom="1440" w:left="1450" w:header="720" w:footer="720"/>'
            '%s</w:sectPr>' % dg)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, para, sect)


def build(name, grid, compat):
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', WRELS); z.writestr('word/settings.xml', settings(compat))
        z.writestr('word/document.xml', doc_xml(grid))
    return p


def measure_word(dx):
    """Return (#chars on line 1, whether 。 is on line1). PDF->fitz, cluster by baseline y."""
    out = tempfile.mktemp(suffix='.json', dir='c:/tmp')
    subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'metrics', 'word_pdf_glyphs.py'), dx, out],
                   capture_output=True, timeout=120)
    g = json.load(io.open(out, encoding='utf-8'))['pages'][0]['glyphs']
    g = [x for x in g if x['char'].strip()]
    if not g:
        return (0, None, 'no glyphs')
    rows = {}
    for x in g:
        rows.setdefault(round(x['y'], 0), []).append(x)
    ys = sorted(rows)
    line1 = sorted(rows[ys[0]], key=lambda x: x['x'])
    n1 = len(line1)
    period_on_l1 = any(x['char'] == '。' for x in line1)
    # right edge of line 1 vs text right boundary (page 595.3 - 72.5 margin = 522.8pt)
    right_edge = max(x['x'] for x in line1)
    return (n1, period_on_l1, 'line1_right_x=%.1f' % right_edge)


def main():
    L = ['S506 oidashi discriminator: jc=left, 37x 亜 + 。, width ~450pt (37 fit, 。 overflows)',
         'text right boundary ~= 595.3 - 72.5 = 522.8pt; expect HANG=。on line1(overflow), OIDASHI=。on line2',
         '%-22s %8s %10s %s' % ('variant', 'L1_chars', '。on_L1?', 'note')]
    variants = [('od_c15_nogrid.docx', 'none', 15), ('od_c12_nogrid.docx', 'none', 12),
                ('od_c15_lines.docx', 'lines', 15), ('od_c12_lines.docx', 'lines', 12),
                ('od_c14_nogrid.docx', 'none', 14)]
    for name, grid, compat in variants:
        dx = build(name, grid, compat)
        n1, p1, note = measure_word(dx)
        verdict = 'HANG' if p1 else ('OIDASHI' if p1 is False else '?')
        L.append('%-22s %8d %10s %s  -> %s' % (name, n1, str(p1), note, verdict))
    with io.open('c:/tmp/_s506_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('wrote c:/tmp/_s506_out.txt')


if __name__ == '__main__':
    main()

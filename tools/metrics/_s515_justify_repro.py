# -*- coding: utf-8 -*-
"""S515 justify distribution derivation: a justified (jc=both) line of identical CJK chars,
measure Word's per-char advances to derive HOW Word distributes the slack (uniform? discrete
units? which gaps get the extra). Compare to Oxi. No docGrid (pure 83/64). cp932-safe."""
import os, zipfile, subprocess, io, json, statistics
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'justify')
os.makedirs(OUT, exist_ok=True)
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CJK = 'あ' * 90  # wraps to multiple justified lines
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
RPR = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'


def doc_xml():
    para = '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (RPR, CJK)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/></w:sectPr>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, para, sect)


p = os.path.join(OUT, 'jb_kanji.docx')
with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
    z.writestr('word/document.xml', doc_xml())


def word_glyphs(dx):
    subprocess.run(['python', os.path.join(ROOT, 'tools', 'metrics', 'word_pdf_glyphs.py'), dx, 'c:/tmp/jb_w.json'], capture_output=True, timeout=120)
    return json.load(io.open('c:/tmp/jb_w.json', encoding='utf-8'))['pages'][0]['glyphs']


def oxi_glyphs(dx):
    subprocess.run([DW, dx, 'c:/tmp/jbox', '150', '--dump-glyphs=c:/tmp/jb_ox.json'], capture_output=True, timeout=120)
    return json.load(io.open('c:/tmp/jb_ox.json', encoding='utf-8'))['pages'][0]['glyphs']


def line1_adv(glyphs, ykey):
    g = [x for x in glyphs if x['char'].strip()]
    rows = {}
    for x in g:
        rows.setdefault(round(x[ykey], 0), []).append(x)
    y0 = sorted(rows)[0]
    r = sorted(rows[y0], key=lambda x: x['x'])
    return [round(r[i + 1]['x'] - r[i]['x'], 3) for i in range(len(r) - 1)], len(r)


wd = word_glyphs(os.path.abspath(p))
ox = oxi_glyphs(os.path.abspath(p))
wa, wn = line1_adv(wd, 'y')
oa, on = line1_adv(ox, 'baseline')
L = ['S515 justify distribution (jc=both, %d kanji, MS Mincho 10.5pt, no docGrid)' % len(CJK)]
L.append('Word line1 n=%d  advances:' % wn)
L.append('  ' + ' '.join('%.2f' % a for a in wa))
L.append('  set=%s  mean=%.3f' % (sorted(set(round(a, 2) for a in wa)), statistics.mean(wa) if wa else 0))
L.append('Oxi  line1 n=%d  advances:' % on)
L.append('  ' + ' '.join('%.2f' % a for a in oa))
L.append('  set=%s  mean=%.3f' % (sorted(set(round(a, 2) for a in oa)), statistics.mean(oa) if oa else 0))
with io.open('c:/tmp/_s515_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('\n'.join(L))

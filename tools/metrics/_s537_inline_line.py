# -*- coding: utf-8 -*-
"""S537: pin Word's INLINE-DRAWING-IN-LINE spec (the flow model for wp:inline objects).

Builds a docx with paragraphs hosting an inline picture (the 3a4f calendar EMF scaled
via v:shape style) under varying conditions, then COM-measures each paragraph's top Y
(Information(6) on collapsed start) so the HOST LINE height = next_top - host_top.

Paragraphs (MS Mincho 10.5pt body, docGrid none):
  P0  marker text (baseline)
  P1  image-only para, default spacing            -> line = extent? extent+leading?
  P2  marker
  P3  image-only para, spacing line=350 atLeast   -> 3a4f pict-para config
  P4  marker
  P5  text BEFORE + image (same run para)         -> line = max(text, extent)?
  P6  marker
  P7  image-only para, spacing line=350 EXACT     -> clipped to 17.5? or extent?
  P8  marker
Image extent: 120pt tall (small enough to fit pages, big vs 17.5 line).
Also renders Oxi (dump-layout) for the same paragraphs. ASCII-out, cp932-safe."""
import os, sys, io, zipfile, subprocess, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join('c:/tmp', 's537_inline.docx')
GDI = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
EMF = r'c:\tmp\3a4f_extract\word\media\image2.emf'

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="emf" ContentType="image/x-emf"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId17" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.emf"/>
</Relationships>'''

PICT = ('<w:r><w:pict><v:shape type="#_x0000_t75" style="width:200pt;height:120pt">'
        '<v:imagedata r:id="rId17"/></v:shape></w:pict></w:r>')


def para(content, ppr=''):
    return '<w:p>%s%s</w:p>' % (('<w:pPr>%s</w:pPr>' % ppr) if ppr else '', content)


def marker(t):
    return para('<w:r><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="ＭＳ 明朝"/></w:rPr><w:t>%s</w:t></w:r>' % t)


def build():
    body = (
        marker('M0') +
        para(PICT) +                                                     # P1 default
        marker('M1') +
        para(PICT, '<w:spacing w:line="350" w:lineRule="atLeast"/>') +   # P3 atLeast350
        marker('M2') +
        para('<w:r><w:t>text</w:t></w:r>' + PICT) +                      # P5 text+image
        marker('M3') +
        para(PICT, '<w:spacing w:line="350" w:lineRule="exact"/>') +     # P7 exact350
        marker('M4') +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
           ' xmlns:v="urn:schemas-microsoft-com:vml"'
           ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<w:body>%s</w:body></w:document>' % body)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.write(EMF, 'word/media/image2.emf')
    print('built', OUT)


WORD = r'''
import sys
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch('Word.Application'); word.Visible=False; word.DisplayAlerts=False
word.AutomationSecurity = 3
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
try:
    n = doc.Paragraphs.Count
    for i in range(1, n+1):
        rng = doc.Paragraphs(i).Range
        s = doc.Range(rng.Start, rng.Start)
        y = s.Information(6)
        t = rng.Text.strip()[:10]
        print('WPARA %d y=%.2f text=%r' % (i, y, t))
finally:
    doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
'''


def main():
    build()
    r = subprocess.run([sys.executable, '-c', WORD, OUT], capture_output=True, text=True,
                       encoding='utf-8', errors='replace', timeout=90)
    wlines = [l for l in (r.stdout or '').splitlines() if l.startswith('WPARA')]
    ys = []
    for l in wlines:
        parts = l.split()
        ys.append(float(parts[2].split('=')[1]))
    print('WORD paragraph tops:', ['%.1f' % y for y in ys])
    labels = ['M0', 'P1 img default', 'M1', 'P3 img atLeast350', 'M2', 'P5 text+img', 'M3', 'P7 img exact350', 'M4']
    for i in range(len(ys) - 1):
        lh = ys[i+1] - ys[i]
        lab = labels[i] if i < len(labels) else '?'
        print('WORD %-18s height=%7.2f' % (lab, lh))
    # Oxi
    subprocess.run([GDI, OUT, 'c:/tmp/s537', '150', '--dump-layout=c:/tmp/s537.json'], capture_output=True)
    d = json.load(io.open('c:/tmp/s537.json', encoding='utf-8'))
    els = d['pages'][0]['elements']
    items = []
    for e in els:
        t = str(e.get('type', ''))
        if t == 'text' and (e.get('text') or '').strip():
            items.append((e.get('y', 0), 'text:' + e.get('text', '')[:8]))
        elif t == 'image':
            items.append((e.get('y', 0), 'image h=%.0f' % e.get('h', 0)))
    items.sort()
    print('OXI elements:')
    for y, lab in items:
        print('  y=%7.2f %s' % (y, lab))


if __name__ == '__main__':
    main()

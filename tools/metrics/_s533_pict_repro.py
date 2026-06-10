# -*- coding: utf-8 -*-
"""S533 repro: VML w:pict v:shape(t75)+v:imagedata EMF image.
P1: pict in a BODY paragraph. P2: pict inside a single-cell TABLE.
Checks Oxi layout (--dump-layout) for image elements and their sizes.
Reuses 3a4f's calendar EMF (c:/tmp/3a4f_extract/word/media/image2.emf)."""
import os, sys, io, zipfile, subprocess, json, shutil
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join('c:/tmp', 's533_pict.docx')
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

PICT = ('<w:r><w:pict><v:shape id="_x0000_i1026" type="#_x0000_t75" style="width:424.5pt;height:321.75pt">'
        '<v:imagedata r:id="rId17" o:title=""/></v:shape></w:pict></w:r>')

DOC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
       ' xmlns:v="urn:schemas-microsoft-com:vml"'
       ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
       ' xmlns:o="urn:schemas-microsoft-com:office:office">'
       '<w:body>'
       '<w:p><w:r><w:t>BODY pict below</w:t></w:r></w:p>'
       '<w:p>' + PICT + '</w:p>'
       '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>'
       '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
       '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
       '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
       '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr>'
       '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
       '<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
       '<w:p><w:r><w:t>CELL pict below</w:t></w:r></w:p>'
       '<w:p>' + PICT + '</w:p>'
       '</w:tc></w:tr></w:tbl>'
       '<w:p/>'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
       '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
       '</w:body></w:document>')


def main():
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', DOC)
        z.write(EMF, 'word/media/image2.emf')
    print('built', OUT)
    env = dict(os.environ)
    env['OXI_S331_CELL_INLINE_IMG'] = '1'
    r = subprocess.run([GDI, OUT, 'c:/tmp/s533_pict', '150', '--dump-layout=c:/tmp/s533_layout.json'],
                       capture_output=True, text=True, env=env)
    print((r.stderr or '')[-200:])
    d = json.load(io.open('c:/tmp/s533_layout.json', encoding='utf-8'))
    for pi, p in enumerate(d['pages']):
        for e in p.get('elements', []):
            t = str(e.get('type', ''))
            if t == 'image' or 'mage' in str(e.get('content', '')):
                print('page %d IMAGE xywh=(%.1f, %.1f, %.1f, %.1f)' % (pi+1, e.get('x', 0), e.get('y', 0), e.get('width', 0), e.get('height', 0)))
    # text element positions to see flow
    for pi, p in enumerate(d['pages']):
        for e in p.get('elements', []):
            if str(e.get('type','')) == 'text' and e.get('text','').strip():
                print('page %d text y=%.1f %r' % (pi+1, e.get('y',0), e.get('text','')[:20]))


if __name__ == '__main__':
    main()

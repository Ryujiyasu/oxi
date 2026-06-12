# -*- coding: utf-8 -*-
"""S553 — does style-level-only w:kern (d77a Normal pattern) reach Oxi's
pair-halving? Build 国、（国国 with kern ONLY in the Normal style rPr,
render with oxi-gdi --dump-layout, print widths. Expected if resolved:
、 = fs/2 (pair-halved). fs=12 (sz24) to match d77a."""
import io
import json
import os
import subprocess
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
GDI = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
DOCX = r'c:\tmp\s553_kerntest.docx'

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
          '<w:rPr><w:kern w:val="2"/></w:rPr></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/></w:settings>')
BODY = ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
        '<w:t xml:space="preserve">国、（国国</w:t></w:r></w:p>')
SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/></w:sectPr>')
DOC = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:body>%s%s</w:body></w:document>') % (BODY, SECT)

with zipfile.ZipFile(DOCX, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT)
    z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', DRELS)
    z.writestr('word/document.xml', DOC)
    z.writestr('word/styles.xml', STYLES)
    z.writestr('word/settings.xml', SETTINGS)

subprocess.run([GDI, DOCX, r'c:\tmp\s553kt', '--dump-layout=c:/tmp/s553_kt.json'],
               capture_output=True)
d = json.load(io.open('c:/tmp/s553_kt.json', encoding='utf-8'))
for e in d['pages'][0].get('elements', []):
    if e.get('type') == 'text' and e.get('text'):
        print('w=%6.3f |%s|' % (e['w'], e['text']))

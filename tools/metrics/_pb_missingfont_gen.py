# -*- coding: utf-8 -*-
"""Which installed font does Word substitute for a font that is not installed and has no
fontTable altName?

policies__00602e8a: runs in "NTPreCursivef" (fontTable panose 0300..., family script, no
altName, not installed) render in Calibri 14 in Word's PDF (= the theme minor / docDefaults
font); Oxi's fallback is ~5% wider and wraps a one-line paragraph.

Sheet: one paragraph of Latin text in the fake font "NoSuchFontQQ" 14pt. Arms: docDefaults
ascii font (Calibri / Times New Roman / Arial, direct rFonts, no theme) × fontTable entry for
the fake font (none / panose script family "script" / panose roman family "roman" /
family "swiss" pitch fixed). Readout: the PDF span font (ExportAsFixedFormat + pymupdf) and
the line's right edge.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/missingfont'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
FAKE = os.environ.get('FAKE', 'NoSuchFontQQ')
THEME_CT = '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
THEME_REL = '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'


def theme_xml(major, minor):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>'
            '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
            f'<a:fontScheme name="Office"><a:majorFont><a:latin typeface="{major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="{minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>')

TEXT = 'Reason: To support the children at lunch time and make sure that they are ready to learn in the afternoon lessons.'


def styles(default_font):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{default_font}" w:hAnsi="{default_font}" w:eastAsia="{default_font}" w:cs="{default_font}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')


def font_table(entry):
    e = ''
    if entry == 'script':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="03000400000000000000"/><w:charset w:val="00"/><w:family w:val="script"/><w:pitch w:val="variable"/></w:font>'
    elif entry == 'roman':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
    elif entry == 'sigscript':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="03000400000000000000"/><w:charset w:val="00"/><w:family w:val="script"/><w:pitch w:val="variable"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigcalibri':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="E00002FF" w:usb1="4000ACFF" w:usb2="00000001" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font>'
    elif entry == 'sigroman':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigmodern':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="020B0609020204030204"/><w:charset w:val="00"/><w:family w:val="modern"/><w:pitch w:val="fixed"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigdecor':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="04020505020E03020304"/><w:charset w:val="00"/><w:family w:val="decorative"/><w:pitch w:val="variable"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigauto':
        e = f'<w:font w:name="{FAKE}"><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="default"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigswiss':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigswissfixed':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="020B0609020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="fixed"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigscriptnop':
        e = f'<w:font w:name="{FAKE}"><w:charset w:val="00"/><w:family w:val="script"/><w:pitch w:val="variable"/><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'sigonly':
        e = f'<w:font w:name="{FAKE}"><w:sig w:usb0="00000003" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
    elif entry == 'swissfixed':
        e = f'<w:font w:name="{FAKE}"><w:panose1 w:val="020B0609020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="fixed"/></w:font>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + e + '</w:fonts>')


def document():
    rpr = f'<w:rPr><w:rFonts w:ascii="{FAKE}" w:hAnsi="{FAKE}"/><w:sz w:val="28"/></w:rPr>'
    body = (f'<w:p><w:r><w:t>control line in the default font</w:t></w:r></w:p>'
            f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>')
    sect = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client, pymupdf
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    ARMS = [('Calibri', 'sigswiss', None), ('Calibri', 'sigswissfixed', None), ('Calibri', 'sigscriptnop', None), ('Calibri', 'sigonly', None)] if os.environ.get('SIG_ARMS3') else [('Calibri', 'sigroman', None), ('Calibri', 'sigmodern', None), ('Calibri', 'sigdecor', None), ('Calibri', 'sigauto', None)] if os.environ.get('SIG_ARMS2') else [('Calibri', 'sigscript', None), ('Calibri', 'sigcalibri', None), ('Times New Roman', 'sigscript', None)] if os.environ.get('SIG_ARMS') else [(df, e, None) for df in ('Calibri', 'Times New Roman', 'Arial') for e in ('none', 'script', 'roman', 'swissfixed')] if os.environ.get('THEME_ARMS') is None else [
        ('Calibri', 'none', ('Cambria', 'Calibri')), ('Calibri', 'none', ('Calibri Light', 'Calibri')), ('Calibri', 'none', ('Arial', 'Arial')),
        ('Calibri', 'none', ('Times New Roman', 'Times New Roman')), ('Arial', 'none', ('Georgia', 'Verdana')), ('Arial', 'script', ('Georgia', 'Verdana'))]
    for default_font, entry, theme in ARMS:
            tag = default_font.replace(' ', '') + '_' + entry + ('' if theme is None else '_th_' + theme[0].replace(' ', '') + '-' + theme[1].replace(' ', ''))
            at = OUT / f'{tag}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                ct = CT if theme is None else CT.replace('</Types>', THEME_CT + '</Types>')
                dr = DR if theme is None else DR.replace('</Relationships>', THEME_REL + '</Relationships>')
                z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', dr)
                if theme is not None: z.writestr('word/theme/theme1.xml', theme_xml(*theme))
                z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', styles(default_font))
                z.writestr('word/fontTable.xml', font_table(entry)); z.writestr('word/document.xml', document())
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                pdf = str((OUT / f'{tag}.pdf').resolve())
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            doc = pymupdf.open(pdf); p = doc[0]
            rows = []
            for b in p.get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    t = ''.join(s['text'] for s in l['spans'])
                    rows.append((round(l['bbox'][1], 1), round(l['bbox'][2], 1), l['spans'][0]['font'], round(l['spans'][0]['size'], 1), t[:24]))
            doc.close()
            print(f'{default_font:16} {entry:10} theme={theme} ' + ' | '.join(f'y{y} x2={x2} {f} {s} «{t}»' for y, x2, f, s, t in sorted(rows)), flush=True)
finally:
    app.Quit()

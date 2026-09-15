# -*- coding: utf-8 -*-
"""How far may a line-end 。 hang past the boundary under a charSpace grid, compat 15?

policies__07543a6b9776a1cf p7, paragraph «このような用語の明確化は…こともある。»: Word
sets it in 4 lines, the 4th holding only «る。»; Oxi fits «る。» on line 3 with the
。 ending at 502.7 (7.7pt past the 495.0 its other lines end at, 2.8pt short of
the 505.45 margin). Word's line 3 ends at «あ» (483). So Word refuses a hang
that Oxi allows -- the one line that starts the document's 14pt drift.

Sheet = section 3 of the doc: A4, margins 1797/1797 (text 415.65pt), docGrid
linePitch 360 charSpace 6144 (no type), docDefaults TNR 12 (sz 24), Normal
minorHAnsi sz 21 kern 2 jc both, compat 15. Paragraph: leftChars 93 /
left 195, Arial + ＭＳ 明朝, sz inherited.

Arms: right indent swept 0..240tw in 20tw steps (13 arms) -- each shifts the
boundary 1pt left, so the hang «る。» needs grows 1pt per arm. Readout: Word
line count + line-3 last char; Oxi line count + line-3 end x.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/hangfloor'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsiaTheme="minorEastAsia" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/><w:pgMar w:top="1418" w:right="1797" w:bottom="1276" w:left="1797" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360" w:charSpace="6144"/></w:sectPr>')
TEXT = 'このような用語の明確化は情報収集時に依頼すべきである。もし、明確化が得られない場合、報告された情報に対する適切な質問がされたことを明らかにするため、「不明（“unknown”）」や「詳細不明（“unspecified”）」のような用語追加が有用なこともある。'
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="ＭＳ 明朝" w:hAnsi="Arial" w:cs="Arial"/><w:szCs w:val="21"/></w:rPr>'


import re as _re
REAL = _re.sub(r' (?:w14:\w+|w:rsid\w*)="[^"]*"', '', Path(r'C:/Users/ryuji/AppData/Local/Temp/p143.xml').read_text(encoding='utf-8'))


def para(rind):
    if os.environ.get('HANG_REAL'):
        ind = f'<w:ind w:leftChars="93" w:left="195" w:right="{rind}"/>' if rind else '<w:ind w:leftChars="93" w:left="195"/>'
        return REAL.replace('<w:ind w:leftChars="93" w:left="195"/>', ind)
    ind = f'<w:ind w:leftChars="93" w:left="195" w:right="{rind}"/>' if rind else '<w:ind w:leftChars="93" w:left="195"/>'
    return f'<w:p><w:pPr>{ind}{RPR}</w:pPr><w:r>{RPR}<w:t>{TEXT}</w:t></w:r></w:p>'


def document(rind):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p>{para(rind)}<w:p><w:r><w:t>B</w:t></w:r></w:p>{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
try:
    for rind in ([0, 20, 40] if os.environ.get('HANG_REAL') else range(0, 260, 20)):
        name = f'rind{rind:03d}'
        at = OUT / f"{name}.docx"
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", document(rind))
        d = app.Documents.Open(str(at.resolve()), False, True)
        try:
            p = d.Paragraphs(2); s, e = p.Range.Start, p.Range.End
            lines = []; last = None; cur = ''
            for k in range(s, e - 1):
                y = round(d.Range(k, k).Information(6), 2)
                if last is None or y != last:
                    if cur: lines.append(cur)
                    cur = ''; last = y
                cur += d.Range(k, k + 1).Text
            if cur: lines.append(cur)
        finally:
            d.Close(False)
        with tempfile.TemporaryDirectory() as t:
            dump = Path(t) / 'l.json'
            subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
            dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
        rows = {}
        for el in dd['pages'][0]['elements']:
            if el.get('type') == 'text' and el.get('text', '').strip() and el['text'] not in ('A', 'B'):
                rows.setdefault(round(el['y'], 2), []).append(el)
        ol = []
        for y in sorted(rows):
            es = sorted(rows[y], key=lambda e: e['x'])
            ol.append((''.join(e['text'] for e in es), round(es[-1]['x'] + es[-1].get('width', 0), 1)))
        w3 = lines[2][-3:] if len(lines) > 2 else ''
        o3 = (ol[2][0][-3:], ol[2][1]) if len(ol) > 2 else ''
        print(f"{name}: WORD {len(lines)} lines, l3 ends {w3!r} | OXI {len(ol)} lines, l3 ends {o3!r}")
finally:
    app.Quit()

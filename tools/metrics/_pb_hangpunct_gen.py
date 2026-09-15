# -*- coding: utf-8 -*-
"""How far may a line-end 。 hang past the right margin (ぶら下げ) on a lines grid?

policies__06c631e8bc061f40 (jablindC50): checklist lines «□ 運転手は、車両の点検
（ライト、ランプの動作確認等）をしている。» at 15pt ＭＳ ゴシック, line 440 exact,
hangingChars 100, A4 with 56.7pt margins (text 481.9pt), jc both, balance compat,
compressPunctuation. Word sets each in ONE line with the trailing 。 STARTING at
x=540 (1.4pt past the 538.6 margin) and hanging its full width; Oxi wraps «る。»
to a 2nd 22pt line -- half the checklist doubles, +1 page.

Arms: '□ ' + N kanji + '。' for N = 26..32 (each +15pt), family K (pure kanji)
and family Y (the real line's four 約物 、（、） in place, so compression can
also fire). Readout: Word line count and last char x / Oxi line count.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/hangpunct'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1191" w:right="1134" w:bottom="1191" w:left="1134" w:header="567" w:footer="340" w:gutter="0"/>'
        '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:hint="eastAsia"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr>'
KANJI = '運転手車両点検確認等動作状態健康出席管理責任者当日欠乗名簿反映職員園長主任担共有緊急連絡用携帯電話'


def para(text):
    return (f'<w:p><w:pPr><w:spacing w:line="440" w:lineRule="exact"/><w:ind w:left="300" w:hangingChars="100" w:hanging="300"/>{RPR}</w:pPr>'
            f'<w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def fam_k(n):
    return '□ ' + KANJI[:n] + '。'


def fam_y(n):
    # the real line's shape: 、 after 4 chars, （ after 9, 、 after 13, ） after 20 (positions inside the kanji run)
    body = KANJI[:n]
    ins = [(4, '、'), (9, '（'), (13, '、'), (20, '）')]
    out = ''; last = 0
    for pos, ch in ins:
        if pos <= len(body):
            out += body[last:pos] + ch; last = pos
    out += body[last:]
    return '□ ' + out + '。'


arms = {}
for n in range(26, 33):
    arms[f'K{n}'] = fam_k(n)
for n in range(22, 29):
    arms[f'Y{n}'] = fam_y(n)


def document(x):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p>{para(x)}<w:p><w:r><w:t>B</w:t></w:r></w:p>{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
try:
    for name, x in arms.items():
        at = OUT / f"{name}.docx"
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", document(x))
        d = app.Documents.Open(str(at.resolve()), False, True)
        try:
            p = d.Paragraphs(2); s, e = p.Range.Start, p.Range.End
            ys = sorted({round(d.Range(k, k).Information(6), 2) for k in range(s, e - 1)})
            lx = round(d.Range(e - 2, e - 2).Information(5), 2)   # the 。
            px = round(d.Range(e - 3, e - 3).Information(5), 2)   # the char before it
        finally:
            d.Close(False)
        with tempfile.TemporaryDirectory() as t:
            dump = Path(t) / 'l.json'
            subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
            dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
        oys = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and el.get('text', '').strip() and el['text'] not in ('A', 'B')})
        nat = 15 * 1 + 7.5 + 15 * (len(x) - 3)   # □ + space + chars (excl. 。)
        print(f"{name}: chars={len(x)} WORD lines {len(ys)} prev_x {px} last_x {lx} (margin 538.6, natural 。 start {round(76.5 + nat, 1)}) | OXI lines {len(oys)}")
finally:
    app.Quit()

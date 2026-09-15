# -*- coding: utf-8 -*-
"""How does Word size body lines on a linesAndChars grid whose charSpace is NEGATIVE?

technical__9e4d04b448f84674 (jablindC50): docGrid type=linesAndChars linePitch=291
(14.55pt) charSpace=-3531, legacy compat list (no compatibilityMode), Normal
10.5pt single. Word's page-1 pitches: an 11pt paragraph 21.0, a 12pt one 23.25 --
neither a grid multiple (14.55 / 29.1) nor the natural line (15.0 / 16.4).
Oxi snaps both to two rows (29.1). 1.4 x natural fits both (21.0 / 22.96).

Arms (three paragraphs A / X / B; X's run size varies; Info6 of all three):
  neg_105 / neg_11 / neg_12 / neg_14   the doc's grid, X at 10.5 / 11 / 12 / 14pt
  zero_11                              charSpace removed, X 11pt
  pos_11                               charSpace +3531, X 11pt
  neg_11_cm15                          the doc's grid + compatibilityMode 15, X 11pt
  lines_11                             type=lines linePitch 291 (no charSpace), X 11pt
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/negcharspace'); OUT.mkdir(parents=True, exist_ok=True)
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
LEGACY = ('<w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/>'
          '<w:adjustLineHeightInTable/><w:doNotBreakWrappedTables/><w:doNotSnapToGridInCell/><w:selectFldWithFirstOrLastChar/><w:doNotWrapTextWithPunct/>'
          '<w:doNotUseEastAsianBreakRules/><w:useWord2002TableStyleRules/><w:growAutofit/><w:useFELayout/><w:useNormalStyleForList/>'
          '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat>')


def settings(cm15=False):
    compat = LEGACY.replace('w:val="11"', 'w:val="15"') if cm15 else LEGACY
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>' + compat + '</w:settings>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>'
          '</w:styles>')


def sect(grid):
    return ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1418" w:right="851" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
            f'<w:cols w:space="425"/>{grid}</w:sectPr>')


NEG = '<w:docGrid w:type="linesAndChars" w:linePitch="291" w:charSpace="-3531"/>'
ZERO = '<w:docGrid w:type="linesAndChars" w:linePitch="291"/>'
POS = '<w:docGrid w:type="linesAndChars" w:linePitch="291" w:charSpace="3531"/>'
LINES = '<w:docGrid w:type="lines" w:linePitch="291"/>'


def para(text, sz=None):
    rpr = f'<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>' if sz else '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
    return f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'


arms = {}
if os.environ.get('CAP'):
    for n in (48, 49, 50, 51, 52, 53):
        arms[f'cap{n}'] = (NEG, None, False, '　' * (n - 2) + '１面')
        arms[f'cap{n}_11'] = (NEG, None, False, ('第２号様式', 22, '　' * (n - 7) + '１面'))
else:
    arms = {
        'neg_105': (NEG, None, False), 'neg_11': (NEG, 22, False), 'neg_12': (NEG, 24, False), 'neg_14': (NEG, 28, False),
        'zero_11': (ZERO, 22, False), 'pos_11': (POS, 22, False), 'neg_11_cm15': (NEG, 22, True), 'lines_11': (LINES, 22, False),
    }


def para_cap(spec):
    ind = '<w:ind w:leftChars="-100" w:left="-193"/>'
    if isinstance(spec, tuple):
        head, sz, rest = spec
        return (f'<w:p><w:pPr>{ind}<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr><w:t xml:space="preserve">{head}</w:t></w:r>'
                f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">{rest}</w:t></w:r></w:p>')
    return f'<w:p><w:pPr>{ind}<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">{spec}</w:t></w:r></w:p>'


def document(grid, sz, cap=None):
    x = para_cap(cap) if cap is not None else para("エックス線装置備付届", sz)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{para("あ")}{x}{para("い")}{sect(grid)}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
try:
    for name, arm in arms.items():
        grid, sz, cm15 = arm[0], arm[1], arm[2]; cap = arm[3] if len(arm) > 3 else None
        at = OUT / f"{name}.docx"
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings(cm15)); z.writestr("word/document.xml", document(grid, sz, cap))
        d = app.Documents.Open(str(at.resolve()), False, True)
        try:
            ys = [round(d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6), 2) for i in (1, 2, 3)]
            if cap is not None:
                r = d.Paragraphs(2).Range; e = r.End
                ys = ys + [('lines', r.ComputeStatistics(1), 'last_x', round(d.Range(e - 2, e - 2).Information(5), 2))]
        finally:
            d.Close(False)
        with tempfile.TemporaryDirectory() as t:
            dump = Path(t) / 'l.json'
            subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
            dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
        oy = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and el.get('text', '').strip()})
        if cap is not None:
            info = ys[-1]; ys = ys[:-1]
            print(f"{name:12s} WORD {info} y {ys} | OXI lines {len(oy) - 2} y {oy[:4]}")
            continue
        wd = [round(b - a, 2) for a, b in zip(ys, ys[1:])]
        od = [round(b - a, 2) for a, b in zip(oy, oy[1:])]
        print(f"{name:12s} WORD y {ys} d {wd} | OXI y {oy[:3]} d {od[:2]}")
finally:
    app.Quit()

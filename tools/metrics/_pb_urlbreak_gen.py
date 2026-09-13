# -*- coding: utf-8 -*-
"""Where does Word break a long Latin token (a URL) inside a JP paragraph?

technical__b80f6caa111ba50d (jablindD50, ＭＳ ゴシック 10.5, docGrid lines 360,
column 461.85pt): the paragraph «#　https://github.com/notepad-plus-plus/
notepad-plus-plus/releases/download/v7.9.5/npp.7.9.5.Installer.exe» is TWO
lines in Word -- line 1 ends at «notepad-plus-» (the last hyphen that fits,
278pt of 462) and line 2 carries «plus/releases/download/v7.9.5/npp.7.9.5.
Installer.exe» whole (283pt). So a hyphen is a break opportunity inside the
token; a slash and a dot are not. Oxi keeps the token whole, moves it to line
2 and margin-breaks it -> three lines, +18pt, one paragraph over the page.

Arms (same section/fonts as the doc; one paragraph «#　» + token, then A2):
  hyphen   token with hyphens only:  aaaa-bbbb-... (each chunk 8 chars, 13 chunks)
  slash    token with slashes only:  aaaa/bbbb/...
  dot      token with dots only:     aaaa.bbbb....
  mixed    the real URL
  en_mixed the real URL in an EN shape (Calibri 11, no docGrid, lang en-US)
  under    token with underscores only
Readout: Word lines via per-character Information(6) (text per line); Oxi
line count from the layout dump (distinct y within the paragraph).
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/urlbreak'); OUT.mkdir(parents=True, exist_ok=True)
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


def settings(balance):
    compat = ('<w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/>'
              '<w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/>') if balance else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>' + compat +
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def styles(jp):
    if jp:
        rf = '<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="ＭＳ ゴシック"/>'
        lang = '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
        sz = '<w:sz w:val="21"/><w:szCs w:val="22"/>'
    else:
        rf = '<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
        lang = '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>'
        sz = '<w:sz w:val="22"/><w:szCs w:val="22"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:docDefaults><w:rPrDefault><w:rPr>{rf}<w:kern w:val="2"/>{sz}{lang}</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>'
            '</w:styles>')


def sect(jp):
    grid = '<w:docGrid w:type="lines" w:linePitch="360"/>' if jp else '<w:docGrid w:linePitch="360"/>'
    return ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1335" w:bottom="1701" w:left="1334" w:header="851" w:footer="992" w:gutter="0"/>'
            f'<w:cols w:space="425"/>{grid}</w:sectPr>')


def para(t, ppr='', rpr=''):
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


L240 = '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
HINT = '<w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="ＭＳ ゴシック" w:hint="eastAsia"/></w:rPr>'


URL = 'https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v7.9.5/npp.7.9.5.Installer.exe'
chunks = ['abcdefgh'] * 13
arms = {
    'hyphen': (True, '#　' + '-'.join(chunks)),
    'slash': (True, '#　' + '/'.join(chunks)),
    'dot': (True, '#　' + '.'.join(chunks)),
    'under': (True, '#　' + '_'.join(chunks)),
    'mixed': (True, '#　' + URL),
    'en_mixed': (False, '# ' + URL),
    'mixed_balance': (True, '#　' + URL, True, '', ''),
    'mixed_bal_l240': (True, '#　' + URL, True, L240, ''),
    'mixed_bal_hint': (True, '#　' + URL, True, '', HINT),
    'mixed_hint': (True, '#　' + URL, False, '', HINT),
    'mixed_cjkbody': (True, '#　' + URL, False, '', '', '漢字の段落'),
}


def document(jp, x, ppr='', rpr='', a1='A1'):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{para(a1)}{para(x, ppr, rpr)}{para("A2")}{sect(jp)}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
try:
    for name, arm in arms.items():
        jp, x = arm[0], arm[1]
        balance, ppr, rpr = (arm[2], arm[3], arm[4]) if len(arm) > 2 else (False, '', '')
        a1 = arm[5] if len(arm) > 5 else 'A1'
        at = OUT / f"{name}.docx"
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", styles(jp))
            z.writestr("word/settings.xml", settings(balance))
            z.writestr("word/document.xml", document(jp, x, ppr, rpr, a1))
        d = app.Documents.Open(str(at.resolve()), False, True)
        try:
            p = d.Paragraphs(2); s, e = p.Range.Start, p.Range.End
            lines = []; last = None; cur = ''
            for k in range(s, e):
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
        els = [el for el in dd['pages'][0]['elements'] if el.get('type') == 'text']
        ys = sorted({round(el['y'], 2) for el in els})
        # paragraph 2 = everything between the A1 line and the A2 line
        ya1 = ys[0]; ya2 = ys[-1]
        oxi_lines = {}
        for el in els:
            y = round(el['y'], 2)
            if ya1 < y < ya2:
                oxi_lines.setdefault(y, []).append((el['x'], el['text']))
        ol = [''.join(t for _x, t in sorted(v)) for _y, v in sorted(oxi_lines.items())]
        print(f"=== {name}: WORD {len(lines)} lines | OXI {len(ol)} lines")
        for l in lines: print('   W', repr(l.rstrip(chr(13)))[:100])
        for l in ol: print('   O', repr(l)[:100])
finally:
    app.Quit()

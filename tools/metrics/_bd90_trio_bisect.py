# -*- coding: utf-8 -*-
"""bd90b00 大臣殿-region +1.5 bisect: faithful specimen (verbatim docDefaults +
a5 style + the two paragraphs), D-bracketed (same-class anchors -> ink
conventions cancel), then property-stripping variants.

Real-doc measured: h(年月日)+h(大臣殿) = Word 48.0 / Oxi 49.5 (+1.5).
G(D3->D4) = h(D3=16.5) + sum(specimen lines).

Run: python tools/metrics/_bd90_trio_bisect.py
"""
import os, sys, json, subprocess
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'gridquant')
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, 'bd90trio.docx')
PDF = os.path.join(OUTDIR, 'bd90trio.pdf')
MINCHO = pg.MINCHO

STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
    '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:line="254" w:lineRule="exact"/>'
    '<w:ind w:left="100" w:rightChars="95" w:right="95" w:hangingChars="100" w:hanging="100"/>'
    '<w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
    '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="a5"><w:name w:val="一太郎"/>'
    '<w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/>'
    '<w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/>'
    '<w:spacing w:line="210" w:lineRule="exact"/></w:pPr>'
    f'<w:rPr><w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}" w:cs="{MINCHO}"/>'
    '<w:spacing w:val="-1"/><w:kern w:val="0"/><w:szCs w:val="21"/></w:rPr></w:style>'
    '</w:styles>'
)

IND0 = ('<w:ind w:leftChars="0" w:left="0" w:rightChars="0" w:right="0" '
        'w:hangingChars="0" w:hanging="0" w:firstLineChars="0" w:firstLine="0"/>')

def dpara(txt):
    r = (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
         '<w:sz w:val="21"/>')
    return (f'<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>{IND0}'
            f'<w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{pg.esc(txt)}</w:t></w:r></w:p>')

def pagebreak():
    return ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:br w:type="page"/></w:r></w:p>')

def nengappi(strip=()):
    ppr = '<w:pStyle w:val="a5"/>' if 'style' not in strip else ''
    ppr += '<w:wordWrap/><w:spacing w:line="240" w:lineRule="auto"/>'
    ppr += '<w:ind w:right="199" w:firstLineChars="0" w:firstLine="0"/>'
    ppr += '<w:jc w:val="right"/>' if 'jc' not in strip else '<w:jc w:val="left"/>'
    sp = '' if 'cspace' in strip else '<w:spacing w:val="-9"/>'
    return (f'<w:p><w:pPr>{ppr}<w:rPr><w:color w:val="000000"/>{sp}</w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:color w:val="000000"/>{sp}</w:rPr>'
            '<w:t xml:space="preserve">　　年　　月　　日</w:t></w:r></w:p>')

def daijin(strip=()):
    ppr = '<w:pStyle w:val="a5"/>' if 'style' not in strip else ''
    if 'wordwrap' not in strip:
        ppr += '<w:wordWrap/>'
    ppr += '<w:spacing w:line="240" w:lineRule="auto"/>'
    if 'ind' not in strip:
        ppr += '<w:ind w:right="199" w:firstLineChars="100" w:firstLine="256"/>'
    ppr += '<w:jc w:val="left"/>'
    if 'pmark24' in strip:
        pm = '<w:rPr><w:color w:val="000000"/><w:spacing w:val="0"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
    else:
        pm = '<w:rPr><w:color w:val="000000"/><w:spacing w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    sp = '' if 'cspace' in strip else '<w:spacing w:val="-12"/>'
    r = f'<w:rPr><w:color w:val="000000"/>{sp}<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
    return (f'<w:p><w:pPr>{ppr}{pm}</w:pPr>'
            f'<w:r>{r}<w:t>厚生労働大臣</w:t></w:r>'
            f'<w:r>{r}<w:t xml:space="preserve">　</w:t></w:r>'
            f'<w:r>{r}<w:t>殿</w:t></w:r></w:p>')

CONFIGS = [
    ('A0 年月日 verbatim',            lambda: [nengappi()]),
    ('B0 大臣殿 verbatim',            lambda: [daijin()]),
    ('C0 both verbatim',              lambda: [nengappi(), daijin()]),
    ('B1 -charspacing',               lambda: [daijin(('cspace',))]),
    ('B2 pmark sz24->28',             lambda: [daijin(('pmark24',))]),
    ('B3 -wordWrap',                  lambda: [daijin(('wordwrap',))]),
    ('B4 -pStyle a5',                 lambda: [daijin(('style',))]),
    ('B5 -ind',                       lambda: [daijin(('ind',))]),
    ('A1 -pStyle a5',                 lambda: [nengappi(('style',))]),
    ('A2 -charspacing',               lambda: [nengappi(('cspace',))]),
]

def build():
    body = []
    for i, (_, mk) in enumerate(CONFIGS):
        if i:
            body.append(pagebreak())
        tag = f'ｘ{i+1:02d}'
        body.append(dpara(f'Ｄ１{tag}'))
        body.append(dpara(f'Ｄ２{tag}'))
        body.append(dpara(f'Ｄ３{tag}'))
        body.extend(mk())
        body.append(dpara(f'Ｄ４{tag}'))
        body.append(dpara(f'Ｄ５{tag}'))
    body.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" '
                'w:header="851" w:footer="567" w:gutter="0"/>'
                '<w:docGrid w:type="lines" w:linePitch="330"/></w:sectPr>')
    pg.write_docx(DOCX, pg.doc(''.join(body)),
                  extra_parts={'word/styles.xml': STYLES})

def anchors_pdf():
    import fitz
    d = fitz.open(PDF)
    out = {}
    for p in range(len(d)):
        for b in d[p].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                txt = ''.join(s['text'] for s in l.get('spans', [])).replace(' ', '')
                if txt.startswith(('Ｄ３', 'Ｄ４')):
                    out[txt[:5]] = round(l['bbox'][1], 2)
    return out

def anchors_oxi():
    dumpjson = os.path.join(OUTDIR, 'bd90trio_oxi.json')
    env = dict(os.environ)
    for v in ('OXI_ROWBOX2', 'OXI_OFFSLOT_INK', 'OXI_S638_ALL'):
        env[v] = '1'
    subprocess.run(['./tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe',
                    DOCX, os.path.join(OUTDIR, 'bd90trio_png'),
                    '--dump-layout=' + dumpjson], capture_output=True, env=env)
    dump = json.load(open(dumpjson, encoding='utf-8'))
    out = {}
    for page in dump['pages']:
        rows = {}
        for e in page['elements']:
            if e['type'] == 'text' and e.get('text', '').strip():
                rows.setdefault(round(e['y'], 2), []).append((e['x'], e['text']))
        for y, frs in rows.items():
            t = ''.join(t for _, t in sorted(frs)).replace('　', '')
            if t.startswith(('Ｄ３', 'Ｄ４')):
                out[t[:5]] = y
    return out

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    build()
    print('wrote', DOCX)
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
    doc.ExportAsFixedFormat(PDF, 17)
    doc.Close(False)
    word.Quit()
    wa = anchors_pdf()
    oa = anchors_oxi()
    print(f'{"config":<22} {"G_word":>8} {"G_oxi":>8} {"d":>7}   (G = 16.5 + specimen lines)')
    for i, (name, _) in enumerate(CONFIGS):
        tag = f'ｘ{i+1:02d}'
        try:
            gw = round(wa[f'Ｄ４{tag}'] - wa[f'Ｄ３{tag}'], 2)
            go = round(oa[f'Ｄ４{tag}'] - oa[f'Ｄ３{tag}'], 2)
            print(f'{name:<22} {gw:8.2f} {go:8.2f} {go-gw:+7.2f}')
        except KeyError:
            print(f'{name:<22} MISSING')

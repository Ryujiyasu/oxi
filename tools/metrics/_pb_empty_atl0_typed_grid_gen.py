# -*- coding: utf-8 -*-
"""An empty `line=0 atLeast` paragraph on a typed (lines) docGrid: which face
sizes it -- the ASCII face (Century) or the eastAsia face (ＭＳ 明朝)?

ja/administrative__05e3143c: the anchor-only paragraph between the masthead
table and the boxed heading (atLeast 0, mark rPr = szCs 16 + bdr, docDefaults
ascii Century / eastAsia ＭＳ 明朝 10.5, docGrid lines 360) is 12.72 tall in
Word (table bottom 95.66 -> heading box top 108.38); Oxi gives it ＭＳ 明朝's
13.62.

Faithful slice of that document (its styles, theme, settings and sectPr): the
body becomes «あいう» / X / «かきく», the text paragraphs ＭＳ 明朝 10.5 atLeast 0,
X = nothing (control), an empty paragraph (mark plain / with bdr / with
rFonts ascii ＭＳ 明朝), or an anchor-only paragraph with the bdr mark.
Readout: the baseline gap «あいう» -> «かきく» minus the control's.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/administrative/05e3143c61176854.docx'
OUT = 'tests/fixtures/empty_atl0_typed_grid'; os.makedirs(OUT, exist_ok=True)
EMU = 12700
SP = '<w:spacing w:line="0" w:lineRule="atLeast"/>'
MIN = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
BDR = '<w:szCs w:val="32"/><w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>'


def text_p(t):
    return f'<w:p><w:pPr>{SP}<w:rPr>{MIN}</w:rPr></w:pPr><w:r><w:rPr>{MIN}</w:rPr><w:t>{t}</w:t></w:r></w:p>'


ANCHOR = ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" '
          'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
          f'<wp:positionH relativeFrom="page"><wp:posOffset>{400 * EMU}</wp:posOffset></wp:positionH>'
          f'<wp:positionV relativeFrom="page"><wp:posOffset>{400 * EMU}</wp:posOffset></wp:positionV>'
          f'<wp:extent cx="{40 * EMU}" cy="{40 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="901" name="P"/><wp:cNvGraphicFramePr/>'
          '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
          f'<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{40 * EMU}" cy="{40 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
          '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
ARMS = {
    'none': '',
    'empty_plain': f'<w:p><w:pPr>{SP}</w:pPr></w:p>',
    'empty_bdr': f'<w:p><w:pPr>{SP}<w:rPr>{BDR}</w:rPr></w:pPr></w:p>',
    'empty_mincho': f'<w:p><w:pPr>{SP}<w:rPr>{MIN}</w:rPr></w:pPr></w:p>',
    'anchor_bdr': f'<w:p><w:pPr>{SP}<w:rPr>{BDR}</w:rPr></w:pPr>{ANCHOR}</w:p>',
    'anchor_plain': f'<w:p><w:pPr>{SP}</w:pPr>{ANCHOR}</w:p>',
}


def build(name, mid):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                x = x[:b0] + text_p('あいう') + mid + text_p('かきく') + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    paths = {n: build(n, m) for n, m in ARMS.items()}
    if '--build' in sys.argv:
        sys.exit()
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    res = {}
    try:
        for n, at in paths.items():
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
            P = pymupdf.open(at[:-5] + '.pdf')[0]
            ys = {}
            for b in P.get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    for s in l['spans']:
                        if s['text'].strip() in ('あいう', 'かきく'):
                            ys[s['text'].strip()] = s['origin'][1]
            res[n] = ys['かきく'] - ys['あいう']
            print(f'{n:14} gap {res[n]:.2f}  extra {res[n] - res.get("none", res[n]):+.2f}', flush=True)
    finally:
        app.Quit()

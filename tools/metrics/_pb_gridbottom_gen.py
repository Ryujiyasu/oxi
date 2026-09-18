"""Where does Word stop putting lines on a docGrid CJK page?

Why: `policies__094c44cd5dce58a8` (JA blind, markers off) puts one 「例示」
paragraph on page 39 that Word starts on page 40.  Everything else about the
paragraph matches -- the line advance (Word 13.56 / Oxi 13.62, compared at the
span ORIGIN, not the bbox top) and the `beforeLines="50"` gap (Word ~6.0 /
Oxi 5.88).  The difference is the page-bottom test: Oxi keeps a line whose
glyph bottom is 769.6 against a content bottom of 770.04, and the deepest
baseline Word uses anywhere in that 48-page document is 764.88.

So Word's limit sits somewhere in [764.88, 767.50) and the ink rule
(`ascent+descent must fit`, Day-33 part 65) is too generous here.  Guessing a
constant from one real document at 0.2pt precision would be reckless -- the
page-bottom test carries the whole corpus -- so this probe measures the limit
directly.

Method: one gridded CJK paragraph of many lines on A4 with a fixed BOTTOM
margin (content bottom stays 770.04).  Sweep the TOP margin in 5tw (0.25pt)
steps over slightly more than one line advance.  For each arm, read the
deepest baseline Word actually placed on page 1.  As the top margin grows the
deepest baseline grows with it until one more step pushes that line to page 2,
at which point the deepest baseline drops by a full advance.  The MAXIMUM over
the sweep is Word's limit, to the step size.

Usage:
  python _pb_gridbottom_gen.py run [arm ...]     # default: all ARMS
  python _pb_gridbottom_gen.py list
"""
import os
import sys
import zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_gridbottom")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')

SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {W_NS}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>')


def styles(font, sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:styles {W_NS}><w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>')


PAGE_W, PAGE_H = 11906, 16840        # A4 portrait, twips
SIDE, BOTTOM = 1797, 1440            # as in policies__094c44cd
CONTENT_BOTTOM = PAGE_H / 20.0 - BOTTOM / 20.0

# (name, font, half-point size, docGrid linePitch tw, snapToGrid, beforeLines)
ARMS = [
    ("mincho105_g360", "ＭＳ 明朝", 21, 360, True, 0),
    ("mincho105_bl50", "ＭＳ 明朝", 21, 360, True, 50),
    ("mincho105_g360_nosnap", "ＭＳ 明朝", 21, 360, False, 0),
    ("mincho12_g360", "ＭＳ 明朝", 24, 360, True, 0),
    ("gothic105_g360", "ＭＳ ゴシック", 21, 360, True, 0),
]
LINE = "本日は晴天なりこの行は頁末の判定を測るための行である"
N_LINES = 60


def build(arm, top_tw):
    _, font, sz, pitch, snap, bl = arm
    rpr = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/>'
    sg = '' if snap else '<w:snapToGrid w:val="0"/>'
    bls = (f'w:beforeLines="{bl}" w:before="120"' if bl else 'w:before="0"')
    paras = ''.join(
        '<w:p><w:pPr>'
        f'<w:spacing {bls} w:after="0" w:line="240" w:lineRule="auto"/>'
        f'{sg}<w:rPr>{rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{i:02d}{LINE}</w:t></w:r></w:p>'
        for i in range(N_LINES))
    sect = (f'<w:sectPr><w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}" w:code="9"/>'
            f'<w:pgMar w:top="{top_tw}" w:right="{SIDE}" w:bottom="{BOTTOM}" '
            f'w:left="{SIDE}" w:header="720" w:footer="720" w:gutter="0"/>'
            f'<w:docGrid w:linePitch="{pitch}"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{paras}{sect}</w:body></w:document>')
    os.makedirs(OUTDIR, exist_ok=True)
    path = os.path.join(OUTDIR, f'gb_{arm[0]}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/styles.xml', styles(font, sz))
        z.writestr('word/document.xml', doc)
    return path


class Word:
    def __init__(self):
        import win32com.client
        self.w = win32com.client.Dispatch('Word.Application')
        self.w.Visible = False
        self.n = 0

    def close(self):
        try:
            self.w.Quit()
        except Exception:
            pass

    def page1(self, path):
        """(deepest baseline on page 1, that line's text, n lines on page 1)."""
        import pymupdf
        pdf = path[:-5] + '.pdf'
        if os.path.exists(pdf):
            os.remove(pdf)
        d = self.w.Documents.Open(os.path.abspath(path), ReadOnly=True)
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
        self.n += 1
        doc = pymupdf.open(pdf)
        deepest, text, count = None, '', 0
        for b in doc[0].get_text('dict')['blocks']:
            if b.get('type') != 0:
                continue
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                if not t:
                    continue
                count += 1
                bl = max(s['origin'][1] for s in l['spans'])
                if deepest is None or bl > deepest:
                    deepest, text = bl, t
        return deepest, text, count


def run_arm(w, arm, lo_tw, hi_tw, step_tw):
    print(f'=== {arm[0]}  content_bottom={CONTENT_BOTTOM:.2f}')
    best = None
    prev_count = None
    for top in range(lo_tw, hi_tw + 1, step_tw):
        bl, text, count = w.page1(build(arm, top))
        if bl is None:
            continue
        flip = prev_count is not None and count != prev_count
        if best is None or bl > best[0]:
            best = (bl, top, text)
        if flip:
            print(f'    top={top / 20.0:7.2f}pt  lines={count:3d}  deepest_baseline={bl:7.2f}  <- flip')
        prev_count = count
    if best:
        print(f'    MAX deepest baseline = {best[0]:.2f} at top={best[1] / 20.0:.2f}pt '
              f'({CONTENT_BOTTOM - best[0]:.2f}pt above the content bottom) {best[2][:14]!r}')
    return best


def main():
    args = [a for a in sys.argv[1:] if a != 'run']
    if args and args[0] == 'list':
        for a in ARMS:
            print(a[0], a[1], a[2] / 2, 'pitch', a[3], 'snap', a[4], 'beforeLines', a[5])
        return 0
    sel = [a for a in ARMS if not args or a[0] in args]
    w = Word()
    try:
        for arm in sel:
            # one full 18pt advance, 0.25pt steps
            run_arm(w, arm, 1440, 1440 + 380, 5)
    finally:
        print(f'({w.n} Word renders)')
        w.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

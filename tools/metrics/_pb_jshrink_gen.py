"""Compat-15 justified space-shrink ALLOW as a function of the SPACE COUNT.

Why: S825 models the compat-15 justify-shrink capacity as a PER-SPACE credit
(0.25 x em_space, unbounded in the space count). educational__0036ed4b
(Times New Roman **10pt**, A4, jc=both, compat 15 explicit) contradicts it:
at 16 spaces Word KEEPS a line 3.75pt over the available width but WRAPS one
7.09pt over -- an allowance in (3.75, 7.09], while S825 grants 16 x 0.625 =
10.0pt.  A flat per-LINE cap (the shape S933 uses for the no-compat class)
fixes that doc but costs three others (EN 296 -> 293), so neither shape is
Word's rule.  This probe measures ALLOW(n_spaces) directly.

Method (all widths RENDER-measured, never metrics-computed -- the S825b
lesson): for one (font, size, word) config,

  natural(k)  = width of line 1 of a LEFT-aligned doc holding exactly k words
                on a page wide enough not to wrap (PDF line bbox, no trailing
                space in the text).
  avail_min(k)= the smallest available width at which the JUSTIFIED doc still
                puts k words on line 1 (binary search over the right margin).
  allow(k)    = natural(k) - avail_min(k)          [n_spaces = k - 1]

A per-space model predicts allow(k) linear in k with slope ~0.25 x space;
a per-line model predicts allow(k) flat.  Every doc carries a settings.xml
declaring compatibilityMode 15 -- without it the host falls into the S933
no-compat class and the allowance collapses to fs/4.

Usage:
  python _pb_jshrink_gen.py run [cfg ...]     # default: all CFGS
  python _pb_jshrink_gen.py list
"""
import os
import sys
import zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_jshrink")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
            '</Relationships>')

SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {W_NS}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>')

# (name, font, half-point size, word)
CFGS = [
    ("tnr10", "Times New Roman", 20, "mnopq"),
    ("tnr10w", "Times New Roman", 20, "mnopqrstuv"),
    ("tnr12", "Times New Roman", 24, "mnopq"),
    ("cal11", "Calibri", 22, "mnopq"),
]

PAGE_W_TW = 11906          # A4 portrait
LEFT_TW = 1440
WIDE_PAGE_TW = 31000       # for the natural-width docs


def _rpr(cfg):
    _, font, sz, _ = cfg
    return (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')


def _para(cfg, n_words, justify):
    rpr = _rpr(cfg)
    word = cfg[3]
    runs = []
    for i in range(n_words):
        if i:
            runs.append(f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve"> </w:t></w:r>')
        runs.append(f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{word}</w:t></w:r>')
    jc = '<w:jc w:val="both"/>' if justify else '<w:jc w:val="left"/>'
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
            f'w:lineRule="auto"/>{jc}<w:rPr>{rpr}</w:rPr></w:pPr>'
            f'{"".join(runs)}</w:p>')


def _document(body, page_w, right_tw):
    sect = (f'<w:sectPr><w:pgSz w:w="{page_w}" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="{right_tw}" w:bottom="1440" '
            f'w:left="{LEFT_TW}" w:header="709" w:footer="709" w:gutter="0"/>'
            '</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}{sect}</w:body></w:document>')


def _write(path, doc_xml):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml)


def build_natural(cfg, k):
    """Left-aligned doc holding exactly k words on one very wide line."""
    return _document(_para(cfg, k, False), WIDE_PAGE_TW, 720)


def build_justified(cfg, right_tw, n_words=90):
    return _document(_para(cfg, n_words, True), PAGE_W_TW, right_tw)


# ---------------------------------------------------------------- measurement

class Renderer:
    """One Word instance; docx -> pdf -> first text line (text, x0, x1)."""

    def __init__(self):
        import win32com.client
        self.word = win32com.client.Dispatch('Word.Application')
        self.word.Visible = False
        self.n = 0
        os.makedirs(OUTDIR, exist_ok=True)

    def close(self):
        try:
            self.word.Quit()
        except Exception:
            pass

    def line1(self, doc_xml, tag):
        import pymupdf
        docx = os.path.abspath(os.path.join(OUTDIR, f'{tag}.docx'))
        pdf = docx[:-5] + '.pdf'
        _write(docx, doc_xml)
        if os.path.exists(pdf):
            os.remove(pdf)
        d = self.word.Documents.Open(docx, ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17)
        d.Close(False)
        self.n += 1
        page = pymupdf.open(pdf)[0]
        best = None
        for blk in page.get_text('dict')['blocks']:
            if blk.get('type') != 0:
                continue
            for ln in blk['lines']:
                t = ''.join(s['text'] for s in ln['spans'])
                if not t.strip():
                    continue
                if best is None or ln['bbox'][1] < best[1]:
                    best = (t, ln['bbox'][1], ln['bbox'][0], ln['bbox'][2])
        if best is None:
            return None
        return best[0], best[2], best[3]


def words_on_line1(r, cfg, right_tw):
    got = r.line1(build_justified(cfg, right_tw), f'{cfg[0]}_j{right_tw:05d}')
    return 0 if got is None else len(got[0].split())


def natural(r, cfg, k):
    got = r.line1(build_natural(cfg, k), f'{cfg[0]}_n{k:03d}')
    if got is None:
        return None
    text, x0, x1 = got
    if len(text.split()) != k:
        return None          # wrapped: page not wide enough
    return x1 - x0


def avail_pt(right_tw):
    return (PAGE_W_TW - LEFT_TW - right_tw) / 20.0


def run_cfg(r, cfg):
    name = cfg[0]
    print(f'=== {name}  {cfg[1]} {cfg[2] / 2:g}pt  word={cfg[3]!r}')
    # Bracket the right margin so that avail runs from ~470pt down to ~150pt.
    hi_tw = PAGE_W_TW - LEFT_TW - int(470 * 20)      # widest avail
    lo_tw = PAGE_W_TW - LEFT_TW - int(150 * 20)      # narrowest avail
    k_hi = words_on_line1(r, cfg, hi_tw)
    k_lo = words_on_line1(r, cfg, lo_tw)
    print(f'    k at avail 470pt = {k_hi}, at 150pt = {k_lo}')
    rows = []
    for k in range(k_lo + 1, k_hi + 1):
        # smallest avail (largest right margin) at which line 1 still holds >= k
        lo, hi = hi_tw, lo_tw                        # lo_tw > hi_tw in twips
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if words_on_line1(r, cfg, mid) >= k:
                lo = mid
            else:
                hi = mid
        a = avail_pt(lo)
        nat = natural(r, cfg, k)
        if nat is None:
            print(f'    k={k:2d}  natural FAILED')
            continue
        rows.append((k, nat, a, nat - a))
        print(f'    k={k:2d} spaces={k - 1:2d} natural={nat:7.2f} '
              f'avail_min={a:7.2f} allow={nat - a:6.2f} '
              f'per_space={(nat - a) / max(k - 1, 1):5.3f}')
    return rows


def main():
    args = sys.argv[1:]
    if args and args[0] == 'list':
        for c in CFGS:
            print(c[0], c[1], c[2] / 2, c[3])
        return 0
    names = [a for a in args if a != 'run']
    sel = [c for c in CFGS if not names or c[0] in names]
    r = Renderer()
    try:
        for cfg in sel:
            run_cfg(r, cfg)
    finally:
        print(f'({r.n} Word renders)')
        r.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""How much room does a clamped CENTER tab keep after the preceding text?

blind-J en/policies__00d4d26d footer: «Date Published: 8th July 2025 ⎵[tab at a
4513tw centre stop]⎵⎵⎵… Last Updated … Review Date: 8th July 2026». The segment
after the tab is too wide to centre on the stop, so it is clamped left. Word
draws the tab with zero width (the space run before it is the only gap) and
keeps «2026» on the line, ending 1.5pt short of the right margin. Oxi's S958
slack keeps one more space width as a separator and wraps «2026».

Faithful slice (the document itself, footer1 edited): the space run before the
tab kept / removed (nosp), the segment's own leading spaces removed (lead), both
(nosp_lead). Readout: the x gap between the last glyph before the tab
and the first one after it, and whether «2026» shares the line.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/en/policies/00d4d26dfb71d6db.docx'
OUT = 'tests/fixtures/center_tab_sep'; os.makedirs(OUT, exist_ok=True)
SP_TAB = b'<w:t xml:space="preserve"> </w:t></w:r><w:r><w:tab/></w:r>'


def build(arm):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, f'{arm}.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/footer1.xml':
                s = re.sub(rb' w:rsid\w*="\w+"', b'', data)
                assert s.count(SP_TAB) == 1
                if arm in ('nosp', 'nosp_lead'):
                    s = s.replace(SP_TAB, b'<w:t></w:t></w:r><w:r><w:tab/></w:r>')
                if arm in ('lead', 'nosp_lead'):
                    # the segment's own leading spaces dropped: it starts at «Last»
                    lead = b'<w:r><w:t xml:space="preserve">         </w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'
                    assert s.count(lead) == 1
                    s = s.replace(lead, b'')
                data = s
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    for arm in sys.argv[1:] or ['orig', 'nosp', 'lead', 'nosp_lead']:
        at = build(arm)
        app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
        finally:
            app.Quit()
        P = pymupdf.open(at[:-5] + '.pdf')[0]
        chars = [(round(s['origin'][1], 1), c['bbox'][0], c['bbox'][2], c['c'])
                 for b in P.get_text('rawdict')['blocks'] for l in b.get('lines', [])
                 for s in l['spans'] for c in s['chars'] if s['origin'][1] > 700]
        five = next(c for c in chars if c[3] == '5' and c[1] > 190 and c[1] < 200)
        after = sorted(c for c in chars if c[1] > five[2] - 0.1 and c[0] == five[0])
        y26 = [c for c in chars if c[3] == '6' and c[1] > 500]
        print(f'{arm}: «5» ends {five[2]:.2f}; next glyphs {[(c[3], round(c[1], 2)) for c in after[:3]]}; '
              f'«6» of 2026 at {[(c[0], round(c[1], 1)) for c in y26]}', flush=True)

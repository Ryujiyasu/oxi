# -*- coding: utf-8 -*-
"""First-line indent channels (twips vs chars) -- Word PDF probe on a faithful
slice of zh dev creative__0f11c107 (Normal: firstLineChars=200 firstLine=200 at
12pt; numId 2 = bullet level left 480 hanging 480). Prints the x of «X» (the
first text glyph) per arm; the left margin is 42.55.

    python tools/metrics/_pb_ind_chars_channels_gen.py <out_dir>

Arms n_* are list paragraphs, p_* plain ones, hc_* / lc_* / l_* use a style HC
(left=960 hangingChars=200 hanging=300). Word 2026-11-10: n_flc0 66.6,
n_none 114.6, n_fl0 114.6, n_flc0_fl0 90.6, n_flc100 90.6, n_hc0 66.6,
n_left0_flc0 42.6, p_flc0 52.6, p_none 66.6, p_fl0 66.6, p_d0_360 60.6,
p_d200* 66.6; hc_none 42.6, hc_flc0 75.6, hc_hc0 75.6, hc_fl0 42.6,
hc_hang0 42.6; lc_flc0 54.6 (continuation 69.6), lc_only 54.6 (78.6),
l_flc0 39.6, lc0_flc0 75.6.
"""
import os, re, sys, zipfile
import fitz, win32com.client

sys.stdout.reconfigure(encoding='utf-8')
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/creative/0f11c10782e6f2a6.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
st = parts['word/styles.xml'].decode('utf8')
parts['word/styles.xml'] = st.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="HC"><w:name w:val="HC"/><w:basedOn w:val="a"/><w:pPr><w:ind w:left="960" w:hangingChars="200" w:hanging="300"/></w:pPr></w:style></w:styles>').encode('utf8')
doc = parts['word/document.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + 8]; sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]
TXT = 'XQZ' + '什么是真正的创新' * 6
R = f'<w:r><w:rPr><w:rFonts w:ascii="DengXian" w:eastAsia="DengXian" w:hAnsi="DengXian"/></w:rPr><w:t>{TXT}</w:t></w:r>'


def P(ind, num=True, style=''):
    n = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' if num else ''
    s = f'<w:pStyle w:val="{style}"/>' if style else ''
    return f'<w:p><w:pPr>{s}{n}<w:spacing w:before="0" w:beforeAutospacing="0" w:after="0" w:afterAutospacing="0"/>{ind}</w:pPr>{R}</w:p>'


ARMS = {
    'n_flc0': P('<w:ind w:firstLineChars="0"/>'),
    'n_none': P(''),
    'n_fl0': P('<w:ind w:firstLine="0"/>'),
    'n_flc0_fl0': P('<w:ind w:firstLineChars="0" w:firstLine="0"/>'),
    'n_flc100': P('<w:ind w:firstLineChars="100"/>'),
    'n_hc0': P('<w:ind w:hangingChars="0"/>'),
    'n_left0_flc0': P('<w:ind w:left="0" w:firstLineChars="0"/>'),
    'p_flc0': P('<w:ind w:firstLineChars="0"/>', False),
    'p_none': P('', False),
    'p_fl0': P('<w:ind w:firstLine="0"/>', False),
    'p_d200_200': P('<w:ind w:firstLineChars="200" w:firstLine="200"/>', False),
    'p_d200_360': P('<w:ind w:firstLineChars="200" w:firstLine="360"/>', False),
    'p_d200_0': P('<w:ind w:firstLineChars="200" w:firstLine="0"/>', False),
    'p_d0_360': P('<w:ind w:firstLineChars="0" w:firstLine="360"/>', False),
    'p_d200': P('<w:ind w:firstLineChars="200"/>', False),
    'hc_none': P('', False, 'HC'),
    'hc_flc0': P('<w:ind w:firstLineChars="0"/>', False, 'HC'),
    'hc_hc0': P('<w:ind w:hangingChars="0"/>', False, 'HC'),
    'hc_fl0': P('<w:ind w:firstLine="0"/>', False, 'HC'),
    'hc_hang0': P('<w:ind w:hanging="0"/>', False, 'HC'),
    'lc_flc0': P('<w:ind w:leftChars="100" w:left="240" w:firstLineChars="0"/>', False, 'HC'),
    'lc_only': P('<w:ind w:leftChars="100" w:left="240"/>', False, 'HC'),
    'l_flc0': P('<w:ind w:left="240" w:firstLineChars="0"/>', False, 'HC'),
    'lc0_flc0': P('<w:ind w:leftChars="0" w:firstLineChars="0"/>', False, 'HC'),
}

w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    for name, body in ARMS.items():
        dst = os.path.join(OUT, name + '.docx')
        xml = head + body + sect + '</w:body></w:document>'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items():
                o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
        pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
        xs = []
        for b in fitz.open(pdf)[0].get_text('rawdict')['blocks']:
            for l in b.get('lines', []):
                c = l['spans'][0]['chars'][0]
                if c['origin'][1] > 60:
                    xs.append(round(c['origin'][0], 1))
        print(name, xs[:3])
finally:
    w.Quit()

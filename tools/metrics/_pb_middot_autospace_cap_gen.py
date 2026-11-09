"""Faithful-slice probe: how much a justified CJK line may overflow and still be kept by
squeezing one «・» and the Latin/CJK auto-space gap (blind-J ja/reference__6c5dcb56's
«2020-2021年度　4R2Z　ゾーン・チェアパーソン», HG丸ｺﾞｼｯｸM-PRO 12, measure 288.6pt).

Arms: A ・ + gap, B ・ no gap (autoSpace off), C ア no gap-less (gap only), D neither.
The Latin chunk replacing «4R2Z» steps the natural overflow by ~0.25pt.
Usage: python _pb_middot_autospace_cap_gen.py <src.docx> <out_dir>
Writes <out_dir>/probe.docx and prints, per arm, overflow -> lines (1 kept / 2 wrapped).
"""
import sys, os, re, zipfile, itertools
from fontTools.ttLib import TTFont

src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
F = 'HG丸ｺﾞｼｯｸM-PRO'
t = TTFont('C:/Windows/Fonts/HGRSMP.TTF'); cm = t.getBestCmap(); hm = t['hmtx']
W = lambda s: sum(hm[cm[ord(c)]][0] for c in s) / 256 * 12
AVAIL = (5866 - 94) / 20
base_latin = W('4R2Z')

# Latin chunks: width offsets from «4R2Z» on a 0.05pt grid (letters/digits only --
# «.» «|» add break chances of their own)
alphabet = '4RZilrtsxcjIabdefghkmnopquvwyz0'
chunks = {}
for n in range(2, 5):
    for combo in itertools.product(alphabet, repeat=n):
        s_ = ''.join(combo)
        w = W(s_) - base_latin
        key = round(w * 20) / 20
        if -40 <= key <= 20 and abs(w - key) < 0.012 and key not in chunks:
            chunks[key] = s_

ARMS = {'A': ('・', 1), 'F': ('・', 2), 'G': ('・', 'G'), 'H': ('・', 'H'), 'E': ('ア', 2)}
TARGET = {'A': (6.5, 7.6), 'F': (5.9, 7.0), 'G': (5.9, 7.6), 'H': (5.9, 7.6), 'E': (2.7, 3.2)}
rpr = f'<w:rPr><w:rFonts w:ascii="{F}" w:eastAsia="{F}" w:hAnsi="{F}" w:hint="eastAsia"/><w:kern w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'


def build(arm, L):
    mid, gap = ARMS[arm]
    if gap == 2:
        text = f'2020-2021　年度{L}ゾーン{mid}チェアパーソン'
    elif gap == 'G':
        return f'2020-2021　年度　{L}ゾーン{mid}チェアパーソン', W(f'2020-2021　年度　{L}ゾーン{mid}チェアパーソン') + 3.0 - AVAIL
    elif gap == 'H':
        return f'2020-2021　年度{L}　ゾーン{mid}チェアパーソン', W(f'2020-2021　年度{L}　ゾーン{mid}チェアパーソン') + 3.0 - AVAIL
    else:
        text = f'2020-2021年度　{L}　ゾーン{mid}チェアパーソン'
    return text, W(text) + 3.0 * gap - AVAIL


paras, meta = [], []
# fixed text per arm («4R2Z»); the overflow is stepped 1 twip at a time with w:right
TARGET = {'A': (6.5, 7.8), 'B': (5.7, 6.3), 'C': (1.2, 1.8), 'E': (2.7, 3.3), 'F': (5.7, 7.8), 'G': (6.5, 7.8), 'H': (6.5, 7.8)}
ARMS = {'A': ('・', 1), 'B': ('・', 0), 'C': ('ア', 1), 'E': ('ア', 2), 'F': ('・', 2), 'G': ('・', 'G'), 'H': ('・', 'H')}
for arm, (mid, gap) in ARMS.items():
    text, ov0 = build(arm, '4R2Z')
    lo, hi = TARGET[arm]
    for k in range(int(round(lo * 20)), int(round(hi * 20)) + 1, 2):
        right = k - int(round(ov0 * 20))          # twips; overflow = ov0 + right/20
        aspr = '' if gap else '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'
        paras.append(f'<w:p><w:pPr>{aspr}<w:ind w:leftChars="45" w:left="94" w:right="{right}"/>{rpr}</w:pPr><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')
        meta.append((arm, right, round(ov0 + right / 20, 2), text))

z = zipfile.ZipFile(src)
doc = z.read('word/document.xml').decode('utf8')
head = doc[:doc.index('<w:body>') + len('<w:body>')]
sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="567" w:right="5019" w:bottom="567" '
        'w:left="1021" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:linePitch="360"/></w:sectPr>')
body = head + ''.join(paras) + sect + '</w:body></w:document>'
dst = os.path.join(out, 'probe.docx')
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        if it.filename == 'word/document.xml':
            o.writestr(it, body)
        elif re.match(r'word/(header|footer)\d*\.xml', it.filename) or it.filename == 'word/_rels/document.xml.rels':
            o.writestr(it, z.read(it.filename))
        else:
            o.writestr(it, z.read(it.filename))

import win32com.client
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    dd = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
    for i, (arm, d, ov, text) in enumerate(meta, 1):
        r = dd.Paragraphs(i).Range
        y0 = dd.Range(r.Start, r.Start).Information(6)
        y1 = dd.Range(r.End - 2, r.End - 2).Information(6)
        x1 = dd.Range(r.End - 2, r.End - 2).Information(5)
        print(arm, f'{ov:+.2f}', 1 if abs(y1 - y0) < 1 else 2, round(x1, 2), ascii(text[9:20]))
    dd.Close(False)
finally:
    w.Quit()

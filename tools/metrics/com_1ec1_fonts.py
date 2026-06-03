# -*- coding: utf-8 -*-
"""S492s — COM-confirm Word's resolved EastAsia font on 1ec1's title/heading runs
(Lever E). 1ec1 theme: majorFont ea='' + Jpan=MS Gothic; minorFont ea='' + Jpan=MS
Mincho. S323 suppresses Jpan -> Oxi major falls to MS Mincho. Does Word use MS Gothic
(Jpan) for the title/headings? If yes, Oxi is wrong (Lever E real). Re-derive vs the
S323 origin (d1e8ac8 body). cp932-safe (UTF-8 file, ASCII-safe output: ord+name)."""
import os, glob
import win32com.client as w32
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/1ec1091177b1_006.docx')[0])
word = w32.DispatchEx('Word.Application'); word.Visible = False
out = []
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        for i, p in enumerate(doc.Paragraphs):
            if i >= 12:
                break
            rng = p.Range
            txt = rng.Text.replace('\r', '').replace('\x07', '')
            if not txt.strip():
                continue
            fe = rng.Font.NameFarEast
            fa = rng.Font.NameAscii
            sz = rng.Font.Size
            sty = ''
            try:
                sty = p.Style.NameLocal
            except Exception:
                pass
            # ASCII-safe: emit font name codepoints + a guess
            fec = '+'.join(hex(ord(c)) for c in (fe or '')[:6])
            out.append({'i': i, 'NameFarEast_codepts': fec, 'NameFarEast_raw': fe,
                        'ascii': fa, 'size': sz, 'style_codepts': '+'.join(hex(ord(c)) for c in (sty or '')[:8]),
                        'text_codepts': '+'.join(hex(ord(c)) for c in txt[:8])})
    finally:
        doc.Close(False)
finally:
    word.Quit()
# MS Gothic = ＭＳ ゴシック (FF2D FF33 3000 30B4...), MS Mincho = ＭＳ 明朝 (FF2D FF33 3000 660E 671D)
import json
print("MS Gothic ea = 0xff2d+0xff33+0x3000+0x30b4... ; MS Mincho = ...+0x660e+0x671d")
for o in out:
    print(o['i'], 'ea=', o['NameFarEast_codepts'], 'sz=', o['size'], 'style=', o['style_codepts'])
json.dump(out, open('c:/tmp/1ec1_fonts.json', 'w', encoding='utf-8'), ensure_ascii=False)

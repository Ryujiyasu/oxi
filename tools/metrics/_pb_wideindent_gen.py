# -*- coding: utf-8 -*-
"""What does Word do with a paragraph whose left + right indents exceed the column width?

reports__28abf02c8c741711: 「　　…調査平成30年７月」 has leftChars=3200 (336pt), right=4164
(208pt) and firstLineChars=2500 on a 540pt column. Word lays it at x = margin on ONE line
(Information(5) = 27.8); Oxi keeps the negative width and wraps one character per line.

Arms: left/right/firstLine combinations around the column width (A4, margins 1134 ->
content 481.9pt): sum below / equal / slightly over / far over the width, firstLine 0 /
positive / huge. Readout: x of the first character, number of lines (y walk).
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/wideindent'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
PREFIX = {'none': '', 'ideo': '　' * 26, 'ascii': ' ' * 26}[os.environ.get('PREFIX', 'none')]
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
# content width 9638 twips
ARMS = [
    ('under', 4000, 4000, 0), ('equal', 4819, 4819, 0), ('over1', 5000, 5000, 0), ('far', 5760, 4164, 0),
    ('far_fl', 5760, 4164, 2500), ('left_only_over', 10000, 0, 0), ('right_only_over', 0, 10000, 0),
    ('under_fl_over', 2000, 2000, 8000), ('far_hang', 5760, 4164, -2500),
]


def document(left, right, fl):
    ind = f'<w:ind w:left="{left}" w:right="{right}"' + (f' w:firstLine="{fl}"' if fl > 0 else (f' w:hanging="{-fl}"' if fl < 0 else '')) + '/>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:p><w:pPr>{RPR}</w:pPr><w:r>{RPR}<w:t>前の段落です。</w:t></w:r></w:p>'
            f'<w:p><w:pPr>{ind}{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">{PREFIX}調査平成30年７月</w:t></w:r></w:p>'
            f'<w:p><w:pPr>{RPR}</w:pPr><w:r>{RPR}<w:t>次の段落です。</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/></w:sectPr></w:body></w:document>')


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for name, l, r, fl in ARMS:
        at = os.path.join(OUT, f'{name}_{os.environ.get("PREFIX", "none")}.docx')
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/document.xml', document(l, r, fl))
        d = app.Documents.Open(os.path.abspath(at), ReadOnly=True)
        try:
            pr = d.Paragraphs(2).Range
            ys = []; xs = []
            for k in range(pr.Start, pr.End - 1):
                c = d.Range(k, k); y = round(c.Information(6), 2)
                if not ys or ys[-1] != y: ys.append(y); xs.append(round(c.Information(5), 2))
            nxt = d.Range(d.Paragraphs(3).Range.Start, d.Paragraphs(3).Range.Start).Information(6)
            print(f'{name:16} {os.environ.get("PREFIX","none"):5} l={l:5} r={r:5} fl={fl:5} lines={len(ys)} x0={xs[0]:.2f} first_y={ys[0]:.2f} next_y={nxt:.2f} LeftIndent={pr.ParagraphFormat.LeftIndent:.1f} Right={pr.ParagraphFormat.RightIndent:.1f} First={pr.ParagraphFormat.FirstLineIndent:.1f}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()

# -*- coding: utf-8 -*-
"""Hard-page-break x nextPage-sectPr collapse derivation.

Matrix (Calibri 11, minimal host, compat15):
  nopb    : [T1][empty sectPr][SECTB]                       control: 1 advance
  pbonly  : [T1 + trailing PB][SECTB]                       control: 1 advance
  tpb     : [T1 + trailing PB][empty sectPr][SECTB]         reports shape
  epb     : [T1][empty PB-only para][empty sectPr][SECTB]   f7115 shape
  epb2    : [T1][empty PB para][empty para][empty sectPr][SECTB]
  dbl     : [T1][PB para][PB para][empty sectPr][SECTB]
  tpbtxt  : [T1 + trailing PB][sectPr para WITH text MARKX][SECTB]
Readout: page of SECTB + total pages.
"""
import os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_sectcollapse")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

SECT1 = ('<w:pgSz w:w="11906" w:h="16838"/>'
         '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
         'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/>')

def tp(text, pb=False):
    b = '<w:r><w:br w:type="page"/></w:r>' if pb else ''
    return f'<w:p><w:r><w:t>{text}</w:t></w:r>{b}</w:p>'

def ep(pb=False):
    b = '<w:r><w:br w:type="page"/></w:r>' if pb else ''
    return f'<w:p>{b}</w:p>'

def sect_para(text=None):
    t = f'<w:r><w:t>{text}</w:t></w:r>' if text else ''
    return f'<w:p><w:pPr><w:sectPr>{SECT1}</w:sectPr></w:pPr>{t}</w:p>'

VARIANTS = {
    'nopb':   [tp('TONE'), sect_para(), tp('SECTB')],
    'pbonly': [tp('TONE', pb=True), tp('SECTB')],
    'tpb':    [tp('TONE', pb=True), sect_para(), tp('SECTB')],
    'epb':    [tp('TONE'), ep(pb=True), sect_para(), tp('SECTB')],
    'epb2':   [tp('TONE'), ep(pb=True), ep(), sect_para(), tp('SECTB')],
    'dbl':    [tp('TONE'), ep(pb=True), ep(pb=True), sect_para(), tp('SECTB')],
    'tpbtxt': [tp('TONE', pb=True), sect_para('MARKX'), tp('SECTB')],
}

def build(v, path):
    body = ''.join(VARIANTS[v]) + f'<w:sectPr>{SECT1}</w:sectPr>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as o:
        o.writestr('[Content_Types].xml', CT)
        o.writestr('_rels/.rels', RELS)
        o.writestr('word/document.xml', doc)

def main():
    os.makedirs(OUT, exist_ok=True)
    import win32com.client, fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for v in VARIANTS:
            p = os.path.join(OUT, f'sp_{v}.docx')
            build(v, p)
            pdf = p.replace('.docx', '.pdf')
            doc = word.Documents.Open(p, ReadOnly=True)
            doc.ExportAsFixedFormat(pdf, 17)
            doc.Close(False)
            d = fitz.open(pdf)
            info = {}
            for pn in range(d.page_count):
                t = d[pn].get_text()
                for tg in ('TONE', 'SECTB', 'MARKX'):
                    if tg in t and tg not in info:
                        for b in d[pn].get_text('dict')['blocks']:
                            if b['type'] != 0: continue
                            for l in b['lines']:
                                if tg in ''.join(s['text'] for s in l['spans']):
                                    info[tg] = (pn + 1, round(l['bbox'][1], 1))
            print(f'{v:7} pages={d.page_count} ' +
                  ' '.join(f'{k}=p{v2[0]}@{v2[1]}' for k, v2 in sorted(info.items())))
            d.close()
    finally:
        word.Quit()

if __name__ == '__main__':
    main()

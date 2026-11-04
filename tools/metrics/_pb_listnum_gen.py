# -*- coding: utf-8 -*-
"""Which list does a LISTNUM field number from, and at which level?

blind-J en/administrative__008d98e2 carries 10 LISTNUM fields with no cached
result (begin/instr/end only); Word shows their numbers, Oxi drew nothing.
Word's PDF: in a numId-34 paragraph at ilvl 4, «LISTNUM» / «LISTNUM \\l 6» ->
(1) (2) (3) (numId 34 ilvl 5 = decimal «(%6)»); in NON-list paragraphs (p6,
p12) «LISTNUM \\l 4» -> (A) (B) = numId 34 ilvl 3 «(%4)» upperLetter, not the
built-in NumberDefault list.

Faithful slice (the document itself): a non-list paragraph is inserted at the
very start of the body (before any list paragraph) carrying «LISTNUM \\l 4»,
«LISTNUM», «LISTNUM \\l 2» (arm 'first'), or the same paragraph inserted right
after the p10 «(ii)» paragraph (arm 'after'). Readout: the inserted paragraph's
text in Word's PDF.
"""
import glob, os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = glob.glob('pipeline_data/docx_corpus/en/administrative/008d98e2*.docx')[0]
OUT = 'tests/fixtures/listnum'; os.makedirs(OUT, exist_ok=True)


def fld(instr):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>' % instr)


PROBE = ('<w:p><w:r><w:t xml:space="preserve">PROBE START </w:t></w:r>' + fld('LISTNUM  \\l 4') +
         '<w:r><w:t xml:space="preserve"> one </w:t></w:r>' + fld('LISTNUM') +
         '<w:r><w:t xml:space="preserve"> two </w:t></w:r>' + fld('LISTNUM  \\l 2') +
         '<w:r><w:t xml:space="preserve"> three PROBE END</w:t></w:r></w:p>')


def build(arm):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, f'{arm}.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                s = data.decode('utf8')
                if arm == 'first':
                    k = s.index('<w:body>') + len('<w:body>')
                else:
                    k = s.index('</w:p>', s.index('received all cost')) + len('</w:p>')
                s = s[:k] + PROBE + s[k:]
                data = s.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    for arm in sys.argv[1:] or ['first', 'after']:
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
        t = ' '.join(P.get_text().replace('\n', ' ') for P in pymupdf.open(at[:-5] + '.pdf'))
        m = re.search(r'PROBE START.*?PROBE END', t)
        print(f'{arm}: {m.group(0) if m else None}', flush=True)

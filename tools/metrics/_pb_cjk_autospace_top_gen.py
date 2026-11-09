"""Faithful-slice probe: does an AUTO space-before (w:beforeAutospacing) survive at the top of
page 1 in a CJK document? (zh dev creative__054408b7: Normal (Web) title, Word's baseline 92.7,
Oxi 108.7 -- S1072 suppresses it only in Latin documents.)

One document per arm (the first paragraph is the page top), from 054408b7's package:
  zh_auto   : Normal (Web) title + body (as the document)        lang zh-CN, 宋体
  zh_plain  : same title with beforeAutospacing="0" before="0"   (control)
  ja_auto   : the same with eastAsia lang ja-JP and ＭＳ 明朝 runs
  ja_plain  : control
  zh_auto_p2: a page break first, the auto title opens page 2
Usage: python _pb_cjk_autospace_top_gen.py <src.docx> <out_dir>
Prints the title's baseline (Word PDF span origin) per arm.
"""
import sys, os, re, zipfile

src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
z = zipfile.ZipFile(src)
parts = {}
for it in z.infolist():
    parts[it.filename] = z.read(it.filename)
doc = parts['word/document.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + len('<w:body>')]
sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]
sty = parts['word/styles.xml'].decode('utf8')


def title(face, lang, plain):
    sp = '<w:spacing w:before="0" w:beforeAutospacing="0"/>' if plain else ''
    rp = f'<w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}"/><w:b/><w:sz w:val="28"/><w:lang w:eastAsia="{lang}"/>'
    return (f'<w:p><w:pPr><w:pStyle w:val="a4"/>{sp}<w:jc w:val="center"/></w:pPr>'
            f'<w:r><w:rPr>{rp}</w:rPr><w:t>标题文字TITLE</w:t></w:r></w:p>')


def body(face, lang):
    rp = f'<w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}"/><w:lang w:eastAsia="{lang}"/>'
    return f'<w:p><w:r><w:rPr>{rp}</w:rPr><w:t>正文第一行。</w:t></w:r></w:p>'


ARMS = {
    'zh_auto': ('宋体', 'zh-CN', False, False),
    'zh_plain': ('宋体', 'zh-CN', True, False),
    'ja_auto': ('ＭＳ 明朝', 'ja-JP', False, False),
    'ja_plain': ('ＭＳ 明朝', 'ja-JP', True, False),
    'zh_auto_p2': ('宋体', 'zh-CN', False, True),
}
import win32com.client
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
import fitz
try:
    for name, (face, lang, plain, p2) in ARMS.items():
        st = sty
        if lang == 'ja-JP':
            st = st.replace('w:eastAsia="zh-CN"', 'w:eastAsia="ja-JP"')
        pre = '<w:p><w:r><w:t>前页</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>' if p2 else ''
        xml = head + pre + title(face, lang, plain) + body(face, lang) + sect + '</w:body></w:document>'
        dst = os.path.join(out, name + '.docx')
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items():
                data = xml if fn == 'word/document.xml' else (st if fn == 'word/styles.xml' else raw)
                o.writestr(fn, data)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
        pdf = os.path.join(out, name + '.pdf')
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
        pg = fitz.open(pdf)[-1]
        ys = [l['spans'][0]['origin'][1] for b in pg.get_text('dict')['blocks'] for l in b.get('lines', [])
              if 'TITLE' in ''.join(s['text'] for s in l['spans'])]
        print(name, 'title baseline', [round(y, 2) for y in ys])
finally:
    w.Quit()

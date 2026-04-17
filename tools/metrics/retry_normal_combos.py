"""Retry crashed variants + test more combinations."""
import os, sys, time, zipfile, re
import win32com.client

TMP = os.path.abspath("pipeline_data/_norm_tmp")
os.makedirs(TMP, exist_ok=True)
SRC = os.path.abspath(r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")

z = zipfile.ZipFile(SRC)
STYLES_FULL = z.read('word/styles.xml').decode('utf-8')
DOCDEFAULTS = re.search(r'<w:docDefaults>(.*?)</w:docDefaults>', STYLES_FULL, re.DOTALL).group(0)

def build(out_path, normal):
    root = re.search(r'<w:styles\b[^>]*>', STYLES_FULL).group(0)
    new_styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{root}{DOCDEFAULTS}{normal}</w:styles>'
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == 'word/styles.xml':
                    zout.writestr(item, new_styles)
                else:
                    zout.writestr(item, zin.read(item))

def measure(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.5)
        results = {}
        for p in doc.Paragraphs:
            text = p.Range.Text
            if '（' not in text: continue
            rng = p.Range
            for ci in range(1, rng.Characters.Count + 1):
                c = rng.Characters(ci)
                if c.Text == '（':
                    try:
                        x1 = c.Information(5); y1 = c.Information(6)
                        nxt = rng.Characters(ci + 1)
                        x2 = nxt.Information(5); y2 = nxt.Information(6)
                        if abs(y1 - y2) > 2: continue
                        fs = round(c.Font.Size, 1)
                        if fs not in results:
                            results[fs] = round(x2 - x1, 2)
                    except: pass
            if len(results) >= 3: break
        doc.Close(False)
        return results
    finally:
        try:
            word.Quit()
        except: pass

BASE = '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>{pPr}{rPr}</w:style>'

TESTS = [
    ("kern_only",          BASE.format(pPr='', rPr='<w:rPr><w:kern w:val="2"/></w:rPr>')),
    ("kern_and_jc",        BASE.format(pPr='<w:pPr><w:jc w:val="both"/></w:pPr>', rPr='<w:rPr><w:kern w:val="2"/></w:rPr>')),
    ("jc_and_widow",       BASE.format(pPr='<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>', rPr='')),
    ("jc_kern_sz",         BASE.format(pPr='<w:pPr><w:jc w:val="both"/></w:pPr>', rPr='<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr>')),
    ("all_pPr",            BASE.format(pPr='<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>', rPr='')),
    ("all_rPr",            BASE.format(pPr='', rPr='<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr>')),
    ("all_pPr_and_rPr",    BASE.format(pPr='<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>', rPr='<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr>')),
    ("full_d77a_normal",   '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:rsid w:val="003B281A"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr><w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>'),
]

print(f"{'variant':<25}  {'fs=14':>7}  {'fs=12':>7}  {'fs=10.5':>8}")
print('-' * 55)
for label, xml in TESTS:
    out = os.path.join(TMP, f"retry_{label}.docx")
    build(out, xml)
    for _ in range(2):  # retry once
        try:
            r = measure(out)
            fs12 = r.get(12.0, '-')
            marker = '' if fs12 == '-' else (' **12→10.5**' if (isinstance(fs12, float) and fs12 < 11.0) else (' (11.5)' if isinstance(fs12, float) and abs(fs12-11.5)<0.1 else ' (no)'))
            print(f"{label:<25}  {r.get(14.0, '-'):>7}  {fs12:>7}  {r.get(10.5, '-'):>8}{marker}")
            break
        except Exception as e:
            time.sleep(2)
            continue
    else:
        print(f"{label:<25} ERROR after retries")

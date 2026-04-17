"""Strip d77a to keep only idx=10 paragraph, measure '（'.

If compression persists → doc-level setting (styles? fonts?)
If disappears → inter-paragraph context (header/structure)
"""
import os, sys, time, zipfile, re
import win32com.client

TMP = os.path.abspath("pipeline_data/_ctx_tmp")
os.makedirs(TMP, exist_ok=True)

SRC = os.path.abspath(r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")

def strip_to_one_para(src, out, keep_text):
    """Extract paragraphs and keep only the one containing keep_text."""
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == 'word/document.xml':
                    xml = data.decode('utf-8')
                    # Extract body
                    body_m = re.search(r'<w:body>(.*?)</w:body>', xml, re.DOTALL)
                    if body_m:
                        body = body_m.group(1)
                        # Split on paragraphs
                        paras = re.findall(r'<w:p\b[^/]*/>|<w:p\b[^>]*>(?:(?!<w:p\b).)*?</w:p>', body, re.DOTALL)
                        # Find sectPr
                        sect_m = re.search(r'<w:sectPr\b[^>]*>.*?</w:sectPr>|<w:sectPr[^/]*/>', body, re.DOTALL)
                        sect = sect_m.group(0) if sect_m else ''
                        # Keep matching para
                        kept = [p for p in paras if keep_text in p]
                        new_body = ''.join(kept[:1]) + sect  # only first match
                        xml = xml.replace(body_m.group(1), new_body)
                        data = xml.encode('utf-8')
                zout.writestr(item, data)

def measure_open(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
        for p in doc.Paragraphs:
            text = p.Range.Text
            if '（' not in text: continue
            rng = p.Range
            for ci in range(1, rng.Characters.Count + 1):
                c = rng.Characters(ci)
                if c.Text == '（':
                    x1 = c.Information(5); y1 = c.Information(6)
                    try:
                        nxt = rng.Characters(ci + 1)
                        x2 = nxt.Information(5); y2 = nxt.Information(6)
                        if abs(y1 - y2) > 2: continue
                        doc.Close(False)
                        return round(x2 - x1, 2)
                    except: pass
        doc.Close(False)
        return None
    finally:
        word.Quit()

KEEP = "公共データ利用規約（第1.0版）"  # matches idx=10

out = os.path.join(TMP, "d77a_idx10_only.docx")
strip_to_one_para(SRC, out, KEEP)
print(f"Stripped to 1 para, size = {os.path.getsize(out)} bytes")
try:
    adv = measure_open(out)
    print(f"'（' advance in stripped d77a: {adv}")
except Exception as e:
    print(f"ERROR: {e}")

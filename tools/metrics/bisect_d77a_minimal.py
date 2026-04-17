"""Strip d77a to minimal settings + characterSpacingControl only.

Check if compression STILL happens with minimal other settings,
to verify characterSpacingControl alone is sufficient.
"""
import os, sys, time, json, zipfile, re
import win32com.client

SRC = os.path.abspath(r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")

def make_variant(out, settings_body):
    """Replace w:settings content with given body (still keep xml header)."""
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == 'word/settings.xml':
                    # Replace with minimal settings
                    data = settings_body.encode('utf-8')
                zout.writestr(item, data)

# Minimal settings header (all xmlns needed)
HEADER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'''
FOOTER = '</w:settings>'

VARIANTS = [
    ("empty", f"{HEADER}{FOOTER}"),
    ("only_cSC", f'{HEADER}<w:characterSpacingControl w:val="compressPunctuation"/>{FOOTER}'),
    ("cSC_plus_useFELayout", f'{HEADER}<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/></w:compat>{FOOTER}'),
    ("cSC_plus_compat15", f'{HEADER}<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>{FOOTER}'),
    ("cSC_plus_balanceByte", f'{HEADER}<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:balanceSingleByteDoubleByteWidth/></w:compat>{FOOTER}'),
]

def measure(docx_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True); time.sleep(0.5)
        results = {}
        for pi, p in enumerate(doc.Paragraphs, 1):
            text = p.Range.Text
            if '（' not in text: continue
            for ci in range(1, p.Range.Characters.Count + 1):
                c = p.Range.Characters(ci)
                if c.Text == '（':
                    try:
                        x1 = c.Information(5); y1 = c.Information(6)
                        nxt = p.Range.Characters(ci + 1)
                        x2 = nxt.Information(5); y2 = nxt.Information(6)
                        if abs(y1-y2) > 2: continue
                        adv = round(x2 - x1, 2)
                        fs = round(c.Font.Size, 1)
                        if fs not in results:
                            results[fs] = {"advance": adv, "para_idx": pi}
                    except: pass
            if len(results) >= 3: break
        doc.Close(False)
        return results
    finally:
        word.Quit()

def main():
    tmp = os.path.abspath("pipeline_data/_bisect_tmp")
    os.makedirs(tmp, exist_ok=True)
    for name, body in VARIANTS:
        out = os.path.join(tmp, f"d77a_{name}.docx")
        make_variant(out, body)
        try:
            r = measure(out)
            print(f"{name:<25}: {r}")
        except Exception as e:
            print(f"{name:<25}: ERROR {e}")

main()

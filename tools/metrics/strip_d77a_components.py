"""Strip individual XML components from d77a to find the doc-level trigger.

Already confirmed: settings.xml features don't trigger '（' compression alone.
Test: remove styles.xml / fontTable.xml / theme.xml individually.
"""
import os, sys, time, shutil, zipfile
import win32com.client

TMP = os.path.abspath("pipeline_data/_comp_tmp")
os.makedirs(TMP, exist_ok=True)
SRC = os.path.abspath(r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")

def strip_files(out_path, files_to_exclude):
    """Copy SRC to out_path, excluding named files."""
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item in files_to_exclude: continue
                zout.writestr(item, zin.read(item))

def replace_file(out_path, file_replacements):
    """Copy SRC to out_path, replacing files with given content."""
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item in file_replacements:
                    zout.writestr(item, file_replacements[item])
                else:
                    zout.writestr(item, zin.read(item))

MINIMAL_STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
</w:styles>'''

MINIMAL_FONTTABLE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:fonts>'''

MINIMAL_THEME = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"/>'''

def measure(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
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
        word.Quit()

TESTS = [
    ("baseline", {}),  # full d77a
    ("replace_styles_minimal", {"word/styles.xml": MINIMAL_STYLES}),
    ("replace_fontTable_minimal", {"word/fontTable.xml": MINIMAL_FONTTABLE}),
    # ("replace_theme_minimal", {"word/theme/theme1.xml": MINIMAL_THEME}),  # may break
    ("replace_both_styles_font", {"word/styles.xml": MINIMAL_STYLES, "word/fontTable.xml": MINIMAL_FONTTABLE}),
]

print(f"{'variant':<30}  {'fs=14':>7}  {'fs=12':>7}  {'fs=10.5':>8}")
print('-' * 60)
for label, replacements in TESTS:
    out = os.path.join(TMP, f"{label}.docx")
    replace_file(out, replacements)
    try:
        r = measure(out)
        fs12 = r.get(12.0, '-')
        marker = '' if fs12 == '-' else (' **compressed**' if (isinstance(fs12, float) and fs12 < 11.5) else ' (no)')
        print(f"{label:<30}  {r.get(14.0, '-'):>7}  {fs12:>7}  {r.get(10.5, '-'):>8}{marker}")
    except Exception as e:
        print(f"{label:<30} ERROR {e}")

"""Binary bisect d77a styles.xml to find yakumono trigger element."""
import os, sys, time, zipfile, re
import win32com.client

TMP = os.path.abspath("pipeline_data/_styles_tmp")
os.makedirs(TMP, exist_ok=True)
SRC = os.path.abspath(r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")

z = zipfile.ZipFile(SRC)
STYLES_FULL = z.read('word/styles.xml').decode('utf-8')

# Extract docDefaults and first few styles as candidates
DOCDEFAULTS_MATCH = re.search(r'<w:docDefaults>(.*?)</w:docDefaults>', STYLES_FULL, re.DOTALL)
DOCDEFAULTS = DOCDEFAULTS_MATCH.group(0) if DOCDEFAULTS_MATCH else ''

NORMAL_MATCH = re.search(r'<w:style\s+w:type="paragraph"\s+w:default="1"[^>]*>.*?</w:style>', STYLES_FULL, re.DOTALL)
NORMAL_STYLE = NORMAL_MATCH.group(0) if NORMAL_MATCH else ''

def build_styles(extra_body):
    # The xmlns attrs must match original — use same root
    root_m = re.search(r'<w:styles\b[^>]*>', STYLES_FULL)
    root = root_m.group(0) if root_m else '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{root}{extra_body}</w:styles>'

def build_docx(out_path, new_styles):
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

MINIMAL_NORMAL = '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>'

TESTS = [
    ("empty", ""),
    ("only_normal_minimal", MINIMAL_NORMAL),
    ("only_docDefaults", DOCDEFAULTS + MINIMAL_NORMAL),  # need Normal to keep valid
    ("docDefaults_rPrDefault_only", re.sub(r'<w:pPrDefault[^/]*/>', '', DOCDEFAULTS) + MINIMAL_NORMAL),
    ("docDefaults_minus_lang", re.sub(r'<w:lang\b[^/]*/>', '', DOCDEFAULTS) + MINIMAL_NORMAL),
    ("docDefaults_minus_rFonts", re.sub(r'<w:rFonts\b[^/]*/>', '', DOCDEFAULTS) + MINIMAL_NORMAL),
    ("full_Normal_no_docDef", NORMAL_STYLE),
    ("docDefaults_plus_full_Normal", DOCDEFAULTS + NORMAL_STYLE),
]

print(f"{'variant':<35}  {'fs=14':>7}  {'fs=12':>7}  {'fs=10.5':>8}")
print('-' * 65)
for label, body in TESTS:
    try:
        xml = build_styles(body)
        out = os.path.join(TMP, f"{label}.docx")
        build_docx(out, xml)
        r = measure(out)
        fs12 = r.get(12.0, '-')
        marker = '' if fs12 == '-' else (' **compressed**' if (isinstance(fs12, float) and fs12 < 11.5) else ' (no)')
        print(f"{label:<35}  {r.get(14.0, '-'):>7}  {fs12:>7}  {r.get(10.5, '-'):>8}{marker}")
    except Exception as e:
        print(f"{label:<35} ERROR {e}")

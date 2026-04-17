"""Bisect d77a settings.xml to find yakumono compression trigger.

Produces variants of d77a with one element removed, measures '（' advance
at font size 12pt (MS Gothic).
"""
import os, sys, time, json, shutil, zipfile, re, tempfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_DOCX = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

# Features to test removing (each runs one variant)
FEATURES_TO_TEST = [
    ("baseline", None),  # unmodified
    ("no_characterSpacingControl", r'<w:characterSpacingControl[^/]*/>'),
    ("no_useFELayout", r'<w:useFELayout/>'),
    ("no_compat_mode_15", r'<w:compatSetting w:name="compatibilityMode"[^/]*/>'),
    ("no_balanceByte", r'<w:balanceSingleByteDoubleByteWidth/>'),
    ("no_overrideTableStyle", r'<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"[^/]*/>'),
    ("no_enableOpenType", r'<w:compatSetting w:name="enableOpenTypeFeatures"[^/]*/>'),
    ("no_spaceForUL", r'<w:spaceForUL/>'),
    ("no_doNotExpShiftRet", r'<w:doNotExpandShiftReturn/>'),
    ("no_drawingGrid", r'<w:drawingGridHorizontalSpacing[^/]*/>|<w:displayHorizontalDrawingGridEvery[^/]*/>|<w:displayVerticalDrawingGridEvery[^/]*/>'),
]

def create_variant(src, out, pattern):
    """Copy src to out, removing matches of pattern from settings.xml."""
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == 'word/settings.xml' and pattern:
                    text = data.decode('utf-8')
                    text = re.sub(pattern, '', text)
                    data = text.encode('utf-8')
                zout.writestr(item, data)

def measure_yakumono(docx_path):
    """Open docx, measure yakumono advances at different font sizes."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        time.sleep(0.5)
        # Find ALL '（' with (font_size, advance) — group by font_size
        results_by_fs = {}
        paras = list(doc.Paragraphs)
        for pi, p in enumerate(paras, 1):
            text = p.Range.Text
            if '（' not in text: continue
            for ci in range(1, p.Range.Characters.Count + 1):
                c = p.Range.Characters(ci)
                if c.Text == '（':
                    x1 = c.Information(5)
                    try:
                        nxt = p.Range.Characters(ci + 1)
                        x2 = nxt.Information(5)
                        y1 = c.Information(6)
                        y2 = nxt.Information(6)
                        if abs(y1 - y2) > 2: continue  # line wrap
                        advance = x2 - x1
                        fs = round(c.Font.Size, 1)
                        if fs not in results_by_fs:
                            results_by_fs[fs] = {"advance": round(advance, 2), "para_idx": pi, "family": c.Font.Name}
                    except:
                        pass
            if len(results_by_fs) >= 3: break  # enough samples
        doc.Close(False)
        return results_by_fs
    finally:
        word.Quit()

def main():
    results = []
    tmp = os.path.abspath("pipeline_data/_bisect_tmp")
    os.makedirs(tmp, exist_ok=True)
    for name, pattern in FEATURES_TO_TEST:
        out_path = os.path.join(tmp, f"d77a_{name}.docx")
        create_variant(SRC_DOCX, out_path, pattern)
        print(f"Testing {name}...", flush=True)
        r = measure_yakumono(out_path)
        r = {"variant": name, **r}
        results.append(r)
        print(f"  → {r}", flush=True)

    out_file = "pipeline_data/d77a_yakumono_bisect.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {out_file}")

    print("\n=== Summary ===")
    # Show advance per font_size for each variant
    all_fs = set()
    for r in results:
        if isinstance(r, dict):
            for k, v in r.items():
                if isinstance(v, dict) and 'advance' in v:
                    all_fs.add(k)
    all_fs = sorted(all_fs)
    header = f"  {'variant':<30}"
    for fs in all_fs:
        header += f" fs={fs}"
    print(header)
    for r in results:
        line = f"  {r.get('variant'):<30}"
        for fs in all_fs:
            if fs in r:
                line += f" {r[fs]['advance']:>6}"
            else:
                line += f" {'-':>6}"
        print(line)

if __name__ == "__main__":
    main()

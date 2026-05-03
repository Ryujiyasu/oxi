"""For each variant, render with Word + Oxi, COM-measure how many chars fit on line 1."""
import time, sys, os, subprocess, json
import win32com.client, pythoncom
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

WD_VPOS = 6
WD_HPOS = 5

DOCX_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
RENDERER = "c:/Users/ryuji/oxi-main/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"

VARIANTS = [
    "v0_baseline", "v1_tcMar_540", "v2_no_adjustR", "v3_no_autoSpace",
    "v4_default_kern", "v5_no_spacing_neg", "v6_no_hanging",
]

# COM measure each variant
print("=== Word COM measurements (chars on line 1 of P1) ===")
print(f"{'variant':22s} {'Word_n_line1':>15s} {'last_x_line1':>14s}")

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
word_results = {}
try:
    for vname in VARIANTS:
        path = os.path.join(DOCX_DIR, f"repro_v_{vname}.docx")
        if not os.path.exists(path):
            print(f"  {vname:22s}  MISSING")
            continue
        try:
            doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
            try:
                doc.Repaginate()
                time.sleep(0.2)
                # Get first paragraph in first cell
                t = doc.Tables(1)
                c1 = t.Range.Cells.Item(1)
                p1 = c1.Range.Paragraphs(1)
                pr = p1.Range
                full_text = pr.Text.replace('\r','').replace('\x07','')
                n = len(full_text)
                # Find first char that's on line 2 (y > line 1's y)
                first_y = float(doc.Range(pr.Start, pr.Start+1).Information(WD_VPOS))
                n_line1 = 0
                last_x_line1 = -1
                for ci in range(n):
                    crng = doc.Range(pr.Start + ci, pr.Start + ci + 1)
                    y = float(crng.Information(WD_VPOS))
                    if abs(y - first_y) < 1.0:
                        n_line1 = ci + 1
                        last_x_line1 = float(crng.Information(WD_HPOS))
                    else:
                        break
                print(f"  {vname:22s}  {n_line1:>15d} {last_x_line1:>14.2f}  total_chars={n}")
                word_results[vname] = (n_line1, last_x_line1)
            finally:
                doc.Close(False)
        except Exception as ex:
            print(f"  {vname:22s}  ERROR: {str(ex)[:60]}")
finally:
    word.Quit()
    pythoncom.CoUninitialize()

# Render with Oxi and find how many chars on line 1
print("\n=== Oxi measurements (last x on line 1) ===")
print(f"{'variant':22s} {'Oxi line1 last x':>20s}")
oxi_results = {}
for vname in VARIANTS:
    path = os.path.join(DOCX_DIR, f"repro_v_{vname}.docx")
    if not os.path.exists(path): continue
    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/repro_v_{vname}.json"
    proc = subprocess.run(
        [RENDERER, path, "/tmp/dummy.png",
         f"--dump-layout={layout_out}"],
        capture_output=True, timeout=60,
    )
    if not os.path.exists(layout_out):
        print(f"  {vname:22s}  RENDER FAIL")
        continue
    with open(layout_out, encoding='utf-8') as f:
        layout = json.load(f)
    page1 = layout['pages'][0]
    # Find smallest y > 50 (skip page top), then group by y
    text_elems = [e for e in page1['elements'] if e.get('type') == 'text']
    if not text_elems: continue
    text_elems.sort(key=lambda e: e['y'])
    first_y = text_elems[0]['y']
    line1_elems = [e for e in text_elems if abs(e['y'] - first_y) < 1.0]
    if not line1_elems: continue
    # Find max (x + w) on line 1
    end_xs = [e['x'] + e.get('w', 0) for e in line1_elems]
    last_x = max(end_xs)
    # Sum chars
    chars_on_line1 = sum(len(e.get('text','')) for e in line1_elems)
    print(f"  {vname:22s}  {last_x:>20.2f}  chars={chars_on_line1}")
    oxi_results[vname] = (chars_on_line1, last_x)

# Diff summary
print("\n=== Word vs Oxi diff (line 1 char count) ===")
print(f"{'variant':22s} {'Word_n':>8s} {'Oxi_n':>8s} {'Δ (Word-Oxi)':>15s}")
for vname in VARIANTS:
    if vname in word_results and vname in oxi_results:
        w_n = word_results[vname][0]
        o_n = oxi_results[vname][0]
        d = w_n - o_n
        marker = ' <<<' if d != 0 else ''
        print(f"  {vname:22s}  {w_n:>8d} {o_n:>8d}  {d:>+15d}{marker}")

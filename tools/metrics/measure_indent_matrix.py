"""For each indent variant, measure Word + Oxi first char x of paragraph."""
import time, sys, os, subprocess, json
import win32com.client, pythoncom
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
WD_VPOS = 6; WD_HPOS = 5
DOCX_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
RENDERER = "c:/Users/ryuji/oxi-main/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"

VARIANTS = ["v0_full", "v1_no_leftChars", "v2_no_left", "v3_small_left",
            "v4_no_hanging", "v5_chars_hanging_only", "v6_twip_hanging_only",
            "v7_huge_left", "v8_zero_indent"]

# Word measurement
print(f"=== Word COM measurement ===")
print(f"  {'variant':24s} {'first_x':>9} {'L_indent':>10} {'fl_indent':>11} {'first_pos':>10}")
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False; word.DisplayAlerts = 0
word_data = {}
try:
    for v in VARIANTS:
        path = os.path.join(DOCX_DIR, f"indent_matrix_{v}.docx")
        doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            doc.Repaginate(); time.sleep(0.2)
            t = doc.Tables(1)
            c1 = t.Range.Cells.Item(1)
            p1 = c1.Range.Paragraphs(1)
            pr = p1.Range
            cs = doc.Range(pr.Start, pr.Start + 1)
            x = float(cs.Information(WD_HPOS))
            pf = p1.Format
            li = pf.LeftIndent
            fl = pf.FirstLineIndent
            fp = li + fl  # First-line position
            print(f"  {v:24s} {x:>9.2f} {li:>10.2f} {fl:>11.2f} {fp:>10.2f}")
            word_data[v] = (x, li, fl, fp)
        finally:
            doc.Close(False)
finally:
    word.Quit()
    pythoncom.CoUninitialize()

# Oxi measurement
print(f"\n=== Oxi measurement ===")
print(f"  {'variant':24s} {'first_x':>9}")
oxi_data = {}
for v in VARIANTS:
    path = os.path.join(DOCX_DIR, f"indent_matrix_{v}.docx")
    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/imatrix_{v}.json"
    proc = subprocess.run([RENDERER, path, "/tmp/dummy.png", f"--dump-layout={layout_out}"], capture_output=True, timeout=60)
    if not os.path.exists(layout_out): continue
    with open(layout_out, encoding='utf-8') as f:
        layout = json.load(f)
    text_elems = sorted([e for e in layout['pages'][0]['elements'] if e.get('type')=='text'], key=lambda e: (e['y'], e['x']))
    if not text_elems: continue
    first_y = text_elems[0]['y']
    line1 = [e for e in text_elems if abs(e['y'] - first_y) < 1.0]
    first_x = min(e['x'] for e in line1)
    print(f"  {v:24s} {first_x:>9.2f}")
    oxi_data[v] = first_x

# Combined
print(f"\n=== Combined ===")
print(f"  {'variant':24s} {'Word':>8} {'Oxi':>8} {'Δ':>7}  {'note'}")
for v in VARIANTS:
    if v not in word_data or v not in oxi_data: continue
    w_x = word_data[v][0]; o_x = oxi_data[v]
    d = o_x - w_x
    li = word_data[v][1]; fl = word_data[v][2]; fp = word_data[v][3]
    mark = " <<<" if abs(d) > 2 else ""
    print(f"  {v:24s} {w_x:>8.2f} {o_x:>8.2f} {d:>+7.2f}  L={li:.2f} fl={fl:.2f} pos={fp:.2f}{mark}")

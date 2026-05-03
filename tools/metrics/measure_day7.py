"""Measure Word + Oxi line 1 + line 2 first char x for Day 7 variants."""
import time, sys, os, subprocess, json
import win32com.client, pythoncom
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
WD_VPOS = 6; WD_HPOS = 5
DOCX_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
RENDERER = "c:/Users/ryuji/oxi-main/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"

VARIANTS = ["d7_v0_hang_only", "d7_v1_hang_offset", "d7_v2_no_hang_left30",
            "d7_v3_huge_hang_like_1636", "d7_v4_bullet", "d7_v5_no_indent_bullet"]

print(f"=== Word: line 1 + line 2 first x ===")
print(f"  {'variant':28s} {'l1_x':>7} {'n_l1':>5} {'l2_x':>7} {'L_indent':>9} {'fl':>7} {'fp':>7}")
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False; word.DisplayAlerts = 0
word_data = {}
try:
    for v in VARIANTS:
        path = os.path.join(DOCX_DIR, f"{v}.docx")
        doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            doc.Repaginate(); time.sleep(0.2)
            t = doc.Tables(1); c1 = t.Range.Cells.Item(1)
            p1 = c1.Range.Paragraphs(1); pr = p1.Range
            txt = pr.Text.replace('\r','').replace('\x07','')
            n = len(txt)
            pf = p1.Format
            li = pf.LeftIndent; fl = pf.FirstLineIndent; fp = li + fl
            l1_y = float(doc.Range(pr.Start, pr.Start+1).Information(WD_VPOS))
            l1_x = float(doc.Range(pr.Start, pr.Start+1).Information(WD_HPOS))
            n_l1 = 0; l2_x = None
            for ci in range(n):
                cs = doc.Range(pr.Start + ci, pr.Start + ci + 1)
                cy = float(cs.Information(WD_VPOS))
                if abs(cy - l1_y) < 1.0:
                    n_l1 = ci + 1
                else:
                    l2_x = float(cs.Information(WD_HPOS))
                    break
            l2s = f"{l2_x:>7.2f}" if l2_x else "  no_wrap"
            print(f"  {v:28s} {l1_x:>7.2f} {n_l1:>5} {l2s} {li:>9.2f} {fl:>7.2f} {fp:>7.2f}")
            word_data[v] = (l1_x, n_l1, l2_x, li, fl, fp)
        finally:
            doc.Close(False)
finally:
    word.Quit()
    pythoncom.CoUninitialize()

print(f"\n=== Oxi: line 1 + line 2 first x ===")
print(f"  {'variant':28s} {'l1_x':>7} {'n_l1':>5} {'l2_x':>7}")
oxi_data = {}
for v in VARIANTS:
    path = os.path.join(DOCX_DIR, f"{v}.docx")
    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/{v}.json"
    proc = subprocess.run([RENDERER, path, "/tmp/dummy.png", f"--dump-layout={layout_out}"], capture_output=True, timeout=60)
    if not os.path.exists(layout_out): continue
    with open(layout_out, encoding='utf-8') as f:
        layout = json.load(f)
    text_elems = sorted([e for e in layout['pages'][0]['elements'] if e.get('type')=='text'], key=lambda e: (e['y'], e['x']))
    if not text_elems: continue
    # Group by Y
    lines = []
    cur_y = None; cur = []
    for e in text_elems:
        if cur_y is None or abs(e['y'] - cur_y) < 1.0:
            cur.append(e); cur_y = e['y']
        else:
            lines.append({"y": cur_y, "elems": cur})
            cur = [e]; cur_y = e['y']
    if cur: lines.append({"y": cur_y, "elems": cur})
    l1 = lines[0] if lines else None
    l2 = lines[1] if len(lines) > 1 else None
    l1_x = min(e['x'] for e in l1['elems']) if l1 else 0
    l1_n = sum(len(e.get('text','')) for e in l1['elems']) if l1 else 0
    l2_x = min(e['x'] for e in l2['elems']) if l2 else None
    l2s = f"{l2_x:>7.2f}" if l2_x else "  no_wrap"
    print(f"  {v:28s} {l1_x:>7.2f} {l1_n:>5} {l2s}")
    oxi_data[v] = (l1_x, l1_n, l2_x)

print(f"\n=== Combined ===")
print(f"  {'variant':28s} {'W_l1':>6} {'O_l1':>6} {'Δl1':>6} | {'W_l2':>6} {'O_l2':>6} {'Δl2':>6}")
for v in VARIANTS:
    if v not in word_data or v not in oxi_data: continue
    w_l1 = word_data[v][0]; o_l1 = oxi_data[v][0]
    w_l2 = word_data[v][2]; o_l2 = oxi_data[v][2]
    dl1 = o_l1 - w_l1
    dl2_str = f"{o_l2 - w_l2:>+6.2f}" if (w_l2 and o_l2) else "  N/A"
    w_l2_str = f"{w_l2:>6.2f}" if w_l2 else "  N/A"
    o_l2_str = f"{o_l2:>6.2f}" if o_l2 else "  N/A"
    print(f"  {v:28s} {w_l1:>6.2f} {o_l1:>6.2f} {dl1:>+6.2f} | {w_l2_str} {o_l2_str} {dl2_str}")

"""For each tcMar variant, COM-measure chars on line 1 + Word-reported cell padding."""
import time, sys, os, subprocess, json
import win32com.client, pythoncom
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

WD_VPOS = 6
WD_HPOS = 5

DOCX_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
RENDERER = "c:/Users/ryuji/oxi-main/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"

SWEEP = [None, 0, 30, 50, 80, 99, 108, 150, 200, 300, 500]
NAMES = ["none" if tw is None else f"{tw:04d}" for tw in SWEEP]

# COM measure
print(f"=== Word COM ===")
print(f"  {'tcMar':10s} {'word_n':>8s} {'last_x':>8s} {'L_pad':>7s} {'R_pad':>7s} {'cell_w':>8s}")

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
word_n = {}
try:
    for tw, name in zip(SWEEP, NAMES):
        path = os.path.join(DOCX_DIR, f"repro_tcmar_{name}.docx")
        if not os.path.exists(path): continue
        try:
            doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
            try:
                doc.Repaginate()
                time.sleep(0.2)
                t = doc.Tables(1)
                c1 = t.Range.Cells.Item(1)
                # Get cell properties
                cell_w = c1.Width
                lpad = c1.LeftPadding
                rpad = c1.RightPadding
                # Measure chars on line 1
                p1 = c1.Range.Paragraphs(1)
                pr = p1.Range
                full_text = pr.Text.replace('\r','').replace('\x07','')
                n = len(full_text)
                first_y = float(doc.Range(pr.Start, pr.Start+1).Information(WD_VPOS))
                n_l1 = 0; lx = 0
                for ci in range(n):
                    crng = doc.Range(pr.Start + ci, pr.Start + ci + 1)
                    y = float(crng.Information(WD_VPOS))
                    if abs(y - first_y) < 1.0:
                        n_l1 = ci + 1
                        lx = float(crng.Information(WD_HPOS))
                    else:
                        break
                tw_str = "(none)" if tw is None else str(tw)
                print(f"  {tw_str:10s} {n_l1:>8d} {lx:>8.2f} {lpad:>7.2f} {rpad:>7.2f} {cell_w:>8.2f}")
                word_n[name] = (n_l1, lx, lpad, rpad, cell_w)
            finally:
                doc.Close(False)
        except Exception as ex:
            print(f"  {name:10s}  ERROR: {str(ex)[:60]}")
finally:
    word.Quit()
    pythoncom.CoUninitialize()

# Oxi measurement
print(f"\n=== Oxi (chars on line 1) ===")
print(f"  {'tcMar':10s} {'oxi_n':>8s} {'last_x':>8s}")
oxi_n = {}
for tw, name in zip(SWEEP, NAMES):
    path = os.path.join(DOCX_DIR, f"repro_tcmar_{name}.docx")
    if not os.path.exists(path): continue
    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/repro_tcmar_{name}.json"
    proc = subprocess.run(
        [RENDERER, path, "/tmp/dummy.png", f"--dump-layout={layout_out}"],
        capture_output=True, timeout=60,
    )
    if not os.path.exists(layout_out): continue
    with open(layout_out, encoding='utf-8') as f:
        layout = json.load(f)
    page1 = layout['pages'][0]
    text_elems = sorted([e for e in page1['elements'] if e.get('type')=='text'], key=lambda e: e['y'])
    if not text_elems: continue
    first_y = text_elems[0]['y']
    line1 = [e for e in text_elems if abs(e['y'] - first_y) < 1.0]
    n_chars = sum(len(e.get('text','')) for e in line1)
    last_x = max(e['x'] + e.get('w', 0) for e in line1)
    tw_str = "(none)" if tw is None else str(tw)
    print(f"  {tw_str:10s} {n_chars:>8d} {last_x:>8.2f}")
    oxi_n[name] = (n_chars, last_x)

# Combined
print(f"\n=== Combined (chars on line 1) ===")
print(f"  {'tcMar_tw':10s} {'Word':>6s} {'Oxi':>6s} {'Δ':>6s}  {'Word_lpad':>10s} {'Word_cellw':>11s}")
for tw, name in zip(SWEEP, NAMES):
    if name not in word_n or name not in oxi_n: continue
    w = word_n[name][0]; o = oxi_n[name][0]
    lpad = word_n[name][2]; cw = word_n[name][4]
    tw_str = "(none)" if tw is None else str(tw)
    marker = ' <<<' if (w - o) != 0 else ''
    print(f"  {tw_str:10s} {w:>6d} {o:>6d} {w-o:>+6d}  {lpad:>10.2f} {cw:>11.2f}{marker}")

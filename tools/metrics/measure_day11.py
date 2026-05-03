"""Measure per-char advance for cs/grid variants."""
import time, sys, os, subprocess, json
import win32com.client, pythoncom
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
WD_VPOS = 6; WD_HPOS = 5
DOCX_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
RENDERER = "c:/Users/ryuji/oxi-main/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"

VARIANTS = ["d11_v0_baseline", "d11_v1_cs_neg9", "d11_v2_cs_neg20",
            "d11_v3_grid_only", "d11_v4_grid_cs_neg9", "d11_v5_grid_cs_neg1"]

def measure_word(docx):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(docx, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            doc.Repaginate(); time.sleep(0.2)
            p1 = doc.Paragraphs(1); pr = p1.Range
            txt = pr.Text.replace('\r','').replace('\x07','')
            xs = []
            first_y = float(doc.Range(pr.Start, pr.Start+1).Information(WD_VPOS))
            for ci in range(len(txt)):
                cs = doc.Range(pr.Start + ci, pr.Start + ci + 1)
                cy = float(cs.Information(WD_VPOS))
                if abs(cy - first_y) > 1: break
                cx = float(cs.Information(WD_HPOS))
                xs.append(cx)
            doc.Close(False)
            return xs
        except Exception as ex:
            doc.Close(False); return None
    finally:
        word.Quit(); pythoncom.CoUninitialize()

def measure_oxi(docx, layout_out):
    proc = subprocess.run([RENDERER, docx, "/tmp/dummy.png", f"--dump-layout={layout_out}"], capture_output=True, timeout=60)
    if not os.path.exists(layout_out): return None
    with open(layout_out, encoding='utf-8') as f:
        layout = json.load(f)
    text_elems = sorted([e for e in layout['pages'][0]['elements'] if e.get('type')=='text'], key=lambda e: (e['y'], e['x']))
    if not text_elems: return None
    first_y = text_elems[0]['y']
    line1 = [e for e in text_elems if abs(e['y']-first_y)<1]
    chars = []
    for e in sorted(line1, key=lambda e: e['x']):
        n = len(e.get('text',''))
        if n == 0: continue
        w = e.get('w', 0)
        for i, ch in enumerate(e['text']):
            chars.append(e['x'] + (i/n)*w if n > 0 else e['x'])
    return chars

print(f"{'variant':22s} {'W_avg':>7} {'O_avg':>7} {'Δ':>7}")
all_data = {}
for v in VARIANTS:
    docx = os.path.join(DOCX_DIR, f"{v}.docx")
    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/{v}_oxi.json"
    word_xs = measure_word(docx)
    oxi_xs = measure_oxi(docx, layout_out)
    if not word_xs or not oxi_xs:
        print(f"  {v:22s} measurement failed"); continue
    n = min(len(word_xs), len(oxi_xs), 19)  # 20 chars text → 19 advances
    if n < 2: continue
    w_avg = (word_xs[n-1] - word_xs[0]) / (n-1)
    o_avg = (oxi_xs[n-1] - oxi_xs[0]) / (n-1)
    d = o_avg - w_avg
    print(f"{v:22s} {w_avg:>7.3f} {o_avg:>7.3f} {d:>+7.3f}")
    all_data[v] = (w_avg, o_avg, d)

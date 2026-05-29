"""S445: measure Word row pitch for each trHeight repro variant + Oxi side."""
import sys, os, glob, json, subprocess, tempfile
import win32com.client as win32
sys.stdout.reconfigure(encoding="utf-8")

REPRO = r"c:\Users\ryuji\oxi-main\tools\metrics\s445_trheight_repro"
RENDERER = r"c:\Users\ryuji\oxi-main\tools\oxi-gdi-renderer\target\release\oxi-gdi-renderer.exe"

def word_pitches(doc):
    tbl = doc.Tables(1)
    cells = tbl.Range.Cells
    rowy = {}
    for ci in range(1, cells.Count + 1):
        c = cells(ci)
        ri = c.RowIndex
        sr = doc.Range(c.Range.Start, c.Range.Start)
        rowy.setdefault(ri, float(sr.Information(6)))
    ys = [rowy[k] for k in sorted(rowy)]
    return ys, [round(b - a, 3) for a, b in zip(ys, ys[1:])]

def oxi_pitches(path):
    with tempfile.TemporaryDirectory() as tmp:
        dump = os.path.join(tmp, "l.json")
        r = subprocess.run([RENDERER, path, os.path.join(tmp, "p_"), "--dump-layout=" + dump],
                           capture_output=True, text=True, timeout=120)
        if r.returncode != 0:
            return None, f"rc={r.returncode} {r.stderr[:200]}"
        d = json.load(open(dump, encoding="utf-8"))
    texts = [e for e in d["pages"][0]["elements"] if e.get("type") == "text" and e.get("text", "").strip()]
    ys = sorted(e["y"] for e in texts)
    return ys, [round(b - a, 3) for a, b in zip(ys, ys[1:])]

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
try:
    for f in sorted(glob.glob(os.path.join(REPRO, "*.docx"))):
        name = os.path.basename(f)
        try:
            doc = word.Documents.Open(f, ReadOnly=True)
            wy, wp = word_pitches(doc)
            doc.Close(False)
        except Exception as e:
            wy, wp = None, f"WORD_ERR {e}"
        oy, op = oxi_pitches(f)
        print(f"\n=== {name} ===")
        print(f"  Word pitches: {wp}")
        print(f"  Oxi  pitches: {op}")
finally:
    word.Quit()

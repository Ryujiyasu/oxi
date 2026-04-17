import os, sys, time, subprocess
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VARIANTS = ["R1_baseline", "R2_kern", "R3_jc_both", "R4_widow", "R5_lang",
            "R6_docdefault_rfonts", "R7_szCs", "R_ALL"]
DOC_DIR = os.path.abspath(r"pipeline_data")


def measure_one(v):
    # Kill any stale Word processes first
    try:
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, timeout=5)
    except Exception: pass
    time.sleep(0.5)
    word = win32com.client.DispatchEx("Word.Application")
    try:
        try: word.Visible = False
        except: pass
        try: word.DisplayAlerts = False
        except: pass
        path = os.path.join(DOC_DIR, f"d77a_{v}.docx")
        if not os.path.exists(path):
            return None
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
        doc.Repaginate()
        result = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            txt = p.Range.Text.replace("\r","").replace("\x07","")
            if "・利用規約名" in txt[:10]:
                pr = p.Range
                c1 = pr.Characters(1); c2 = pr.Characters(2)
                x1 = c1.Information(5); x2 = c2.Information(5)
                result = x2 - x1
                break
        doc.Close(False)
        return result
    finally:
        try: word.Quit()
        except: pass


for v in VARIANTS:
    try:
        adv = measure_one(v)
        if adv is not None:
            status = "COMPRESSED" if adv < 11 else "natural"
            print(f"[{v:>22}] adv={adv:.2f}pt  {status}", flush=True)
        else:
            print(f"[{v:>22}] no data", flush=True)
    except Exception as e:
        print(f"[{v:>22}] err: {e}", flush=True)
